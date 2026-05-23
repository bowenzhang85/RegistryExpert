using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RegistryExpert.Core.Services
{
    /// <summary>
    /// Minimal BCD (Boot Configuration Data) hive reader. Parses just enough of the
    /// BCD store to identify which partition holds the boot loader (System) and the
    /// Windows OS installation (Boot). Used by <see cref="DiskLayoutBcdEnricher"/>
    /// to authoritatively tag the corresponding partitions in a DiskLayoutModel.
    /// </summary>
    /// <remarks>
    /// Format references are based on Microsoft's public BCD documentation and the
    /// reverse-engineered DeviceElement binary layout observed in the field. We
    /// intentionally parse a narrow subset (device + osdevice + path + systemroot
    /// for Boot Loader and Resume Loader entries) and follow inherit chains
    /// (Elements\14000006) up to a small depth to find non-zero descriptors.
    /// </remarks>
    public class BcdParser : IDisposable
    {
        private readonly OfflineRegistryParser _parser;

        // BCD Description Type constants (from bcd.h conceptual values)
        public const uint TypeBootManager   = 0x10100002;
        public const uint TypeBootLoader    = 0x10200003;
        public const uint TypeResumeLoader  = 0x10200004;

        // Well-known BCD element IDs
        public const string ElementDevice     = "11000001"; // device — boot loader's system partition
        public const string ElementPath       = "12000002"; // path  — loader application path
        public const string ElementInherit    = "14000006"; // inherit chain
        public const string ElementOsDevice   = "21000001"; // osdevice — partition where Windows is installed
        public const string ElementSystemRoot = "22000002"; // systemroot — \Windows

        // Maximum depth to follow inherit chains (avoid cycles in malformed stores)
        private const int MaxInheritDepth = 8;

        public BcdParser()
        {
            _parser = new OfflineRegistryParser();
        }

        /// <summary>Load the BCD hive from disk.</summary>
        public bool Load(string filePath) => _parser.LoadHive(filePath);

        public string? FilePath => _parser.FilePath;

        public void Dispose() => _parser.Dispose();

        /// <summary>
        /// Enumerate all interesting BCD boot entries — Boot Loader + Resume Loader
        /// objects with their resolved device descriptors.
        /// </summary>
        public List<BcdBootEntry> EnumerateBootEntries()
        {
            var result = new List<BcdBootEntry>();
            var objectsKey = _parser.GetKey("Objects");
            if (objectsKey == null) return result;

            // Build a lookup of objectId → object key for inherit resolution
            var objectsById = new Dictionary<string, RegistryParser.Abstractions.RegistryKey>(StringComparer.OrdinalIgnoreCase);
            foreach (var obj in objectsKey.SubKeys)
            {
                objectsById[obj.KeyName.Trim('{', '}').ToLowerInvariant()] = obj;
                objectsById[obj.KeyName] = obj;
            }

            foreach (var obj in objectsKey.SubKeys)
            {
                var descKey = obj.SubKeys.FirstOrDefault(k =>
                    string.Equals(k.KeyName, "Description", StringComparison.OrdinalIgnoreCase));
                if (descKey == null) continue;

                uint typeVal = ReadDword(descKey, "Type") ?? 0;
                if (typeVal != TypeBootLoader && typeVal != TypeResumeLoader) continue;

                var entry = new BcdBootEntry
                {
                    ObjectId = obj.KeyName,
                    Type = typeVal,
                    TypeName = typeVal switch
                    {
                        TypeBootLoader   => "Windows Boot Loader",
                        TypeResumeLoader => "Windows Resume Loader",
                        _                => $"Unknown (0x{typeVal:X8})"
                    },
                };

                var elementsKey = obj.SubKeys.FirstOrDefault(k =>
                    string.Equals(k.KeyName, "Elements", StringComparison.OrdinalIgnoreCase));
                if (elementsKey == null) { result.Add(entry); continue; }

                // device (11000001), osdevice (21000001) — read inline or via inherit chain
                entry.Device = ReadDeviceDescriptor(obj, ElementDevice, objectsById, 0);
                entry.OsDevice = ReadDeviceDescriptor(obj, ElementOsDevice, objectsById, 0);
                entry.Path = ReadString(elementsKey, ElementPath);
                entry.SystemRoot = ReadString(elementsKey, ElementSystemRoot);

                result.Add(entry);
            }

            return result;
        }

        /// <summary>
        /// Read a device descriptor element. If the descriptor is empty (all zeros
        /// in the partition fields), follow the inherit chain (14000006) up to
        /// MaxInheritDepth levels looking for a non-empty descriptor.
        /// </summary>
        private BcdDeviceDescriptor? ReadDeviceDescriptor(
            RegistryParser.Abstractions.RegistryKey objKey,
            string elementId,
            Dictionary<string, RegistryParser.Abstractions.RegistryKey> objectsById,
            int depth)
        {
            if (depth >= MaxInheritDepth) return null;

            var elementsKey = objKey.SubKeys.FirstOrDefault(k =>
                string.Equals(k.KeyName, "Elements", StringComparison.OrdinalIgnoreCase));
            if (elementsKey == null) return null;

            var elemKey = elementsKey.SubKeys.FirstOrDefault(k =>
                string.Equals(k.KeyName, elementId, StringComparison.OrdinalIgnoreCase));
            if (elemKey != null)
            {
                var bytes = ReadElementBytes(elemKey);
                if (bytes != null)
                {
                    var desc = ParseDeviceDescriptor(bytes);
                    if (desc != null && desc.HasUsableTarget)
                        return desc;
                }
            }

            // Follow inherit chain
            var inheritElem = elementsKey.SubKeys.FirstOrDefault(k =>
                string.Equals(k.KeyName, ElementInherit, StringComparison.OrdinalIgnoreCase));
            if (inheritElem == null) return null;

            var inheritList = ReadStringElement(inheritElem) ?? "";
            foreach (var refId in inheritList
                .Split(new[] { ' ', '\t', '\r', '\n', '\0' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = refId.Trim();
                if (objectsById.TryGetValue(trimmed, out var refObj) ||
                    objectsById.TryGetValue(trimmed.Trim('{', '}').ToLowerInvariant(), out refObj))
                {
                    var desc = ReadDeviceDescriptor(refObj, elementId, objectsById, depth + 1);
                    if (desc != null && desc.HasUsableTarget)
                        return desc;
                }
            }

            return null;
        }

        /// <summary>
        /// Parse a BCD DeviceElement binary blob (typically 88+ bytes). Recognises
        /// the common type 6 (PartitionEx) MBR layout and type 8 GPT layout.
        /// </summary>
        /// <remarks>
        /// Observed layout for type 6 MBR:
        ///   bytes  0–15: zero header
        ///   bytes 16–19: DeviceType (DWORD LE) = 6
        ///   bytes 20–23: zero
        ///   bytes 24–31: Size (QWORD LE) = 0x48
        ///   bytes 32–47: reserved / padding (typically zeros with a small struct hint at byte 34)
        ///   bytes 48–55: PartitionNumber (QWORD LE) — for MBR
        ///   bytes 56–59: MBR disk signature (DWORD LE)
        ///   bytes 60–63: padding
        /// For GPT, partition + disk GUIDs would occupy a similar region (offset 32 area).
        /// We prefer MBR detection when a valid signature is present, and treat GPT
        /// fields only when no MBR signature is detected.
        /// </remarks>
        private static BcdDeviceDescriptor? ParseDeviceDescriptor(byte[] bytes)
        {
            if (bytes.Length < 32) return null;

            // Read DeviceType at offset 16
            uint deviceType = BitConverter.ToUInt32(bytes, 16);
            var desc = new BcdDeviceDescriptor { DeviceType = deviceType };

            if (deviceType == 5 || deviceType == 6)
            {
                // First check for MBR signature at offset 56. Non-zero → MBR.
                if (bytes.Length >= 60)
                {
                    uint sig = BitConverter.ToUInt32(bytes, 56);
                    if (sig != 0)
                    {
                        desc.MbrDiskSignature = sig;
                        desc.PartitionStyle = "MBR";
                        return desc;
                    }
                }

                // Otherwise, try GPT: partition GUID at offset 32 + disk GUID at offset 48.
                // Only accept GUIDs that look "real" — a valid Microsoft type GUID has
                // a non-zero variant field. Reject obvious garbage like
                // {00100000-0000-0000-0000-000000000000}.
                if (bytes.Length >= 48)
                {
                    var partGuidBytes = new byte[16];
                    Array.Copy(bytes, 32, partGuidBytes, 0, 16);
                    if (LooksLikeValidGuid(partGuidBytes))
                    {
                        try
                        {
                            var partGuid = new Guid(partGuidBytes);
                            desc.GptPartitionGuid = partGuid.ToString("B");
                            desc.PartitionStyle = "GPT";
                        }
                        catch { }
                    }
                }

                if (bytes.Length >= 64)
                {
                    var diskGuidBytes = new byte[16];
                    Array.Copy(bytes, 48, diskGuidBytes, 0, 16);
                    if (LooksLikeValidGuid(diskGuidBytes))
                    {
                        try
                        {
                            var diskGuid = new Guid(diskGuidBytes);
                            desc.GptDiskGuid = diskGuid.ToString("B");
                            if (string.IsNullOrEmpty(desc.PartitionStyle))
                                desc.PartitionStyle = "GPT";
                        }
                        catch { }
                    }
                }
            }

            return desc;
        }

        /// <summary>
        /// Heuristic to reject obvious garbage GUIDs. A real GPT partition or disk
        /// GUID has at least 8 distinct non-zero bytes (real GUIDs are essentially
        /// random). Sequences with mostly zero bytes are typically misread fields.
        /// </summary>
        private static bool LooksLikeValidGuid(byte[] bytes)
        {
            if (bytes.Length != 16) return false;
            int nonZero = 0;
            for (int i = 0; i < 16; i++)
                if (bytes[i] != 0) nonZero++;
            return nonZero >= 8;
        }

        private static byte[]? ReadElementBytes(RegistryParser.Abstractions.RegistryKey elemKey)
        {
            var val = elemKey.Values.FirstOrDefault(v =>
                string.Equals(v.ValueName, "Element", StringComparison.OrdinalIgnoreCase));
            return val?.ValueDataRaw;
        }

        private static string? ReadString(RegistryParser.Abstractions.RegistryKey elementsKey, string elementId)
        {
            var elem = elementsKey.SubKeys.FirstOrDefault(k =>
                string.Equals(k.KeyName, elementId, StringComparison.OrdinalIgnoreCase));
            return elem == null ? null : ReadStringElement(elem);
        }

        private static string? ReadStringElement(RegistryParser.Abstractions.RegistryKey elemKey)
        {
            var val = elemKey.Values.FirstOrDefault(v =>
                string.Equals(v.ValueName, "Element", StringComparison.OrdinalIgnoreCase));
            return val?.ValueData;
        }

        private static uint? ReadDword(RegistryParser.Abstractions.RegistryKey key, string name)
        {
            var val = key.Values.FirstOrDefault(v =>
                string.Equals(v.ValueName, name, StringComparison.OrdinalIgnoreCase));
            if (val?.ValueDataRaw == null || val.ValueDataRaw.Length < 4) return null;
            return BitConverter.ToUInt32(val.ValueDataRaw, 0);
        }
    }

    /// <summary>One BCD boot entry (Boot Loader or Resume Loader).</summary>
    public class BcdBootEntry
    {
        public string ObjectId { get; set; } = "";
        public uint Type { get; set; }
        public string TypeName { get; set; } = "";

        /// <summary>The boot loader's "device" — partition that holds the loader.</summary>
        public BcdDeviceDescriptor? Device { get; set; }

        /// <summary>The boot loader's "osdevice" — partition where Windows is installed.</summary>
        public BcdDeviceDescriptor? OsDevice { get; set; }

        /// <summary>Loader binary path, e.g. "\Windows\system32\winload.exe".</summary>
        public string? Path { get; set; }

        /// <summary>Windows system root, e.g. "\Windows".</summary>
        public string? SystemRoot { get; set; }
    }

    /// <summary>A parsed BCD device descriptor — partition identification.</summary>
    public class BcdDeviceDescriptor
    {
        public uint DeviceType { get; set; }
        public string PartitionStyle { get; set; } = "";
        public uint? MbrDiskSignature { get; set; }
        public string? GptPartitionGuid { get; set; }
        public string? GptDiskGuid { get; set; }

        /// <summary>True when the descriptor identifies a specific partition we can match.</summary>
        public bool HasUsableTarget =>
            MbrDiskSignature.HasValue ||
            !string.IsNullOrEmpty(GptPartitionGuid) ||
            !string.IsNullOrEmpty(GptDiskGuid);
    }
}
