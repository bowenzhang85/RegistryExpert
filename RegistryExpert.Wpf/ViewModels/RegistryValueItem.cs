using RegistryParser.Abstractions;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// DataGrid row model wrapping a KeyValue with formatted display.
    /// </summary>
    public class RegistryValueItem
    {
        private readonly KeyValue _keyValue;

        public RegistryValueItem(KeyValue keyValue)
        {
            _keyValue = keyValue;
            Name = string.IsNullOrEmpty(keyValue.ValueName) ? "(Default)" : keyValue.ValueName;
            Type = keyValue.ValueType ?? "Unknown";
            Data = FormatValueData(keyValue);
            ImageKey = GetImageKey(Type);
        }

        /// <summary>The underlying KeyValue.</summary>
        public KeyValue KeyValue => _keyValue;

        /// <summary>Display name ("(Default)" for empty name).</summary>
        public string Name { get; }

        /// <summary>Registry value type string (e.g. "RegSz", "RegDword").</summary>
        public string Type { get; }

        /// <summary>Formatted data string for display.</summary>
        public string Data { get; }

        /// <summary>Image key for value type icon ("reg_str", "reg_num", "reg_bin").</summary>
        public string ImageKey { get; }

        /// <summary>Raw bytes for hex dump display.</summary>
        public byte[] RawBytes => _keyValue.ValueDataRaw ?? Array.Empty<byte>();

        /// <summary>Slack bytes count.</summary>
        public int SlackSize => _keyValue.ValueSlackRaw?.Length ?? 0;

        private static string FormatValueData(KeyValue value)
        {
            var type = value.ValueType?.ToUpperInvariant() ?? "";

            switch (type)
            {
                case "REGBINARY":
                {
                    var bytes = value.ValueDataRaw;
                    if (bytes == null || bytes.Length == 0)
                        return "(zero-length binary value)";

                    var display = BitConverter.ToString(bytes, 0, Math.Min(bytes.Length, 64))
                        .Replace("-", " ");
                    return bytes.Length > 64
                        ? $"{display}... ({bytes.Length} bytes)"
                        : display;
                }

                case "REGMULTISZ":
                case "REGMULTISTRING":
                {
                    var data = value.ValueData ?? "";
                    return data.Replace("\0", " | ");
                }

                case "REGQWORD":
                {
                    var bytes = value.ValueDataRaw;
                    if (bytes != null && bytes.Length >= 8)
                    {
                        var val = BitConverter.ToUInt64(bytes, 0);
                        return $"{val} (0x{val:X})";
                    }
                    return value.ValueData ?? "";
                }

                case "REGDWORD":
                {
                    var bytes = value.ValueDataRaw;
                    if (bytes != null && bytes.Length >= 4)
                    {
                        var val = BitConverter.ToUInt32(bytes, 0);
                        return $"{val} (0x{val:X})";
                    }
                    return value.ValueData ?? "";
                }

                default:
                    return value.ValueData ?? "";
            }
        }

        internal static string GetImageKey(string valueType)
        {
            return (valueType?.ToUpperInvariant() ?? "") switch
            {
                "REGBINARY" => "reg_bin",
                "REGDWORD" or "REGQWORD" => "reg_num",
                _ => "reg_str"
            };
        }
    }
}
