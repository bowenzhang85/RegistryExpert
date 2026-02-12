using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using RegistryParser.Other;

using static RegistryParser.Other.Helpers;

namespace RegistryParser;

public class RegistryBase : IRegistry
{
    public RegistryBase()
    {
        throw new NotSupportedException("Call the other constructor and pass in the path to the Registry hive!");
    }

    public RegistryBase(byte[] rawBytes, string hivePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        FileBytes = rawBytes;
        HivePath = "None";


        if (!HasValidSignature())
        {
            Debug.WriteLine("ERROR: Data in byte array is not a Registry hive (bad signature)");

            throw new ArgumentException("Data in byte array is not a Registry hive (bad signature)");
        }

        HivePath = hivePath;

        Initialize();
    }

    public RegistryBase(string hivePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (hivePath == null) throw new ArgumentNullException("hivePath cannot be null");

        if (!File.Exists(hivePath))
        {
            var fullPath = Path.GetFullPath(hivePath);
            throw new FileNotFoundException($"The specified file {fullPath} was not found.", fullPath);
        }

        var fileStream = new FileStream(hivePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var binaryReader = new BinaryReader(fileStream);

        binaryReader.BaseStream.Seek(0, SeekOrigin.Begin);

        FileBytes = binaryReader.ReadBytes((int) binaryReader.BaseStream.Length);

        binaryReader.Close();
        fileStream.Close();


        if (!HasValidSignature())
        {
            Debug.WriteLine($"ERROR: {hivePath} is not a Registry hive (bad signature)");

            throw new Exception($"{hivePath} is not a Registry hive (bad signature)");
        }

        HivePath = hivePath;

        //    Logger.Trace("Set HivePath to {0}", hivePath);

        Initialize();
    }

    public long TotalBytesRead { get; internal set; }
    public string Version { get; private set; }

    public byte[] FileBytes { get; internal set; }

    public HiveTypeEnum HiveType { get; private set; }

    public string HivePath { get; }

    public RegistryHeader Header { get; set; }

    public byte[] ReadBytesFromHive(long offset, int length)
    {
        var readLength = Math.Abs(length);

        if (readLength == 0) return Array.Empty<byte>();

        // Guard against negative or out-of-bounds offsets
        if (offset < 0 || offset >= FileBytes.Length)
        {
            // Return zero-filled array of requested size so callers (BitConverter, etc.) don't crash
            return new byte[readLength];
        }

        var remaining = FileBytes.Length - offset;

        if (readLength <= remaining)
        {
            // Normal case: enough data available
            var r = new ArraySegment<byte>(FileBytes, (int) offset, readLength);
            return r.ToArray();
        }

        // Partial read: not enough data remaining. Return a full-size zero-padded array
        // so callers always get the exact number of bytes they requested.
        var result = new byte[readLength];
        Array.Copy(FileBytes, (int) offset, result, 0, (int) remaining);
        return result;
    }

    internal void Initialize()
    {
        var header = ReadBytesFromHive(0, 4096);

        //    Logger.Trace("Getting header");

        Header = new RegistryHeader(header);

        var fileNameSegs = Header.FileName.Split('\\');

        var fNameBase = fileNameSegs.Last().ToLowerInvariant();
        
        Debug.WriteLine($"Got hive header. Embedded file name {Header.FileName}. Base Name {fNameBase}");

        switch (fNameBase)
        {
            case "ntuser.dat":
                HiveType = HiveTypeEnum.NtUser;
                break;
            case "sam":
                HiveType = HiveTypeEnum.Sam;
                break;
            case "security":
                HiveType = HiveTypeEnum.Security;
                break;
            case "software":
                HiveType = HiveTypeEnum.Software;
                break;
            case "system":
                HiveType = HiveTypeEnum.System;
                break;
            case "drivers":
                HiveType = HiveTypeEnum.Drivers;
                break;
            case "usrclass.dat":
                HiveType = HiveTypeEnum.UsrClass;
                break;
            case "components":
                HiveType = HiveTypeEnum.Components;
                break;
            case "bcd":
                HiveType = HiveTypeEnum.Bcd;
                break;
            case "amcache.hve":
            case "amcache.hve.tmp":
                HiveType = HiveTypeEnum.Amcache;
                break;
            case "syscache.hve":
                HiveType = HiveTypeEnum.Syscache;
                break;
            case "elam": 
                HiveType = HiveTypeEnum.Elam;
                break;
            case "default": 
                HiveType = HiveTypeEnum.Default;
                break;
            case "Vsmidk": 
                HiveType = HiveTypeEnum.Vsmidk;
                break;
            case "BcdTemplate":
                HiveType = HiveTypeEnum.BcdTemplate;
                break;
            case "bbi": 
                HiveType = HiveTypeEnum.Bbi;
                break;
            case "userdiff": 
                HiveType = HiveTypeEnum.Userdiff;
                break;
            case "user.dat": 
                HiveType = HiveTypeEnum.User;
                break;
            case "userclasses.dat": 
                HiveType = HiveTypeEnum.UserClasses;
                break;
            case "settings.dat": 
                HiveType = HiveTypeEnum.settings;
                break;
            case "registry.dat": 
                HiveType = HiveTypeEnum.Registry;
                break;
            default:
                HiveType = HiveTypeEnum.Other;
                break;
        }

        //    Logger.Trace("Hive is a {0} hive", HiveType);

        Version = $"{Header.MajorVersion}.{Header.MinorVersion}";

        //   Logger.Trace("Hive version is {0}", version);
    }

    public bool HasValidSignature()
    {
        var sig = BitConverter.ToInt32(FileBytes, 0);

        return sig.Equals(RegfSignature);
    }
}
