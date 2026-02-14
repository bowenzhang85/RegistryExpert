using System;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Text;
using RegistryParser.Other;

using static RegistryParser.Other.Helpers;

namespace RegistryParser;

public class RegistryBase : IRegistry, IDisposable
{
    private MemoryMappedFile? _mmf;
    private MemoryMappedViewAccessor? _accessor;
    private long _fileLength;
    private byte[]? _lazyFileBytes;
    private bool _useMemoryMap;
    private bool _disposed;
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

        _fileLength = new FileInfo(hivePath).Length;
        _mmf = MemoryMappedFile.CreateFromFile(hivePath, FileMode.Open, null, 0, MemoryMappedFileAccess.Read);
        _accessor = _mmf.CreateViewAccessor(0, _fileLength, MemoryMappedFileAccess.Read);
        _useMemoryMap = true;

        if (!HasValidSignature())
        {
            Debug.WriteLine($"ERROR: {hivePath} is not a Registry hive (bad signature)");
            DisposeMemoryMap();
            throw new Exception($"{hivePath} is not a Registry hive (bad signature)");
        }

        HivePath = hivePath;

        Initialize();
    }

    public long TotalBytesRead { get; internal set; }
    public string Version { get; private set; }

    public byte[] FileBytes
    {
        get
        {
            if (_useMemoryMap)
            {
                if (_lazyFileBytes == null)
                {
                    if (_fileLength > int.MaxValue)
                        throw new NotSupportedException($"File size ({_fileLength} bytes) exceeds maximum supported size for full-file byte array materialization.");
                    _lazyFileBytes = new byte[_fileLength];
                    _accessor!.ReadArray(0, _lazyFileBytes, 0, (int)_fileLength);
                }
                return _lazyFileBytes;
            }
            return _lazyFileBytes!;
        }
        internal set
        {
            _lazyFileBytes = value;
            _fileLength = value.Length;
        }
    }

    public HiveTypeEnum HiveType { get; private set; }

    public string HivePath { get; }

    public RegistryHeader Header { get; set; }

    public byte[] ReadBytesFromHive(long offset, int length)
    {
        var readLength = Math.Abs(length);

        if (readLength == 0) return Array.Empty<byte>();

        if (offset < 0 || offset >= _fileLength)
        {
            return new byte[readLength];
        }

        var remaining = _fileLength - offset;

        if (_useMemoryMap && _accessor != null)
        {
            if (readLength <= remaining)
            {
                var result = new byte[readLength];
                _accessor.ReadArray(offset, result, 0, readLength);
                return result;
            }

            var padded = new byte[readLength];
            _accessor.ReadArray(offset, padded, 0, (int)remaining);
            return padded;
        }

        if (readLength <= remaining)
        {
            var r = new ArraySegment<byte>(_lazyFileBytes!, (int)offset, readLength);
            return r.ToArray();
        }

        var result2 = new byte[readLength];
        Array.Copy(_lazyFileBytes!, (int)offset, result2, 0, (int)remaining);
        return result2;
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
        if (_useMemoryMap && _accessor != null)
        {
            var sig = _accessor.ReadInt32(0);
            return sig.Equals(RegfSignature);
        }
        var sigBytes = BitConverter.ToInt32(_lazyFileBytes!, 0);
        return sigBytes.Equals(RegfSignature);
    }

    public long FileLength => _fileLength;

    private void DisposeMemoryMap()
    {
        _accessor?.Dispose();
        _accessor = null;
        _mmf?.Dispose();
        _mmf = null;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                DisposeMemoryMap();
                _lazyFileBytes = null;
            }
            _disposed = true;
        }
    }
}
