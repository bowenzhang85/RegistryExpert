namespace RegistryParser.Other;

public interface IRecordBase
{
    /// <summary>
    ///     The offset in the registry hive file to a record
    /// </summary>
    long AbsoluteOffset { get; }

    string Signature { get; }
}