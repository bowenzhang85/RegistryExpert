using System;
using System.Text;
using System.Threading;
using RegistryParser.Cells;

// namespaces...

namespace RegistryParser.Abstractions;

// public classes...
/// <summary>
///     Represents a value that is associated with a RegistryKey
///     <remarks>Also contains references to low level structures related to a given value</remarks>
/// </summary>
public class KeyValue
{
    private static long _nextId;

    // public constructors...
    /// <summary>
    ///     Initializes a new instance of the <see cref="KeyValue" /> class.
    /// </summary>
    public KeyValue(VkCellRecord vk)
    {
        VkRecord = vk;
        InternalGuid = Interlocked.Increment(ref _nextId).ToString();
    }

    // public properties...
    /// <summary>
    ///     A unique value that can be used to find this key in a collection
    /// </summary>
    public string InternalGuid { get; }

    /// <summary>
    ///     The normalized representation of the value's value.
    /// </summary>
    public string ValueData
    {
        get
        {
            if (VkRecord.ValueData is byte[]) return BitConverter.ToString((byte[]) VkRecord.ValueData);

            return VkRecord.ValueData.ToString();
        }
    }

    /// <summary>
    ///     The value as stored in the hive as a series of bytes
    /// </summary>
    public byte[] ValueDataRaw => VkRecord.ValueDataRaw;

    public string ValueName => VkRecord.ValueName;

    /// <summary>
    ///     If present, the value slack as a string of bytes delimited by hyphens
    /// </summary>
    public string ValueSlack => BitConverter.ToString(VkRecord.ValueDataSlack);

    /// <summary>
    ///     The value slack as stored in the hive as a series of bytes
    /// </summary>
    public byte[] ValueSlackRaw => VkRecord.ValueDataSlack;

    /// <summary>
    ///     The values type (VKCellRecord.DataTypeEnum)
    /// </summary>
    public string ValueType => VkRecord.DataType.ToString();

    /// <summary>
    ///     The underlying VKRecord for this Key. This allows access to all info about the VK Record
    /// </summary>
    public VkCellRecord VkRecord { get; }

    // public methods...
    public override string ToString()
    {
        var sb = new StringBuilder();

        sb.AppendLine($"VK Record: {VkRecord}");

        return sb.ToString();
    }
}