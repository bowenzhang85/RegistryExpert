using RegistryExpert.Core;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// Tracks a single loaded registry hive with its parser, extractor, and tree root.
    /// </summary>
    public class LoadedHiveInfo : IDisposable
    {
        public required OfflineRegistryParser Parser { get; init; }
        public required RegistryInfoExtractor InfoExtractor { get; init; }
        public required string FilePath { get; init; }
        public required RegistryKeyNode RootNode { get; init; }

        public OfflineRegistryParser.HiveType HiveType => Parser.CurrentHiveType;

        public void Dispose()
        {
            InfoExtractor.Dispose();
            Parser.Dispose();
        }
    }
}
