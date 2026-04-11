using System.IO;
using RegistryExpert.Core;
using RegistryExpert.Core.Models;

namespace RegistryExpert.Wpf.ViewModels
{
    public class HivePickerItem : ViewModelBase
    {
        private bool _isSelected = true;

        public HivePickerItem(DiscoveredHive hive)
        {
            Hive = hive;
        }

        public DiscoveredHive Hive { get; }

        public string TypeName
        {
            get
            {
                if (Hive.DetectedType != OfflineRegistryParser.HiveType.Unknown)
                    return Hive.DetectedType.ToString();

                // For unknown types, derive a friendly name from the filename.
                // TSS pattern: {hostname}_reg_{name}.hiv → show "{name}"
                var nameWithoutExt = Path.GetFileNameWithoutExtension(Hive.FilePath);
                var regIndex = nameWithoutExt.IndexOf("_reg_", StringComparison.OrdinalIgnoreCase);
                if (regIndex >= 0)
                    return nameWithoutExt.Substring(regIndex + 5);

                // Fallback: just show the filename without extension
                return nameWithoutExt;
            }
        }

        public string RelativePath => Hive.RelativePath;

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}
