using System.Collections.ObjectModel;
using RegistryParser.Abstractions;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// TreeView binding model wrapping a RegistryKey with lazy-loading children.
    /// </summary>
    public class RegistryKeyNode : ViewModelBase
    {
        private static readonly RegistryKeyNode DummyChild = new();

        private readonly RegistryKey? _registryKey;
        private bool _isExpanded;
        private bool _isSelected;
        private bool _childrenLoaded;

        /// <summary>Root node display name override (e.g. "SYSTEM" instead of raw key name).</summary>
        private readonly string? _displayNameOverride;

        /// <summary>Private constructor for dummy placeholder node.</summary>
        private RegistryKeyNode()
        {
            _registryKey = null;
            Children = new ObservableCollection<RegistryKeyNode>();
        }

        public RegistryKeyNode(RegistryKey key, string? displayNameOverride = null)
        {
            _registryKey = key;
            _displayNameOverride = displayNameOverride;
            Children = new ObservableCollection<RegistryKeyNode>();

            // Add dummy child if this key has subkeys (makes expand arrow appear)
            if (key.SubKeys?.Count > 0)
            {
                Children.Add(DummyChild);
            }
        }

        /// <summary>The underlying registry key (null only for dummy node).</summary>
        public RegistryKey? RegistryKey => _registryKey;

        /// <summary>Display name: override for root nodes, otherwise KeyName.</summary>
        public string DisplayName => _displayNameOverride ?? _registryKey?.KeyName ?? "";

        /// <summary>Parent node (null for root-level hive nodes).</summary>
        public RegistryKeyNode? Parent { get; internal set; }

        /// <summary>Whether this is a root-level hive node (has display name override).</summary>
        public bool IsRootNode => _displayNameOverride != null;

        /// <summary>Full key path from the RegistryKey.</summary>
        public string KeyPath => _registryKey?.KeyPath ?? "";

        /// <summary>Last write time of this key.</summary>
        public DateTimeOffset? LastWriteTime => _registryKey?.LastWriteTime;

        /// <summary>Number of subkeys.</summary>
        public int SubKeyCount => _registryKey?.SubKeys?.Count ?? 0;

        /// <summary>Number of values.</summary>
        public int ValueCount => _registryKey?.Values?.Count ?? 0;

        /// <summary>Child nodes (lazy-loaded on expand).</summary>
        public ObservableCollection<RegistryKeyNode> Children { get; }

        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (SetProperty(ref _isExpanded, value) && value)
                {
                    LoadChildren();
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        /// <summary>Whether to show a separator line above this root node (not the first hive).</summary>
        private bool _showSeparator;
        public bool ShowSeparator
        {
            get => _showSeparator;
            set => SetProperty(ref _showSeparator, value);
        }

        /// <summary>Load real children, replacing the dummy placeholder.</summary>
        private void LoadChildren()
        {
            if (_childrenLoaded || _registryKey == null)
                return;

            _childrenLoaded = true;
            Children.Clear();

            if (_registryKey.SubKeys != null)
            {
                foreach (var subKey in _registryKey.SubKeys.OrderBy(k => k.KeyName))
                {
                    var child = new RegistryKeyNode(subKey) { Parent = this };
                    Children.Add(child);
                }
            }
        }

        /// <summary>Force reload of children (used after navigation).</summary>
        public void EnsureChildrenLoaded()
        {
            if (!_childrenLoaded)
            {
                LoadChildren();
                _isExpanded = true;
                OnPropertyChanged(nameof(IsExpanded));
            }
        }
    }
}
