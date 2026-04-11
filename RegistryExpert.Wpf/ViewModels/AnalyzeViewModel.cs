using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Input;
using RegistryExpert.Core;
using RegistryExpert.Core.Models;
using HiveType = RegistryExpert.Core.OfflineRegistryParser.HiveType;

namespace RegistryExpert.Wpf.ViewModels
{
    public enum ContentMode
    {
        DefaultGrid,
        Services,
        NetworkInterfaces,
        Firewall,
        DeviceManager,
        ScheduledTasks,
        CertificateStores,
        RolesFeatures,
        DiskPartitions,
        PhysicalDisks,
        Health,
        CbsPackages,
        ComponentsOverview,
        GroupPolicy
    }

    public class AnalyzeViewModel : ViewModelBase
    {
        // ── Inner types ─────────────────────────────────────────────────────

        public class CategoryItem : ViewModelBase
        {
            public string Name { get; init; } = "";
            public string Key { get; init; } = "";
            public string IconPath { get; init; } = "";
            public bool IsEnabled { get; init; }
            public string? Tooltip { get; init; }
        }

        public class SubcategoryItem : ViewModelBase
        {
            private bool _isActive;
            private bool _isEnabled = true;

            public string Title { get; init; } = "";

            public bool IsEnabled
            {
                get => _isEnabled;
                set => SetProperty(ref _isEnabled, value);
            }

            public string? Tooltip { get; init; }
            public ICommand SelectCommand { get; init; } = null!;

            public bool IsActive
            {
                get => _isActive;
                set => SetProperty(ref _isActive, value);
            }
        }

        public class AnalyzeGridRow
        {
            public string Column1 { get; init; } = "";
            public string Column2 { get; init; } = "";
            public string Column3 { get; init; } = "";
            public string Column4 { get; init; } = "";
            public string Column5 { get; init; } = "";
            public bool IsSubSection { get; init; }
            public bool IsWarning { get; init; }
            public bool IsHealthy { get; init; }
            public AnalysisItem? SourceItem { get; init; }
            public object? Tag { get; init; }
        }

        public class FirewallProfileItem : ViewModelBase
        {
            private bool _isActive;
            public string Name { get; init; } = "";        // "Domain", "Private", "Public"
            public string ProfileKey { get; init; } = "";   // Same as Name for GetFirewallRulesForProfile
            public string Label { get; init; } = "";        // "✅ Domain: Enabled" or "❌ Private: Disabled"
            public bool IsProfileEnabled { get; init; }     // Whether the profile is enabled in registry
            public ICommand SelectCommand { get; init; } = null!;

            public bool IsActive
            {
                get => _isActive;
                set => SetProperty(ref _isActive, value);
            }
        }

        public class DeviceTreeNode : ViewModelBase
        {
            private bool _isExpanded;
            private bool _isSelected;

            public string DisplayName { get; init; } = "";
            public object Tag { get; init; } = null!;  // DeviceClassItem or DeviceItem
            public ObservableCollection<DeviceTreeNode> Children { get; } = new();
            public string ForegroundBrushKey { get; init; } = "";  // "ErrorBrush", "TextDisabledBrush", or "" for default

            public bool IsExpanded
            {
                get => _isExpanded;
                set => SetProperty(ref _isExpanded, value);
            }

            public bool IsSelected
            {
                get => _isSelected;
                set => SetProperty(ref _isSelected, value);
            }
        }

        public class ScheduledTaskTreeNode : ViewModelBase
        {
            private bool _isExpanded;
            private bool _isSelected;

            public string DisplayName { get; init; } = "";
            public object? Tag { get; init; }  // AnalysisItem, CertificateStoreInfo, or CertificateInfo
            public bool IsFolder { get; init; }
            public ObservableCollection<ScheduledTaskTreeNode> Children { get; } = new();
            public string ForegroundBrushKey { get; init; } = "";  // "ErrorBrush", "TextDisabledBrush", or "" for default

            public bool IsExpanded
            {
                get => _isExpanded;
                set => SetProperty(ref _isExpanded, value);
            }

            public bool IsSelected
            {
                get => _isSelected;
                set => SetProperty(ref _isSelected, value);
            }
        }

        /// <summary>
        /// A row in the GPResult-style scrollable document.
        /// Can be a section header bar or a policy setting row.
        /// </summary>
        public class GpDocumentRow
        {
            public bool IsSectionHeader { get; init; }
            public bool IsListChild { get; init; }             // Indented child of enabledList policy
            public string CategoryPath { get; init; } = "";    // Full ADMX path for headers, e.g. "Network/BITS"
            public int SettingCount { get; init; }              // Number of settings in this section (headers only)
            public string PolicyName { get; init; } = "";       // ADMX display name or raw value name (rows only)
            public string SettingValue { get; init; } = "";     // Friendly value / Enabled / Disabled (rows only)
            public GroupPolicyItem? PolicyItem { get; init; }   // Full data for detail pane (rows only)
        }

        // ── Backing fields ──────────────────────────────────────────────────

        private CategoryItem? _selectedCategory;
        private ContentMode _currentMode;
        private string _contentHeader = "";
        private string _detailRegistryPath = "";
        private string _detailValueText = "";
        private bool _showSubcategories;
        private bool _showSubTabs;
        private string _infoBannerText = "";
        private AnalyzeGridRow? _selectedGridRow;
        private string _gridColumn1Header = "Property";
        private string _gridColumn2Header = "Value";
        private string _gridColumn3Header = "";
        private string _gridColumn4Header = "";
        private string _gridColumn5Header = "";
        private int _gridColumnCount = 2;
        private double _gridColumn1Star = 2;
        private double _gridColumn2Star = 3;
        private double _gridColumn3Star = 1;
        private double _gridColumn4Star = 1;

        // ── State ───────────────────────────────────────────────────────────

        private readonly IReadOnlyList<LoadedHiveInfo> _loadedHives;
        private readonly Dictionary<string, List<LoadedHiveInfo>> _categoryHiveMap = new();
        private readonly HashSet<string> _enabledCategories = new();
        private readonly Dictionary<string, List<AnalysisSection>> _contentCache = new();
        private readonly Dictionary<AnalysisSection, LoadedHiveInfo> _sectionHiveMap = new();

        private OfflineRegistryParser? _activeParser;
        private RegistryInfoExtractor? _activeExtractor;

        private List<AnalysisSection>? _currentSections;
        private AnalysisSection? _currentSection;

        // Services state
        private List<ServiceInfo>? _allServicesCache;
        private string _activeServiceFilter = "All";

        // Firewall state
        private List<FirewallRuleInfo> _currentFirewallRules = new();
        private string _currentFirewallProfile = "";

        // Health state
        private Dictionary<LoadedHiveInfo, List<AnalysisSection>>? _healthCache;
        private LoadedHiveInfo? _activeHealthHive;

        // ── Properties ──────────────────────────────────────────────────────

        public ObservableCollection<CategoryItem> Categories { get; } = new();

        public CategoryItem? SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                if (SetProperty(ref _selectedCategory, value))
                    OnSelectedCategoryChanged();
            }
        }

        public ObservableCollection<SubcategoryItem> Subcategories { get; } = new();
        public ObservableCollection<SubcategoryItem> SubTabs { get; } = new();

        public ContentMode CurrentMode
        {
            get => _currentMode;
            set => SetProperty(ref _currentMode, value);
        }

        public string ContentHeader
        {
            get => _contentHeader;
            set => SetProperty(ref _contentHeader, value);
        }

        public string DetailRegistryPath
        {
            get => _detailRegistryPath;
            set => SetProperty(ref _detailRegistryPath, value);
        }

        public string DetailValueText
        {
            get => _detailValueText;
            set => SetProperty(ref _detailValueText, value);
        }

        public bool ShowSubcategories
        {
            get => _showSubcategories;
            set => SetProperty(ref _showSubcategories, value);
        }

        public bool ShowSubTabs
        {
            get => _showSubTabs;
            set => SetProperty(ref _showSubTabs, value);
        }

        public string InfoBannerText
        {
            get => _infoBannerText;
            set => SetProperty(ref _infoBannerText, value);
        }

        // ── Default grid properties ─────────────────────────────────────────

        public ObservableCollection<AnalyzeGridRow> DefaultGridRows { get; } = new();

        public AnalyzeGridRow? SelectedGridRow
        {
            get => _selectedGridRow;
            set
            {
                if (SetProperty(ref _selectedGridRow, value))
                    OnGridRowSelected();
            }
        }

        public string GridColumn1Header
        {
            get => _gridColumn1Header;
            set => SetProperty(ref _gridColumn1Header, value);
        }

        public string GridColumn2Header
        {
            get => _gridColumn2Header;
            set => SetProperty(ref _gridColumn2Header, value);
        }

        public string GridColumn3Header
        {
            get => _gridColumn3Header;
            set => SetProperty(ref _gridColumn3Header, value);
        }

        public string GridColumn4Header
        {
            get => _gridColumn4Header;
            set => SetProperty(ref _gridColumn4Header, value);
        }

        public string GridColumn5Header
        {
            get => _gridColumn5Header;
            set => SetProperty(ref _gridColumn5Header, value);
        }

        public int GridColumnCount
        {
            get => _gridColumnCount;
            set => SetProperty(ref _gridColumnCount, value);
        }

        public double GridColumn1Star
        {
            get => _gridColumn1Star;
            set => SetProperty(ref _gridColumn1Star, value);
        }

        public double GridColumn2Star
        {
            get => _gridColumn2Star;
            set => SetProperty(ref _gridColumn2Star, value);
        }

        public double GridColumn3Star
        {
            get => _gridColumn3Star;
            set => SetProperty(ref _gridColumn3Star, value);
        }

        public double GridColumn4Star
        {
            get => _gridColumn4Star;
            set => SetProperty(ref _gridColumn4Star, value);
        }

        // ── Services properties ─────────────────────────────────────────────

        public ObservableCollection<ServiceInfo> FilteredServices { get; } = new();

        private ServiceInfo? _selectedService;
        public ServiceInfo? SelectedService
        {
            get => _selectedService;
            set
            {
                if (SetProperty(ref _selectedService, value))
                    OnServiceSelected();
            }
        }

        private bool _showServiceSearch;
        public bool ShowServiceSearch
        {
            get => _showServiceSearch;
            set => SetProperty(ref _showServiceSearch, value);
        }

        private string _serviceSearchText = "";
        public string ServiceSearchText
        {
            get => _serviceSearchText;
            set
            {
                if (SetProperty(ref _serviceSearchText, value))
                    ApplyServiceFilter(_activeServiceFilter);
            }
        }

        // ── CBS Packages search/filter properties ───────────────────────────

        private bool _showCbsSearch;
        public bool ShowCbsSearch
        {
            get => _showCbsSearch;
            set => SetProperty(ref _showCbsSearch, value);
        }

        private string _cbsSearchText = "";
        public string CbsSearchText
        {
            get => _cbsSearchText;
            set
            {
                if (SetProperty(ref _cbsSearchText, value))
                    FilterCbsPackages();
            }
        }

        private bool _cbsDismFilter;
        public bool CbsDismFilter
        {
            get => _cbsDismFilter;
            set
            {
                if (SetProperty(ref _cbsDismFilter, value))
                    FilterCbsPackages();
            }
        }

        // Cached CBS All Packages data for filtering
        private List<(string group, string package, string state, string installed, string user, int visibility, AnalysisItem item)> _allCbsPackagesData = new();

        // ── Network properties ──────────────────────────────────────────────

        public ObservableCollection<NetworkAdapterItem> NetworkAdapters { get; } = new();

        private NetworkAdapterItem? _selectedNetworkAdapter;
        public NetworkAdapterItem? SelectedNetworkAdapter
        {
            get => _selectedNetworkAdapter;
            set
            {
                if (SetProperty(ref _selectedNetworkAdapter, value))
                    OnNetworkAdapterSelected();
            }
        }

        public ObservableCollection<NetworkPropertyItem> NetworkProperties { get; } = new();

        private NetworkPropertyItem? _selectedNetworkProperty;
        public NetworkPropertyItem? SelectedNetworkProperty
        {
            get => _selectedNetworkProperty;
            set
            {
                if (SetProperty(ref _selectedNetworkProperty, value))
                    OnNetworkPropertySelected();
            }
        }

        // ── Firewall properties ────────────────────────────────────────────

        public ObservableCollection<FirewallProfileItem> FirewallProfiles { get; } = new();

        private string _firewallDirection = "Inbound";
        public string FirewallDirection
        {
            get => _firewallDirection;
            set
            {
                if (SetProperty(ref _firewallDirection, value))
                    PopulateFirewallGrid();
            }
        }

        public ICommand SetFirewallDirectionCommand { get; }

        public ObservableCollection<FirewallRuleInfo> FirewallRules { get; } = new();

        private string _firewallRulesHeader = "Firewall Rules";
        public string FirewallRulesHeader
        {
            get => _firewallRulesHeader;
            set => SetProperty(ref _firewallRulesHeader, value);
        }

        private FirewallRuleInfo? _selectedFirewallRule;
        public FirewallRuleInfo? SelectedFirewallRule
        {
            get => _selectedFirewallRule;
            set
            {
                if (SetProperty(ref _selectedFirewallRule, value))
                    OnFirewallRuleSelected();
            }
        }

        // ── Device Manager properties ──────────────────────────────────────

        public ObservableCollection<DeviceTreeNode> DeviceManagerNodes { get; } = new();

        private DeviceTreeNode? _selectedDeviceNode;
        public DeviceTreeNode? SelectedDeviceNode
        {
            get => _selectedDeviceNode;
            set
            {
                if (SetProperty(ref _selectedDeviceNode, value))
                    OnDeviceNodeSelected();
            }
        }

        public ObservableCollection<DevicePropertyItem> DeviceProperties { get; } = new();
        public ObservableCollection<DevicePropertyItem> DriverProperties { get; } = new();

        private DevicePropertyItem? _selectedDeviceProperty;
        public DevicePropertyItem? SelectedDeviceProperty
        {
            get => _selectedDeviceProperty;
            set
            {
                if (SetProperty(ref _selectedDeviceProperty, value))
                    OnDevicePropertySelected();
            }
        }

        private DevicePropertyItem? _selectedDriverProperty;
        public DevicePropertyItem? SelectedDriverProperty
        {
            get => _selectedDriverProperty;
            set
            {
                if (SetProperty(ref _selectedDriverProperty, value))
                    OnDriverPropertySelected();
            }
        }

        private bool _showDriverTab;
        public bool ShowDriverTab
        {
            get => _showDriverTab;
            set => SetProperty(ref _showDriverTab, value);
        }

        private bool _isDevicePropertiesTabActive = true;
        public bool IsDevicePropertiesTabActive
        {
            get => _isDevicePropertiesTabActive;
            set => SetProperty(ref _isDevicePropertiesTabActive, value);
        }

        private bool _isDriverDetailsTabActive;
        public bool IsDriverDetailsTabActive
        {
            get => _isDriverDetailsTabActive;
            set => SetProperty(ref _isDriverDetailsTabActive, value);
        }

        public ObservableCollection<SubcategoryItem> DeviceDetailTabs { get; } = new();

        private string _deviceDetailHeader = "Device Details";
        public string DeviceDetailHeader
        {
            get => _deviceDetailHeader;
            set => SetProperty(ref _deviceDetailHeader, value);
        }

        // ── Scheduled Tasks properties ─────────────────────────────────────

        public ObservableCollection<ScheduledTaskTreeNode> ScheduledTaskNodes { get; } = new();

        private ScheduledTaskTreeNode? _selectedScheduledTaskNode;
        public ScheduledTaskTreeNode? SelectedScheduledTaskNode
        {
            get => _selectedScheduledTaskNode;
            set
            {
                if (SetProperty(ref _selectedScheduledTaskNode, value))
                    OnScheduledTaskNodeSelected();
            }
        }

        public ObservableCollection<AnalyzeGridRow> ScheduledTaskDetails { get; } = new();

        private AnalyzeGridRow? _selectedScheduledTaskDetail;
        public AnalyzeGridRow? SelectedScheduledTaskDetail
        {
            get => _selectedScheduledTaskDetail;
            set
            {
                if (SetProperty(ref _selectedScheduledTaskDetail, value))
                    OnScheduledTaskDetailSelected();
            }
        }

        private string _scheduledTasksHeader = "Task Folders";
        public string ScheduledTasksHeader
        {
            get => _scheduledTasksHeader;
            set => SetProperty(ref _scheduledTasksHeader, value);
        }

        private string _scheduledTaskDetailHeader = "Task Details";
        public string ScheduledTaskDetailHeader
        {
            get => _scheduledTaskDetailHeader;
            set => SetProperty(ref _scheduledTaskDetailHeader, value);
        }

        // ── Certificate Stores properties ──────────────────────────────────

        public ObservableCollection<ScheduledTaskTreeNode> CertStoreNodes { get; } = new();

        private ScheduledTaskTreeNode? _selectedCertStoreNode;
        public ScheduledTaskTreeNode? SelectedCertStoreNode
        {
            get => _selectedCertStoreNode;
            set
            {
                if (SetProperty(ref _selectedCertStoreNode, value))
                    OnCertStoreNodeSelected();
            }
        }

        public ObservableCollection<AnalyzeGridRow> CertificateDetails { get; } = new();

        private AnalyzeGridRow? _selectedCertificateDetail;
        public AnalyzeGridRow? SelectedCertificateDetail
        {
            get => _selectedCertificateDetail;
            set
            {
                if (SetProperty(ref _selectedCertificateDetail, value))
                    OnCertificateDetailSelected();
            }
        }

        private string _certStoresTreeHeader = "Certificate Stores";
        public string CertStoresTreeHeader
        {
            get => _certStoresTreeHeader;
            set => SetProperty(ref _certStoresTreeHeader, value);
        }

        private string _certStoreDetailHeader = "Certificate Details";
        public string CertStoreDetailHeader
        {
            get => _certStoreDetailHeader;
            set => SetProperty(ref _certStoreDetailHeader, value);
        }

        // ── Group Policy properties ────────────────────────────────────────

        public ObservableCollection<GpDocumentRow> GroupPolicyDocRows { get; } = new();

        private List<GpDocumentRow> _allGpDocRows = new();

        private GpDocumentRow? _selectedGroupPolicyRow;
        public GpDocumentRow? SelectedGroupPolicyRow
        {
            get => _selectedGroupPolicyRow;
            set
            {
                if (SetProperty(ref _selectedGroupPolicyRow, value))
                    OnGroupPolicyRowSelected();
            }
        }

        private string _groupPolicyHeader = "Group Policy";
        public string GroupPolicyHeader
        {
            get => _groupPolicyHeader;
            set => SetProperty(ref _groupPolicyHeader, value);
        }

        private string _gpSearchText = "";
        public string GpSearchText
        {
            get => _gpSearchText;
            set
            {
                if (SetProperty(ref _gpSearchText, value))
                    FilterGpDocRows();
            }
        }

        // ── Roles & Features properties ────────────────────────────────────

        public ObservableCollection<DeviceTreeNode> RolesFeatureNodes { get; } = new();

        private DeviceTreeNode? _selectedRolesFeatureNode;
        public DeviceTreeNode? SelectedRolesFeatureNode
        {
            get => _selectedRolesFeatureNode;
            set
            {
                if (SetProperty(ref _selectedRolesFeatureNode, value))
                    OnRolesFeatureNodeSelected();
            }
        }

        public ObservableCollection<AnalyzeGridRow> RoleFeatureDetails { get; } = new();

        private AnalyzeGridRow? _selectedRoleFeatureDetail;
        public AnalyzeGridRow? SelectedRoleFeatureDetail
        {
            get => _selectedRoleFeatureDetail;
            set
            {
                if (SetProperty(ref _selectedRoleFeatureDetail, value))
                    OnRoleFeatureDetailSelected();
            }
        }

        private string _rolesTreeHeader = "Roles & Features";
        public string RolesTreeHeader
        {
            get => _rolesTreeHeader;
            set => SetProperty(ref _rolesTreeHeader, value);
        }

        private string _roleFeatureDetailHeader = "Details";
        public string RoleFeatureDetailHeader
        {
            get => _roleFeatureDetailHeader;
            set => SetProperty(ref _roleFeatureDetailHeader, value);
        }

        // ── Disk Partitions (Mounted Devices) properties ──────────────────

        public ObservableCollection<MountedDeviceEntry> MountedDevices { get; } = new();

        private MountedDeviceEntry? _selectedMountedDevice;
        public MountedDeviceEntry? SelectedMountedDevice
        {
            get => _selectedMountedDevice;
            set
            {
                if (SetProperty(ref _selectedMountedDevice, value))
                    OnMountedDeviceSelected();
            }
        }

        public ObservableCollection<AnalyzeGridRow> MountedDeviceDetails { get; } = new();

        private AnalyzeGridRow? _selectedMountedDeviceDetail;
        public AnalyzeGridRow? SelectedMountedDeviceDetail
        {
            get => _selectedMountedDeviceDetail;
            set
            {
                if (SetProperty(ref _selectedMountedDeviceDetail, value))
                    OnMountedDeviceDetailSelected();
            }
        }

        private string _mountedDeviceDetailHeader = "Device Details";
        public string MountedDeviceDetailHeader
        {
            get => _mountedDeviceDetailHeader;
            set => SetProperty(ref _mountedDeviceDetailHeader, value);
        }

        // ── Physical Disks properties ─────────────────────────────────────

        public ObservableCollection<PhysicalDiskEntry> PhysicalDisksList { get; } = new();

        private PhysicalDiskEntry? _selectedPhysicalDisk;
        public PhysicalDiskEntry? SelectedPhysicalDisk
        {
            get => _selectedPhysicalDisk;
            set
            {
                if (SetProperty(ref _selectedPhysicalDisk, value))
                    OnPhysicalDiskSelected();
            }
        }

        public ObservableCollection<AnalyzeGridRow> PhysicalDiskDetails { get; } = new();

        private AnalyzeGridRow? _selectedPhysicalDiskDetail;
        public AnalyzeGridRow? SelectedPhysicalDiskDetail
        {
            get => _selectedPhysicalDiskDetail;
            set
            {
                if (SetProperty(ref _selectedPhysicalDiskDetail, value))
                    OnPhysicalDiskDetailSelected();
            }
        }

        private string _physicalDiskDetailHeader = "Disk Details";
        public string PhysicalDiskDetailHeader
        {
            get => _physicalDiskDetailHeader;
            set => SetProperty(ref _physicalDiskDetailHeader, value);
        }

        // ── Constructor ─────────────────────────────────────────────────────

        public AnalyzeViewModel(IReadOnlyList<LoadedHiveInfo> loadedHives)
        {
            _loadedHives = loadedHives;

            // Set initial active parser/extractor from first hive
            if (loadedHives.Count > 0)
            {
                _activeParser = loadedHives[0].Parser;
                _activeExtractor = loadedHives[0].InfoExtractor;
            }

            SetFirewallDirectionCommand = new RelayCommand((param) =>
            {
                if (param is string direction)
                    FirewallDirection = direction;
            });

            BuildCategoryHiveMap();
            BuildCategories();

            // Auto-select the first available category
            SelectedCategory = Categories.FirstOrDefault(c => c.IsEnabled);
        }

        // ── Category-to-hive mapping ────────────────────────────────────────

        private void BuildCategoryHiveMap()
        {
            void MapHive(LoadedHiveInfo hive, params string[] categories)
            {
                foreach (var cat in categories)
                {
                    _enabledCategories.Add(cat);
                    if (!_categoryHiveMap.TryGetValue(cat, out var list))
                        _categoryHiveMap[cat] = list = new List<LoadedHiveInfo>();
                    list.Add(hive);
                }
            }

            foreach (var hive in _loadedHives)
            {
                var ht = hive.HiveType;

                if (ht == HiveType.SYSTEM)
                    MapHive(hive, "System", "Services", "Storage", "Network", "RDP", "Health", "Software");
                else if (ht == HiveType.SOFTWARE)
                    MapHive(hive, "Profiles", "System", "Software", "Update", "Health");
                else if (ht == HiveType.COMPONENTS)
                    MapHive(hive, "Update", "Health");
                else if (ht == HiveType.SAM)
                    MapHive(hive, "Profiles", "Health");
                else if (ht == HiveType.NTUSER)
                    MapHive(hive, "Software", "Health");
                else
                    MapHive(hive, "Health");
            }
        }

        private void BuildCategories()
        {
            var categoryDefs = new (string Name, string Key, string Icon)[]
            {
                ("System", "System", "registry_icon_system.png"),
                ("Profiles", "Profiles", "registry_icon_profiles.png"),
                ("Services", "Services", "registry_icon_services.png"),
                ("Software", "Software", "registry_icon_software.png"),
                ("Storage", "Storage", "registry_icon_storage.png"),
                ("Network", "Network", "registry_icon_network.png"),
                ("RDP", "RDP", "registry_icon_rdp.png"),
                ("Update", "Update", "registry_icon_update.png"),
                ("Health", "Health", "registry_icon_health.png"),
            };

            // Available categories first, then unavailable
            var available = new List<CategoryItem>();
            var unavailable = new List<CategoryItem>();

            foreach (var (name, key, icon) in categoryDefs)
            {
                var isEnabled = _enabledCategories.Contains(key);
                var item = new CategoryItem
                {
                    Name = name,
                    Key = key,
                    IconPath = $"pack://application:,,,/Assets/{icon}",
                    IsEnabled = isEnabled,
                    Tooltip = isEnabled ? null : "No hive loaded for this category"
                };

                if (isEnabled)
                    available.Add(item);
                else
                    unavailable.Add(item);
            }

            foreach (var item in available) Categories.Add(item);
            foreach (var item in unavailable) Categories.Add(item);
        }

        // ── Category selection ──────────────────────────────────────────────

        private void OnSelectedCategoryChanged()
        {
            var cat = SelectedCategory;
            if (cat == null || !cat.IsEnabled) return;

            var key = cat.Key;

            // Hide services search when switching away from Services
            ShowServiceSearch = false;

            // Hide CBS search when switching away from CBS Packages
            ShowCbsSearch = false;

            // Clear info banner from previous subcategory
            InfoBannerText = "";

            // Clear sub-tabs from previous category (e.g. Health)
            SubTabs.Clear();
            ShowSubTabs = false;

            // Set active parser/extractor to the first hive for this category
            if (_categoryHiveMap.TryGetValue(key, out var hives) && hives.Count > 0)
            {
                var newParser = hives[0].Parser;
                var newExtractor = hives[0].InfoExtractor;

                // Invalidate caches if extractor changed
                if (!ReferenceEquals(newExtractor, _activeExtractor))
                {
                    _allServicesCache = null;
                    _healthCache = null;
                }

                _activeParser = newParser;
                _activeExtractor = newExtractor;
            }
            else
            {
                // No hives available for this category -- nothing to display
                return;
            }

            // Special handling for Services (no subcategories, uses filter buttons)
            if (key == "Services")
            {
                HandleServicesCategory();
                return;
            }

            // Special handling for Health (always re-runs, optional hive selector)
            if (key == "Health")
            {
                HandleHealthCategory();
                return;
            }

            // Regular categories — load from cache or extract
            if (!_contentCache.ContainsKey(key))
            {
                LoadCategoryData(key);
            }

            _currentSections = _contentCache.GetValueOrDefault(key);
            if (_currentSections == null || _currentSections.Count == 0)
            {
                Subcategories.Clear();
                ShowSubcategories = false;
                DefaultGridRows.Clear();
                ContentHeader = cat.Name;
                return;
            }

            // Build subcategory buttons
            BuildSubcategoryButtons(key, _currentSections);

            // Auto-select first available subcategory
            var first = Subcategories.FirstOrDefault(s => s.IsEnabled);
            if (first != null)
            {
                first.SelectCommand.Execute(null);
            }
            else
            {
                // No enabled subcategories — clear stale content from previous category
                ContentHeader = cat.Name;
                DefaultGridRows.Clear();
                CurrentMode = ContentMode.DefaultGrid;
                InfoBannerText = "";
            }
        }

        private void LoadCategoryData(string key)
        {
            var allSections = new List<AnalysisSection>();

            if (!_categoryHiveMap.TryGetValue(key, out var hives)) return;

            foreach (var hive in hives)
            {
                var ext = hive.InfoExtractor;
                List<AnalysisSection> hiveSections = key switch
                {
                    "System" => ext.GetSystemAnalysis(),
                    "Profiles" => ext.GetUserAnalysis(),
                    "Network" => ext.GetNetworkAnalysis(),
                    "RDP" => ext.GetRdpAnalysis(),
                    "Update" => ext.GetUpdateAnalysis(),
                    "Storage" => ext.GetStorageAnalysis(),
                    "Software" => ext.GetSoftwareAnalysis(),
                    _ => new List<AnalysisSection>()
                };

                // Add special System subcategories not included in GetSystemAnalysis()
                if (key == "System")
                {
                    if (hive.HiveType == HiveType.SOFTWARE)
                    {
                        try { hiveSections.Add(ext.GetActivationAnalysis()); }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Activation analysis error: {ex.Message}"); }

                        try
                        {
                            if (ext.HasCertificateStores())
                            {
                                var certData = ext.GetCertificateStoresData();
                                hiveSections.Add(new AnalysisSection
                                {
                                    Title = "\U0001f4dc Certificate Stores",
                                    Tag = certData
                                });
                            }
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Certificate stores error: {ex.Message}"); }

                        try
                        {
                            if (ext.HasGroupPolicies())
                            {
                                var gpData = ext.GetGroupPolicyData();
                                hiveSections.Add(new AnalysisSection
                                {
                                    Title = "\U0001f4dc Group Policy",
                                    Tag = gpData
                                });
                            }
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Group policy analysis error: {ex.Message}"); }
                    }

                    if (hive.HiveType == HiveType.SYSTEM)
                    {
                        try
                        {
                            if (ext.HasBootConfigurations())
                                hiveSections.Add(ext.GetBootConfigurationAnalysis());
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Boot config analysis error: {ex.Message}"); }
                    }
                }

                // Merge sections by title across hives
                foreach (var section in hiveSections)
                {
                    _sectionHiveMap[section] = hive;

                    var existing = allSections.FirstOrDefault(s => s.Title == section.Title);
                    if (existing == null)
                    {
                        allSections.Add(section);
                    }
                    else if (existing.Items.Count == 0 && section.Items.Count > 0)
                    {
                        // Replace empty section with one that has data
                        var idx = allSections.IndexOf(existing);
                        allSections[idx] = section;
                    }
                    else if (existing.Items.Count > 0 && section.Items.Count > 0)
                    {
                        // Merge items from both hives into one flat list
                        existing.Items.AddRange(section.Items);

                        // Remove from sectionHiveMap since it's now merged from multiple hives
                        _sectionHiveMap.Remove(existing);
                    }
                }
            }

            // Cross-source dedup for Security Software:
            // Priority: WSC > Installed Software > ELAM
            // WSC has richest data (ON/OFF state, definitions), Installed Software has version,
            // ELAM has start type. Remove lower-priority duplicates.
            var secSoftSection = allSections.FirstOrDefault(s => s.Title.Contains("Security Software"));
            if (secSoftSection != null)
            {
                bool systemHiveLoaded = hives.Any(h => h.HiveType == HiveType.SYSTEM);

                if (secSoftSection.Items.Count > 1)
                {
                    // Categorize items by source
                    var wscItems = secSoftSection.Items
                        .Where(i => i.SubItems != null && i.SubItems.Count > 0
                                  && i.SubItems[0].Value != "Security Service"
                                  && i.SubItems[0].Value != "Installed Software")
                        .ToList();
                    var elamItems = secSoftSection.Items
                        .Where(i => i.SubItems != null && i.SubItems.Count > 0
                                  && i.SubItems[0].Value == "Security Service")
                        .ToList();
                    var installedItems = secSoftSection.Items
                        .Where(i => i.SubItems != null && i.SubItems.Count > 0
                                  && i.SubItems[0].Value == "Installed Software")
                        .ToList();

                    static string First2Words(string name) =>
                        string.Join(' ', name.Split(' ', StringSplitOptions.RemoveEmptyEntries).Take(2));

                    static bool NamesMatch(string a, string b) =>
                        a.Contains(b, StringComparison.OrdinalIgnoreCase) ||
                        b.Contains(a, StringComparison.OrdinalIgnoreCase) ||
                        First2Words(a).Equals(First2Words(b), StringComparison.OrdinalIgnoreCase);

                    // Remove ELAM entries that match WSC or Installed Software
                    var higherPriority = wscItems.Concat(installedItems).ToList();
                    if (higherPriority.Count > 0 && elamItems.Count > 0)
                    {
                        var elamToRemove = elamItems
                            .Where(elam => higherPriority.Any(hp => NamesMatch(hp.Name, elam.Name)))
                            .ToList();
                        foreach (var item in elamToRemove)
                            secSoftSection.Items.Remove(item);
                    }

                    // Remove Installed Software entries that match WSC
                    if (wscItems.Count > 0 && installedItems.Count > 0)
                    {
                        var instToRemove = installedItems
                            .Where(inst => wscItems.Any(wsc => NamesMatch(wsc.Name, inst.Name)))
                            .ToList();
                        foreach (var item in instToRemove)
                            secSoftSection.Items.Remove(item);
                    }
                }

                // Remove ELAM hint when SYSTEM hive is loaded (regardless of ELAM count)
                if (systemHiveLoaded)
                {
                    secSoftSection.Items.RemoveAll(i =>
                        i.SubItems == null && i.Name == "Info"
                        && i.Value != null
                        && i.Value.Contains("ELAM", StringComparison.OrdinalIgnoreCase));
                }
            }

            _contentCache[key] = allSections;
        }

        // ── Subcategory management ──────────────────────────────────────────

        private void BuildSubcategoryButtons(string categoryKey, List<AnalysisSection> sections)
        {
            Subcategories.Clear();

            var isSoftwareHive = _loadedHives.Any(h => h.HiveType == HiveType.SOFTWARE);
            var isSystemHive = _loadedHives.Any(h => h.HiveType == HiveType.SYSTEM);
            var isComponentsHive = _loadedHives.Any(h => h.HiveType == HiveType.COMPONENTS);
            var isNtuserHive = _loadedHives.Any(h => h.HiveType == HiveType.NTUSER);
            var isSamHive = _loadedHives.Any(h => h.HiveType == HiveType.SAM);

            // Subcategory availability sets for System category
            var softwareSubcats = new HashSet<string> { "🪟 Build Information", "\U0001f4dc Certificate Stores", "\U0001f4dc Group Policy" };
            var systemSubcats = new HashSet<string>
            {
                "💻 Computer Information", "🔄 CPU Hyper-Threading",
                "💥 Crash Dump Configuration", "🕐 System Time Config",
                "\U0001f5a5\ufe0f Device Manager"
            };
            var bothSubcats = new HashSet<string> { "📁 Hive Information" };

            var available = new List<SubcategoryItem>();
            var unavailable = new List<SubcategoryItem>();

            foreach (var section in sections)
            {
                // Skip "Notice" sections for Update category when only COMPONENTS hive loaded
                if (categoryKey == "Update" && isComponentsHive && !isSoftwareHive && section.Title.Contains("Notice"))
                    continue;

                bool isAvailable = IsSubcategoryAvailable(categoryKey, section.Title,
                    isSoftwareHive, isSystemHive, isNtuserHive,
                    softwareSubcats, systemSubcats, bothSubcats,
                    out string reqHive);

                var sectionRef = section; // Capture for closure
                var item = new SubcategoryItem
                {
                    Title = section.Title,
                    IsEnabled = isAvailable,
                    Tooltip = isAvailable ? null : $"Requires {reqHive} hive to be loaded",
                    SelectCommand = new RelayCommand(() =>
                    {
                        if (!isAvailable) return;

                        // Switch active extractor to the section's originating hive
                        if (_sectionHiveMap.TryGetValue(sectionRef, out var sectionHive))
                        {
                            _activeParser = sectionHive.Parser;
                            _activeExtractor = sectionHive.InfoExtractor;
                        }

                        DisplaySection(sectionRef);
                        UpdateSubcategoryActiveState(sectionRef.Title);
                    })
                };

                if (isAvailable)
                    available.Add(item);
                else
                    unavailable.Add(item);
            }

            // Add placeholder buttons for subcategories that would exist if the missing hive were loaded
            var existingTitles = new HashSet<string>(available.Select(s => s.Title).Concat(unavailable.Select(s => s.Title)));

            if (categoryKey == "System")
            {
                // Missing SOFTWARE subcategories when only SYSTEM is loaded
                if (!isSoftwareHive)
                {
                    var missingSoftware = new[] { "🪟 Build Information", "🔑 Windows Activation", "\U0001f4dc Certificate Stores", "\U0001f4dc Group Policy" };
                    foreach (var title in missingSoftware)
                    {
                        if (existingTitles.Contains(title)) continue;
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = title,
                            IsEnabled = false,
                            Tooltip = "Requires SOFTWARE hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }
                }

                // Missing SYSTEM subcategories when only SOFTWARE is loaded
                if (!isSystemHive)
                {
                    var missingSystem = new[]
                    {
                        "💻 Computer Information", "🔄 CPU Hyper-Threading",
                        "💥 Crash Dump Configuration", "🕐 System Time Config",
                        "\U0001f5a5\ufe0f Device Manager", "\U0001f4bb Boot Configurations"
                    };
                    foreach (var title in missingSystem)
                    {
                        if (existingTitles.Contains(title)) continue;
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = title,
                            IsEnabled = false,
                            Tooltip = "Requires SYSTEM hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }
                }
            }

            // Add extra grayed-out Update subcategories when only COMPONENTS hive loaded
            if (categoryKey == "Update" && isComponentsHive && !isSoftwareHive)
            {
                var updateSoftwareSubcats = new[]
                {
                    "📋 Update Policy", "🏢 Windows Update for Business",
                    "📦 Delivery Optimization", "📜 Update Configuration",
                    "🔧 Servicing Stack Update (SSU)", "📦 CBS Packages"
                };

                foreach (var title in updateSoftwareSubcats)
                {
                    if (existingTitles.Contains(title)) continue;

                    unavailable.Add(new SubcategoryItem
                    {
                        Title = title,
                        IsEnabled = false,
                        Tooltip = "Requires SOFTWARE hive to be loaded",
                        SelectCommand = new RelayCommand(() => { })
                    });
                }
            }

            // Profiles: placeholders for missing SOFTWARE or NTUSER subcategories
            if (categoryKey == "Profiles")
            {
                if (!isSoftwareHive)
                {
                    var missing = new[] { "📂 User Profiles" };
                    foreach (var title in missing)
                    {
                        if (existingTitles.Contains(title)) continue;
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = title,
                            IsEnabled = false,
                            Tooltip = "Requires SOFTWARE hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }
                }


            }

            // Network: placeholders for missing SYSTEM or SOFTWARE subcategories
            if (categoryKey == "Network")
            {
                if (!isSystemHive)
                {
                    var missing = new[]
                    {
                        "🔌 Network Interfaces", "🧭 DNS Registered Adapters",
                        "📁 Network Shares", "🔑 NTLM Authentication",
                        "🔐 TLS/SSL Protocols", "🔥 Windows Firewall"
                    };
                    foreach (var title in missing)
                    {
                        if (existingTitles.Contains(title)) continue;
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = title,
                            IsEnabled = false,
                            Tooltip = "Requires SYSTEM hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }
                }

                if (!isSoftwareHive)
                {
                    // No SOFTWARE-only Network subcategories currently
                }
            }

            // RDP: all subcategories require SYSTEM
            if (categoryKey == "RDP" && !isSystemHive)
            {
                var missing = new[]
                {
                    "🍊 Citrix Detection", "🖥️ RDP Configuration",
                    "⏱️ Session Limits", "⚙️ Terminal Service",
                    "📜 RDP Licensing", "🏢 Remote Desktop Services (RDS)"
                };
                foreach (var title in missing)
                {
                    if (existingTitles.Contains(title)) continue;
                    unavailable.Add(new SubcategoryItem
                    {
                        Title = title,
                        IsEnabled = false,
                        Tooltip = "Requires SYSTEM hive to be loaded",
                        SelectCommand = new RelayCommand(() => { })
                    });
                }
            }

            // Storage: all subcategories require SYSTEM
            if (categoryKey == "Storage" && !isSystemHive)
            {
                var missing = new[] { "🔧 Filters", "💿 Mounted Devices", "💽 Physical Disks" };
                foreach (var title in missing)
                {
                    if (existingTitles.Contains(title)) continue;
                    unavailable.Add(new SubcategoryItem
                    {
                        Title = title,
                        IsEnabled = false,
                        Tooltip = "Requires SYSTEM hive to be loaded",
                        SelectCommand = new RelayCommand(() => { })
                    });
                }
            }

            // Software: placeholders for missing SOFTWARE or NTUSER subcategories
            if (categoryKey == "Software")
            {
                if (!isSoftwareHive)
                {
                    var missing = new[]
                    {
                        "📦 Installed Programs", "🚀 Startup Programs",
                        "📱 Appx Packages", "🔷 .NET Framework", "📅 Scheduled Tasks"
                    };
                    foreach (var title in missing)
                    {
                        if (existingTitles.Contains(title)) continue;
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = title,
                            IsEnabled = false,
                            Tooltip = "Requires SOFTWARE hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }

                    // Security Software: only unavailable if NEITHER SOFTWARE nor SYSTEM is loaded
                    if (!isSystemHive && !existingTitles.Contains("🛡️ Security Software"))
                    {
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = "🛡️ Security Software",
                            IsEnabled = false,
                            Tooltip = "Requires SOFTWARE or SYSTEM hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }

                    // Guest Agent: only unavailable if NEITHER SOFTWARE nor SYSTEM is loaded
                    if (!isSystemHive && !existingTitles.Contains("☁️ Guest Agent"))
                    {
                        unavailable.Add(new SubcategoryItem
                        {
                            Title = "☁️ Guest Agent",
                            IsEnabled = false,
                            Tooltip = "Requires SOFTWARE or SYSTEM hive to be loaded",
                            SelectCommand = new RelayCommand(() => { })
                        });
                    }
                }

                if (!isNtuserHive)
                {
                    // NTUSER contributes startup programs too, but SOFTWARE covers it;
                    // no unique NTUSER-only subcategories in Software currently
                }
            }

            // Available first, then unavailable (greyed out) at the end
            foreach (var item in available) Subcategories.Add(item);
            foreach (var item in unavailable) Subcategories.Add(item);

            ShowSubcategories = Subcategories.Count > 0;
        }

        private static bool IsSubcategoryAvailable(string categoryKey, string title,
            bool isSoftwareHive, bool isSystemHive, bool isNtuserHive,
            HashSet<string> softwareSubcats, HashSet<string> systemSubcats, HashSet<string> bothSubcats,
            out string reqHive)
        {
            reqHive = "";

            if (categoryKey == "Update")
            {
                if (!isSoftwareHive)
                {
                    reqHive = "SOFTWARE";
                    return false;
                }
                return true;
            }

            if (categoryKey == "System")
            {
                if (bothSubcats.Contains(title)) return true;
                if (softwareSubcats.Contains(title))
                {
                    reqHive = "SOFTWARE";
                    return isSoftwareHive;
                }
                if (systemSubcats.Contains(title))
                {
                    reqHive = "SYSTEM";
                    return isSystemHive;
                }
                return true;
            }

            // Network: some subcategories need SYSTEM, some need SOFTWARE
            if (categoryKey == "Network")
            {
                var networkSoftwareSubcats = new HashSet<string>();
                var networkSystemSubcats = new HashSet<string>
                {
                    "🔌 Network Interfaces", "🧭 DNS Registered Adapters",
                    "📁 Network Shares", "🔑 NTLM Authentication",
                    "🔐 TLS/SSL Protocols", "🔥 Windows Firewall"
                };

                if (networkSoftwareSubcats.Contains(title))
                {
                    reqHive = "SOFTWARE";
                    return isSoftwareHive;
                }
                if (networkSystemSubcats.Contains(title))
                {
                    reqHive = "SYSTEM";
                    return isSystemHive;
                }
                return true;
            }

            // Profiles: User Profiles needs SOFTWARE, typed paths/run/userassist need NTUSER
            if (categoryKey == "Profiles")
            {
                if (title.Contains("User Profiles"))
                {
                    reqHive = "SOFTWARE";
                    return isSoftwareHive;
                }
                return true;
            }

            // Software: most subcategories require SOFTWARE, but Security Software
            // and Guest Agent are available with either SOFTWARE or SYSTEM hive
            if (categoryKey == "Software")
            {
                if (title.Contains("Security Software") || title.Contains("Guest Agent"))
                {
                    if (!isSoftwareHive && !isSystemHive)
                    {
                        reqHive = "SOFTWARE or SYSTEM";
                        return false;
                    }
                    return true;
                }
                if (!isSoftwareHive)
                {
                    reqHive = "SOFTWARE";
                    return false;
                }
                return true;
            }

            // RDP: all subcategories require SYSTEM
            if (categoryKey == "RDP")
            {
                if (!isSystemHive)
                {
                    reqHive = "SYSTEM";
                    return false;
                }
                return true;
            }

            // Storage: all subcategories require SYSTEM
            if (categoryKey == "Storage")
            {
                if (!isSystemHive)
                {
                    reqHive = "SYSTEM";
                    return false;
                }
                return true;
            }

            return true;
        }

        private void UpdateSubcategoryActiveState(string activeTitle)
        {
            foreach (var sub in Subcategories)
            {
                sub.IsActive = sub.Title == activeTitle;
            }
        }

        private void UpdateSubTabActiveState(string activeTitle)
        {
            foreach (var tab in SubTabs)
            {
                tab.IsActive = tab.Title == activeTitle;
            }
        }

        // ── Display routing ─────────────────────────────────────────────────

        /// <summary>
        /// Central dispatcher — routes a section to the appropriate view type.
        /// </summary>
        private void DisplaySection(AnalysisSection section)
        {
            _currentSection = section;
            ContentHeader = section.Title;

            // Clear detail pane
            DetailRegistryPath = "";
            DetailValueText = "";

            // Clear sub-tabs (individual display methods re-populate if needed)
            SubTabs.Clear();
            ShowSubTabs = false;

            // Route to specialized views based on section title
            var title = section.Title;

            if (title.Contains("Network Interfaces"))
            {
                DisplayNetworkInterfaces(section);
                return;
            }
            if (title.Contains("Windows Firewall"))
            {
                DisplayFirewall(section);
                return;
            }
            if (title.Contains("Device Manager"))
            {
                DisplayDeviceManager(section);
                return;
            }
            if (title.Contains("Scheduled Tasks"))
            {
                DisplayScheduledTasks(section);
                return;
            }
            if (title.Contains("Certificate Stores"))
            {
                DisplayCertificateStores(section);
                return;
            }
            if (title.Contains("Group Policy"))
            {
                DisplayGroupPolicy(section);
                return;
            }
            if (title.Contains("Roles") && title.Contains("Features"))
            {
                DisplayRolesFeatures();
                return;
            }
            if (title.Contains("Mounted Devices"))
            {
                DisplayDiskPartitions();
                return;
            }
            if (title.Contains("Physical Disks"))
            {
                DisplayPhysicalDisks();
                return;
            }
            if (title.Contains("Appx") || title.Contains("AppX"))
            {
                DisplayAppxPackages(section);
                return;
            }
            if (title.Contains("CBS Packages"))
            {
                DisplayCbsPackages(section);
                return;
            }
            if (title.Contains("User Profiles"))
            {
                DisplayUserProfiles(section);
                return;
            }
            if (title.Contains("Guest Agent"))
            {
                DisplayGuestAgent(section);
                return;
            }
            if (title.Contains("Filters") && SelectedCategory?.Key == "Storage")
            {
                DisplayStorageFilters(section);
                return;
            }

            // Default: render in generic grid
            DisplayDefaultGrid(section);
        }

        // ── Default grid display ────────────────────────────────────────────

        private void DisplayDefaultGrid(AnalysisSection section)
        {
            CurrentMode = ContentMode.DefaultGrid;
            DefaultGridRows.Clear();
            InfoBannerText = "";

            // Reset column proportions to default (callers can override after)
            GridColumn1Star = 2;
            GridColumn2Star = 3;
            GridColumn3Star = 1;
            GridColumn4Star = 1;

            if (section.Items.Count == 0) return;

            // Determine column layout from data shape
            bool hasSubItems = section.Items.Any(i => i.SubItems != null && i.SubItems.Count > 0);
            bool hasValues = section.Items.Any(i => !string.IsNullOrEmpty(i.Value));

            if (hasSubItems)
            {
                // Multi-column: check sub-item counts for column layout
                var maxSubs = section.Items.Where(i => i.SubItems != null).Max(i => i.SubItems!.Count);
                
                // Special column headers for known sections
                if (section.Title.Contains("TLS") || section.Title.Contains("SSL"))
                {
                    GridColumn1Header = "Protocol";
                    GridColumn2Header = "Client";
                    GridColumn3Header = "Server";
                    GridColumnCount = 3;
                }
                else if (section.Title.Contains("Installed Programs"))
                {
                    GridColumn1Header = "Name";
                    GridColumn2Header = "Version";
                    GridColumn3Header = "Publisher";
                    GridColumn4Header = "Installed Date";
                    GridColumnCount = 4;
                    GridColumn1Star = 3;
                    GridColumn2Star = 1.2;
                    GridColumn3Star = 2;
                    GridColumn4Star = 1;
                }
                else if (section.Title.Contains("Security Software"))
                {
                    GridColumn1Header = "Name";
                    GridColumn2Header = "Type";
                    GridColumn3Header = "Status";
                    GridColumn4Header = "Definition";
                    GridColumn5Header = "Executable";
                    GridColumnCount = 5;
                    GridColumn1Star = 2.5;
                    GridColumn2Star = 1;
                    GridColumn3Star = 2;
                    GridColumn4Star = 2;
                }
                else
                {
                    if (maxSubs >= 2)
                    {
                        GridColumn1Header = "Name";
                        GridColumn2Header = "Detail 1";
                        GridColumn3Header = "Detail 2";
                        GridColumnCount = 3;
                    }
                    else
                    {
                        GridColumn1Header = "Property";
                        GridColumn2Header = "Value";
                        GridColumnCount = 2;
                    }
                }

                foreach (var item in section.Items)
                {
                    // Extract info hints into the banner instead of a grid row
                    if (item.SubItems == null && item.Name == "Info"
                        && !string.IsNullOrEmpty(item.Value))
                    {
                        InfoBannerText = item.Value;
                        continue;
                    }

                    var row = new AnalyzeGridRow
                    {
                        Column1 = item.Name,
                        Column2 = item.SubItems?.ElementAtOrDefault(0)?.Value ?? item.Value,
                        Column3 = item.SubItems?.ElementAtOrDefault(1)?.Value ?? "",
                        Column4 = item.SubItems?.ElementAtOrDefault(2)?.Value ?? "",
                        Column5 = item.SubItems?.ElementAtOrDefault(3)?.Value ?? "",
                        IsSubSection = item.IsSubSection,
                        IsWarning = item.IsWarning,
                        SourceItem = item
                    };
                    DefaultGridRows.Add(row);
                }
            }
            else if (hasValues)
            {
                // Two-column: Property / Value
                GridColumn1Header = "Property";
                GridColumn2Header = "Value";
                GridColumnCount = 2;

                foreach (var item in section.Items)
                {
                    // Extract info hints into the banner instead of a grid row
                    if (item.SubItems == null && item.Name == "Info"
                        && !string.IsNullOrEmpty(item.Value))
                    {
                        InfoBannerText = item.Value;
                        continue;
                    }

                    DefaultGridRows.Add(new AnalyzeGridRow
                    {
                        Column1 = item.Name,
                        Column2 = item.Value,
                        IsSubSection = item.IsSubSection,
                        IsWarning = item.IsWarning,
                        SourceItem = item
                    });
                }
            }
            else
            {
                // Single column: Name only
                GridColumn1Header = "Name";
                GridColumnCount = 1;

                foreach (var item in section.Items)
                {
                    DefaultGridRows.Add(new AnalyzeGridRow
                    {
                        Column1 = item.Name,
                        IsSubSection = item.IsSubSection,
                        IsWarning = item.IsWarning,
                        SourceItem = item
                    });
                }
            }
        }

        private void OnGridRowSelected()
        {
            var row = SelectedGridRow;
            if (row?.SourceItem == null)
            {
                DetailRegistryPath = "";
                DetailValueText = "";
                return;
            }

            var item = row.SourceItem;
            DetailRegistryPath = item.RegistryPath;
            DetailValueText = !string.IsNullOrEmpty(item.RegistryValue)
                ? $"{item.Name}: {item.RegistryValue}"
                : !string.IsNullOrEmpty(item.Value)
                    ? $"{item.Name}: {item.Value}"
                    : item.Name;
        }

        // ── Services view ───────────────────────────────────────────────────

        private void HandleServicesCategory()
        {
            CurrentMode = ContentMode.Services;
            ContentHeader = "Services";

            // Clear sub-tabs from previous subcategory (e.g. Profiles)
            SubTabs.Clear();
            ShowSubTabs = false;

            // Reset search and show search box
            _serviceSearchText = "";
            OnPropertyChanged(nameof(ServiceSearchText));
            ShowServiceSearch = true;

            // Load services from cache or extract
            if (_allServicesCache == null)
            {
                try
                {
                    _allServicesCache = _activeExtractor.GetAllServices();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GetAllServices error: {ex.Message}");
                    _allServicesCache = new List<ServiceInfo>();
                }
            }

            // Build filter buttons using subcategory items
            var all = _allServicesCache;
            var disabledCount = all.Count(s => s.IsDisabled);
            var autoCount = all.Count(s => s.IsAutoStart);
            var bootSystemCount = all.Count(s => s.IsBoot || s.IsSystem);
            var manualCount = all.Count(s => s.IsManual);

            Subcategories.Clear();

            var filters = new (string Key, string Label, int Count)[]
            {
                ("All", "All", all.Count),
                ("Disabled", "Disabled", disabledCount),
                ("AutoStart", "Auto-Start", autoCount),
                ("BootSystem", "Boot/System", bootSystemCount),
                ("Manual", "Manual", manualCount),
            };

            foreach (var (key, label, count) in filters)
            {
                var filterKey = key;
                Subcategories.Add(new SubcategoryItem
                {
                    Title = $"{label} ({count})",
                    IsEnabled = true,
                    SelectCommand = new RelayCommand(() =>
                    {
                        ApplyServiceFilter(filterKey);
                        UpdateSubcategoryActiveState($"{label} ({count})");
                    })
                });
            }

            ShowSubcategories = true;

            // Apply default filter
            ApplyServiceFilter("All");
            Subcategories.FirstOrDefault()!.IsActive = true;
        }

        private void ApplyServiceFilter(string filterKey)
        {
            _activeServiceFilter = filterKey;
            FilteredServices.Clear();

            if (_allServicesCache == null) return;

            var filtered = filterKey switch
            {
                "Disabled" => _allServicesCache.Where(s => s.IsDisabled),
                "AutoStart" => _allServicesCache.Where(s => s.IsAutoStart),
                "BootSystem" => _allServicesCache.Where(s => s.IsBoot || s.IsSystem),
                "Manual" => _allServicesCache.Where(s => s.IsManual),
                _ => _allServicesCache.AsEnumerable()
            };

            // Apply search filter
            var searchText = _serviceSearchText.Trim();
            if (!string.IsNullOrEmpty(searchText))
            {
                filtered = filtered.Where(s =>
                    s.ServiceName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    (s.DisplayName?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (s.ImagePath?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false));
            }

            foreach (var svc in filtered.OrderBy(s => s.ServiceName, StringComparer.OrdinalIgnoreCase))
                FilteredServices.Add(svc);
        }

        private void OnServiceSelected()
        {
            var svc = SelectedService;
            if (svc == null)
            {
                DetailRegistryPath = "";
                DetailValueText = "";
                return;
            }

            DetailRegistryPath = svc.RegistryPath;

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Service Name:   {svc.ServiceName}");
            sb.AppendLine($"Display Name:   {svc.DisplayName}");
            sb.AppendLine($"Description:    {svc.Description}");
            sb.AppendLine($"Start Type:     {svc.StartTypeName}");
            if (svc.IsDelayedAutoStart)
                sb.AppendLine($"Delayed Start:  Yes");
            sb.AppendLine($"Image Path:     {svc.ImagePath}");
            sb.AppendLine($"Registry Path:  {svc.RegistryPath}");

            DetailValueText = sb.ToString();
        }

        private void HandleHealthCategory()
        {
            ContentHeader = "Health";
            SubTabs.Clear();
            ShowSubTabs = false;

            // Get all hives mapped to Health
            if (!_categoryHiveMap.TryGetValue("Health", out var hives) || hives.Count == 0)
                return;

            // Build health cache (run analysis per hive, cache results)
            if (_healthCache == null)
            {
                _healthCache = new Dictionary<LoadedHiveInfo, List<AnalysisSection>>();
                foreach (var hive in hives)
                {
                    try
                    {
                        var sections = hive.InfoExtractor.GetHealthAnalysis();
                        _healthCache[hive] = sections;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Health analysis error for {hive.FilePath}: {ex.Message}");
                        _healthCache[hive] = new List<AnalysisSection>();
                    }
                }
            }

            if (hives.Count == 1)
            {
                // Single hive: show sections directly as subcategory buttons
                _activeHealthHive = hives[0];
                var sections = _healthCache.GetValueOrDefault(hives[0]) ?? new();
                BuildHealthSectionButtons(sections);
            }
            else
            {
                // Multiple hives: show hive selector as subcategory buttons
                Subcategories.Clear();
                foreach (var hive in hives)
                {
                    var label = $"{hive.Parser.FriendlyName} ({System.IO.Path.GetFileName(hive.FilePath)})";
                    var hiveRef = hive; // capture for closure
                    Subcategories.Add(new SubcategoryItem
                    {
                        Title = label,
                        IsEnabled = true,
                        Tooltip = hive.FilePath,
                        SelectCommand = new RelayCommand(() => SelectHealthHive(hiveRef, label))
                    });
                }
                ShowSubcategories = true;

                // Auto-select first hive
                Subcategories.First().SelectCommand.Execute(null);
            }
        }

        private void SelectHealthHive(LoadedHiveInfo hive, string label)
        {
            _activeHealthHive = hive;
            _activeParser = hive.Parser;
            _activeExtractor = hive.InfoExtractor;
            UpdateSubcategoryActiveState(label);

            var sections = _healthCache?.GetValueOrDefault(hive) ?? new();
            BuildHealthSectionButtons(sections, useTabs: true);
        }

        private void BuildHealthSectionButtons(List<AnalysisSection> sections, bool useTabs = false)
        {
            var target = useTabs ? SubTabs : Subcategories;
            target.Clear();

            foreach (var section in sections)
            {
                var sectionRef = section;
                target.Add(new SubcategoryItem
                {
                    Title = section.Title,
                    IsEnabled = true,
                    Tooltip = null,
                    SelectCommand = new RelayCommand(() =>
                    {
                        DisplayDefaultGrid(sectionRef);
                        ContentHeader = sectionRef.Title;
                        if (useTabs)
                            UpdateSubTabActiveState(sectionRef.Title);
                        else
                            UpdateSubcategoryActiveState(sectionRef.Title);
                    })
                });
            }

            if (useTabs)
                ShowSubTabs = sections.Count > 0;
            else
                ShowSubcategories = sections.Count > 0;

            // Auto-select first section
            if (target.Count > 0)
                target[0].SelectCommand.Execute(null);
        }

        // ── Network Interfaces view ────────────────────────────────────────

        private void DisplayNetworkInterfaces(AnalysisSection section)
        {
            CurrentMode = ContentMode.NetworkInterfaces;

            NetworkAdapters.Clear();
            NetworkProperties.Clear();

            // Parse section items into NetworkAdapterItem objects
            // Each top-level item with SubItems represents an adapter
            foreach (var item in section.Items)
            {
                if (item.SubItems == null || item.SubItems.Count == 0) continue;

                var adapter = new NetworkAdapterItem
                {
                    DisplayName = item.Name,
                    RegistryPath = item.RegistryPath,
                    FullGuid = item.Value ?? ""
                };

                foreach (var sub in item.SubItems)
                {
                    adapter.Properties.Add(new NetworkPropertyItem
                    {
                        Name = sub.Name,
                        Value = sub.Value,
                        RegistryValueName = sub.RegistryValue ?? "",
                        RegistryPath = sub.RegistryPath
                    });
                }

                NetworkAdapters.Add(adapter);
            }

            // Auto-select first adapter
            if (NetworkAdapters.Count > 0)
                SelectedNetworkAdapter = NetworkAdapters[0];
        }

        private void OnNetworkAdapterSelected()
        {
            NetworkProperties.Clear();

            var adapter = SelectedNetworkAdapter;
            if (adapter == null)
            {
                DetailRegistryPath = "";
                DetailValueText = "";
                return;
            }

            DetailRegistryPath = adapter.RegistryPath;
            DetailValueText = adapter.DisplayName;

            foreach (var prop in adapter.Properties)
                NetworkProperties.Add(prop);
        }

        private void OnNetworkPropertySelected()
        {
            var prop = SelectedNetworkProperty;
            if (prop == null) return;

            DetailRegistryPath = prop.RegistryPath;
            DetailValueText = $"{prop.Name}: {prop.Value}";
        }

        private void DisplayFirewall(AnalysisSection section)
        {
            CurrentMode = ContentMode.Firewall;

            FirewallProfiles.Clear();
            FirewallRules.Clear();

            // Define profiles: registry key name, display name
            var profiles = new[]
            {
                ("DomainProfile", "Domain"),
                ("StandardProfile", "Private"),
                ("PublicProfile", "Public")
            };

            foreach (var (registryKey, displayName) in profiles)
            {
                var isEnabled = _activeExtractor.IsFirewallProfileEnabled(registryKey);
                var statusIcon = isEnabled ? "✅" : "❌";
                var statusText = isEnabled ? "Enabled" : "Disabled";
                var profileKey = displayName; // Used for GetFirewallRulesForProfile

                var capturedKey = profileKey;
                var item = new FirewallProfileItem
                {
                    Name = displayName,
                    ProfileKey = capturedKey,
                    Label = $"{statusIcon} {displayName}: {statusText}",
                    IsProfileEnabled = isEnabled,
                    SelectCommand = new RelayCommand(() => SelectFirewallProfile(capturedKey))
                };

                FirewallProfiles.Add(item);
            }

            // Reset direction to Inbound
            _firewallDirection = "Inbound";
            OnPropertyChanged(nameof(FirewallDirection));

            // Auto-select first profile
            if (FirewallProfiles.Count > 0)
            {
                SelectFirewallProfile(FirewallProfiles[0].ProfileKey);
            }
        }

        private void SelectFirewallProfile(string profileKey)
        {
            _currentFirewallProfile = profileKey;

            // Update active state on profile buttons
            foreach (var p in FirewallProfiles)
                p.IsActive = p.ProfileKey == profileKey;

            // Get rules for this profile
            try
            {
                _currentFirewallRules = _activeExtractor.GetFirewallRulesForProfile(profileKey);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetFirewallRulesForProfile error: {ex.Message}");
                _currentFirewallRules = new List<FirewallRuleInfo>();
            }

            PopulateFirewallGrid();
        }

        private void PopulateFirewallGrid()
        {
            FirewallRules.Clear();

            var direction = FirewallDirection;
            var filtered = _currentFirewallRules
                .Where(r => r.Direction.Equals(direction, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(r => r.IsActive && r.Action.Equals("Block", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(r => r.IsActive)
                .ThenBy(r => r.Name)
                .ToList();

            foreach (var rule in filtered)
                FirewallRules.Add(rule);

            FirewallRulesHeader = $"Firewall Rules — {_currentFirewallProfile} ({direction}) — {filtered.Count} rules";
        }

        private void OnFirewallRuleSelected()
        {
            var rule = SelectedFirewallRule;
            if (rule == null)
            {
                DetailRegistryPath = "";
                DetailValueText = "";
                return;
            }

            DetailRegistryPath = rule.RegistryPath;

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Rule Name:     {rule.Name}");
            if (!string.IsNullOrEmpty(rule.Description))
                sb.AppendLine($"Description:   {rule.Description}");
            sb.AppendLine($"Action:        {rule.Action}");
            sb.AppendLine($"Direction:     {rule.Direction}");
            sb.AppendLine($"Active:        {(rule.IsActive ? "Yes" : "No")}");
            if (!string.IsNullOrEmpty(rule.Protocol))
                sb.AppendLine($"Protocol:      {rule.Protocol}");
            if (!string.IsNullOrEmpty(rule.Profiles))
                sb.AppendLine($"Profiles:      {rule.Profiles}");
            if (!string.IsNullOrEmpty(rule.LocalPorts))
                sb.AppendLine($"Local Ports:   {rule.LocalPorts}");
            if (!string.IsNullOrEmpty(rule.RemotePorts))
                sb.AppendLine($"Remote Ports:  {rule.RemotePorts}");
            if (!string.IsNullOrEmpty(rule.LocalAddresses))
                sb.AppendLine($"Local Addr:    {rule.LocalAddresses}");
            if (!string.IsNullOrEmpty(rule.RemoteAddresses))
                sb.AppendLine($"Remote Addr:   {rule.RemoteAddresses}");
            if (!string.IsNullOrEmpty(rule.Application))
                sb.AppendLine($"Application:   {rule.Application}");
            if (!string.IsNullOrEmpty(rule.Service))
                sb.AppendLine($"Service:       {rule.Service}");
            if (!string.IsNullOrEmpty(rule.PackageFamilyName))
                sb.AppendLine($"Package:       {rule.PackageFamilyName}");
            if (!string.IsNullOrEmpty(rule.EmbedContext))
                sb.AppendLine($"Context:       {rule.EmbedContext}");
            sb.AppendLine();
            sb.AppendLine($"Raw Data:      {rule.RawData}");

            DetailValueText = sb.ToString();
        }

        private void DisplayDeviceManager(AnalysisSection section)
        {
            CurrentMode = ContentMode.DeviceManager;

            DeviceManagerNodes.Clear();
            DeviceProperties.Clear();
            DriverProperties.Clear();
            BuildDeviceDetailTabs();

            foreach (var classItem in section.Items)
            {
                bool isUnknownClass = classItem.Value == "N/A";
                var deviceClassData = new DeviceClassItem
                {
                    ClassName = classItem.Name,
                    ClassGuid = classItem.Value ?? "",
                    RegistryPath = classItem.RegistryPath ?? "",
                    Devices = new List<DeviceItem>()
                };

                var classNode = new DeviceTreeNode
                {
                    DisplayName = classItem.Name,
                    Tag = deviceClassData,
                    ForegroundBrushKey = isUnknownClass ? "ErrorBrush" : ""
                };

                if (classItem.SubItems != null)
                {
                    foreach (var deviceAnalysis in classItem.SubItems)
                    {
                        var deviceData = new DeviceItem
                        {
                            DisplayName = deviceAnalysis.Name,
                            RegistryPath = deviceAnalysis.RegistryPath ?? ""
                        };

                        if (deviceAnalysis.SubItems != null)
                        {
                            foreach (var prop in deviceAnalysis.SubItems)
                            {
                                // Check for driver details marker
                                if (prop.Name == "__DriverDetails__" && prop.SubItems != null)
                                {
                                    deviceData.DriverRegistryPath = prop.RegistryPath ?? "";
                                    foreach (var driverProp in prop.SubItems)
                                    {
                                        deviceData.DriverProperties.Add(new DevicePropertyItem
                                        {
                                            Name = driverProp.Name,
                                            Value = driverProp.Value ?? "",
                                            RegistryValueName = driverProp.RegistryValue ?? driverProp.Name,
                                            RegistryPath = driverProp.RegistryPath ?? prop.RegistryPath ?? ""
                                        });
                                    }
                                }
                                else
                                {
                                    deviceData.Properties.Add(new DevicePropertyItem
                                    {
                                        Name = prop.Name,
                                        Value = prop.Value ?? "",
                                        RegistryValueName = prop.RegistryValue ?? prop.Name,
                                        RegistryPath = prop.RegistryPath ?? deviceAnalysis.RegistryPath ?? ""
                                    });
                                }
                            }
                        }

                        deviceClassData.Devices.Add(deviceData);

                        // Determine color for device node
                        var status = deviceData.Properties.FirstOrDefault(p => p.Name == "Status")?.Value;
                        string brushKey = "";
                        if (status == "Disabled")
                            brushKey = "TextDisabledBrush";
                        else if (isUnknownClass)
                            brushKey = "ErrorBrush";

                        var deviceNode = new DeviceTreeNode
                        {
                            DisplayName = deviceData.DisplayName,
                            Tag = deviceData,
                            ForegroundBrushKey = brushKey
                        };

                        classNode.Children.Add(deviceNode);
                    }
                }

                DeviceManagerNodes.Add(classNode);
            }

            // Auto-select first node
            if (DeviceManagerNodes.Count > 0)
                DeviceManagerNodes[0].IsSelected = true;
        }

        private void BuildDeviceDetailTabs()
        {
            DeviceDetailTabs.Clear();

            var propsTab = new SubcategoryItem
            {
                Title = "Device Properties",
                IsEnabled = true,
                IsActive = true,
                SelectCommand = new RelayCommand(() => SwitchDeviceDetailTab("Properties"))
            };
            var driverTab = new SubcategoryItem
            {
                Title = "Driver Details",
                IsEnabled = true,
                IsActive = false,
                SelectCommand = new RelayCommand(() => SwitchDeviceDetailTab("Driver"))
            };

            DeviceDetailTabs.Add(propsTab);
            DeviceDetailTabs.Add(driverTab);

            SwitchDeviceDetailTab("Properties");
        }

        private void SwitchDeviceDetailTab(string tab)
        {
            IsDevicePropertiesTabActive = tab == "Properties";
            IsDriverDetailsTabActive = tab == "Driver";

            foreach (var t in DeviceDetailTabs)
                t.IsActive = (tab == "Properties" && t.Title == "Device Properties")
                          || (tab == "Driver" && t.Title == "Driver Details");
        }

        private void OnDeviceNodeSelected()
        {
            var node = SelectedDeviceNode;
            if (node == null) return;

            DeviceProperties.Clear();
            DriverProperties.Clear();

            if (node.Tag is DeviceItem device)
            {
                DeviceDetailHeader = $"Details — {device.DisplayName}";

                foreach (var prop in device.Properties)
                    DeviceProperties.Add(prop);

                if (device.DriverProperties.Count > 0)
                {
                    foreach (var prop in device.DriverProperties)
                        DriverProperties.Add(prop);
                    ShowDriverTab = true;
                }
                else
                {
                    ShowDriverTab = false;
                }

                // Update driver tab enabled state and reset to properties tab
                if (DeviceDetailTabs.Count >= 2)
                    DeviceDetailTabs[1].IsEnabled = ShowDriverTab;
                SwitchDeviceDetailTab("Properties");

                DetailRegistryPath = device.RegistryPath;
                DetailValueText = "Select a property to view registry details";
            }
            else if (node.Tag is DeviceClassItem classItem)
            {
                DeviceDetailHeader = $"{classItem.ClassName} — {classItem.Devices.Count} device(s)";
                ShowDriverTab = false;
                if (DeviceDetailTabs.Count >= 2)
                    DeviceDetailTabs[1].IsEnabled = false;
                SwitchDeviceDetailTab("Properties");

                // Show class overview: device list with status
                foreach (var dev in classItem.Devices)
                {
                    var status = dev.Properties.FirstOrDefault(p => p.Name == "Status")?.Value ?? "Unknown";
                    DeviceProperties.Add(new DevicePropertyItem
                    {
                        Name = dev.DisplayName,
                        Value = status,
                        RegistryPath = dev.RegistryPath
                    });
                }

                DetailRegistryPath = classItem.RegistryPath;
                DetailValueText = $"Class = {classItem.ClassName} | ClassGUID = {classItem.ClassGuid}";
            }
        }

        private void OnDevicePropertySelected()
        {
            var prop = SelectedDeviceProperty;
            if (prop == null) return;

            DetailRegistryPath = prop.RegistryPath;
            DetailValueText = $"{prop.Name} = {prop.Value} | Registry Value: {prop.RegistryValueName}";
        }

        private void OnDriverPropertySelected()
        {
            var prop = SelectedDriverProperty;
            if (prop == null) return;

            DetailRegistryPath = prop.RegistryPath;
            DetailValueText = $"{prop.Name} = {prop.Value} | Registry Value: {prop.RegistryValueName}";
        }

        private void DisplayScheduledTasks(AnalysisSection section)
        {
            CurrentMode = ContentMode.ScheduledTasks;

            ScheduledTaskNodes.Clear();
            ScheduledTaskDetails.Clear();

            int totalTasks = 0;

            // Recursive tree builder
            void AddTaskNodes(List<AnalysisItem> sourceItems, ObservableCollection<ScheduledTaskTreeNode> targetNodes)
            {
                var regularNodes = new List<ScheduledTaskTreeNode>();
                var disabledFolderNodes = new List<ScheduledTaskTreeNode>();

                foreach (var item in sourceItems)
                {
                    if (item.IsSubSection && item.SubItems != null)
                    {
                        // Folder node — first build children
                        var folderNode = new ScheduledTaskTreeNode
                        {
                            DisplayName = item.Name,
                            Tag = item,
                            IsFolder = true,
                            ForegroundBrushKey = ""
                        };
                        AddTaskNodes(item.SubItems, folderNode.Children);

                        // Check if folder has any ready tasks
                        bool hasReady = HasReadyTask(folderNode);
                        if (!hasReady)
                        {
                            // Create dimmed copy
                            var dimmedFolder = new ScheduledTaskTreeNode
                            {
                                DisplayName = item.Name,
                                Tag = item,
                                IsFolder = true,
                                ForegroundBrushKey = "TextDisabledBrush"
                            };
                            foreach (var c in folderNode.Children)
                                dimmedFolder.Children.Add(c);
                            disabledFolderNodes.Add(dimmedFolder);
                        }
                        else
                        {
                            regularNodes.Add(folderNode);
                        }
                    }
                    else
                    {
                        // Task leaf
                        totalTasks++;
                        bool isDisabled = item.Name.EndsWith(" [Disabled]", StringComparison.Ordinal);
                        regularNodes.Add(new ScheduledTaskTreeNode
                        {
                            DisplayName = item.Name,
                            Tag = item,
                            IsFolder = false,
                            ForegroundBrushKey = isDisabled ? "TextDisabledBrush" : ""
                        });
                    }
                }

                foreach (var n in regularNodes) targetNodes.Add(n);
                foreach (var n in disabledFolderNodes) targetNodes.Add(n);
            }

            AddTaskNodes(section.Items, ScheduledTaskNodes);

            ScheduledTasksHeader = $"Task Folders ({totalTasks} tasks)";

            // Auto-select first node
            if (ScheduledTaskNodes.Count > 0)
                ScheduledTaskNodes[0].IsSelected = true;
        }

        private static bool HasReadyTask(ScheduledTaskTreeNode folderNode)
        {
            foreach (var child in folderNode.Children)
            {
                if (child.Children.Count > 0)
                {
                    if (HasReadyTask(child)) return true;
                }
                else
                {
                    if (!child.DisplayName.EndsWith(" [Disabled]", StringComparison.Ordinal))
                        return true;
                }
            }
            return false;
        }

        private void OnScheduledTaskNodeSelected()
        {
            var node = SelectedScheduledTaskNode;
            if (node?.Tag is not AnalysisItem item) return;

            ScheduledTaskDetails.Clear();

            if (node.IsFolder)
            {
                // Folder summary
                ScheduledTaskDetailHeader = $"Folder: {item.Name}";

                int CountTasks(AnalysisItem folder)
                {
                    int count = 0;
                    if (folder.SubItems == null) return 0;
                    foreach (var sub in folder.SubItems)
                    {
                        if (sub.IsSubSection) count += CountTasks(sub);
                        else count++;
                    }
                    return count;
                }

                int folderTaskCount = CountTasks(item);
                int subFolderCount = item.SubItems?.Count(i => i.IsSubSection) ?? 0;
                int directTaskCount = item.SubItems?.Count(i => !i.IsSubSection) ?? 0;

                ScheduledTaskDetails.Add(new AnalyzeGridRow { Column1 = "Total Tasks", Column2 = folderTaskCount.ToString() });
                ScheduledTaskDetails.Add(new AnalyzeGridRow { Column1 = "Sub-folders", Column2 = subFolderCount.ToString() });
                ScheduledTaskDetails.Add(new AnalyzeGridRow { Column1 = "Direct Tasks", Column2 = directTaskCount.ToString() });
                ScheduledTaskDetails.Add(new AnalyzeGridRow { Column1 = "Registry Path", Column2 = item.RegistryPath ?? "" });

                DetailRegistryPath = item.RegistryPath ?? "";
                DetailValueText = $"Folder: {item.Name} ({folderTaskCount} tasks)";
            }
            else
            {
                // Task leaf — show properties
                ScheduledTaskDetailHeader = $"Task: {item.Name}";

                // GUID as first row
                ScheduledTaskDetails.Add(new AnalyzeGridRow { Column1 = "GUID", Column2 = item.Value ?? "" });

                if (item.SubItems != null)
                {
                    foreach (var prop in item.SubItems)
                    {
                        ScheduledTaskDetails.Add(new AnalyzeGridRow { Column1 = prop.Name, Column2 = prop.Value ?? "" });
                    }
                }

                DetailRegistryPath = item.RegistryPath ?? "";
                DetailValueText = $"Task: {item.Name}";
            }
        }

        private void OnScheduledTaskDetailSelected()
        {
            var row = SelectedScheduledTaskDetail;
            if (row == null) return;
            DetailValueText = $"{row.Column1}: {row.Column2}";
        }

        private void DisplayCertificateStores(AnalysisSection section)
        {
            CurrentMode = ContentMode.CertificateStores;

            CertStoreNodes.Clear();
            CertificateDetails.Clear();

            if (section.Tag is List<CertificateStoreInfo> storeList)
            {
                int totalCerts = storeList.Sum(s => s.Certificates.Count);
                CertStoresTreeHeader = $"Certificate Stores ({storeList.Count} stores, {totalCerts} certificates)";

                foreach (var store in storeList)
                {
                    bool isEmpty = store.Certificates.Count == 0;
                    var storeNode = new ScheduledTaskTreeNode
                    {
                        DisplayName = $"{store.FriendlyName} ({store.Certificates.Count})",
                        Tag = store,
                        IsFolder = true,
                        ForegroundBrushKey = isEmpty ? "TextDisabledBrush" : ""
                    };

                    foreach (var cert in store.Certificates.OrderBy(c => c.DisplayName, StringComparer.OrdinalIgnoreCase))
                    {
                        bool isExpired = cert.ValidTo.HasValue && cert.ValidTo.Value < DateTime.Now;
                        storeNode.Children.Add(new ScheduledTaskTreeNode
                        {
                            DisplayName = cert.DisplayName,
                            Tag = cert,
                            IsFolder = false,
                            ForegroundBrushKey = isExpired ? "ErrorBrush" : ""
                        });
                    }

                    CertStoreNodes.Add(storeNode);
                }
            }

            // Auto-select first store
            if (CertStoreNodes.Count > 0)
                CertStoreNodes[0].IsSelected = true;
        }

        private void OnCertStoreNodeSelected()
        {
            var node = SelectedCertStoreNode;
            if (node == null) return;

            CertificateDetails.Clear();

            if (node.Tag is CertificateStoreInfo store)
            {
                CertStoreDetailHeader = $"Store: {store.FriendlyName}";

                CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Store Name", Column2 = store.FriendlyName });
                CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Registry Name", Column2 = store.RegistryName });
                CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Certificate Count", Column2 = store.Certificates.Count.ToString() });
                CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Registry Path", Column2 = store.RegistryPath });

                DetailRegistryPath = store.RegistryPath;
                DetailValueText = $"Store: {store.FriendlyName} ({store.Certificates.Count} certificates)";
            }
            else if (node.Tag is CertificateInfo cert)
            {
                CertStoreDetailHeader = $"Certificate: {cert.DisplayName}";

                CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Display Name", Column2 = cert.DisplayName });
                if (!string.IsNullOrEmpty(cert.Subject))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Subject", Column2 = cert.Subject });
                if (!string.IsNullOrEmpty(cert.Issuer))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Issuer", Column2 = cert.Issuer });
                if (!string.IsNullOrEmpty(cert.SerialNumber))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Serial Number", Column2 = cert.SerialNumber });
                if (cert.ValidFrom.HasValue)
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Valid From", Column2 = cert.ValidFrom.Value.ToString("yyyy-MM-dd HH:mm:ss") });
                if (cert.ValidTo.HasValue)
                {
                    bool isExpired = cert.ValidTo.Value < DateTime.Now;
                    CertificateDetails.Add(new AnalyzeGridRow
                    {
                        Column1 = "Valid To",
                        Column2 = cert.ValidTo.Value.ToString("yyyy-MM-dd HH:mm:ss") + (isExpired ? " (EXPIRED)" : ""),
                        IsWarning = isExpired
                    });
                }
                if (!string.IsNullOrEmpty(cert.Thumbprint))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Thumbprint", Column2 = cert.Thumbprint });
                if (!string.IsNullOrEmpty(cert.Sha1Hash))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "SHA-1 Hash", Column2 = cert.Sha1Hash });
                if (!string.IsNullOrEmpty(cert.SignatureAlgorithm))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Signature Algorithm", Column2 = cert.SignatureAlgorithm });
                if (!string.IsNullOrEmpty(cert.FriendlyName))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Friendly Name", Column2 = cert.FriendlyName });
                if (!string.IsNullOrEmpty(cert.KeyProvider))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Key Provider", Column2 = cert.KeyProvider });
                if (!string.IsNullOrEmpty(cert.KeyContainer))
                    CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Key Container", Column2 = cert.KeyContainer });
                CertificateDetails.Add(new AnalyzeGridRow { Column1 = "Registry Path", Column2 = cert.RegistryPath });

                DetailRegistryPath = cert.RegistryPath;
                DetailValueText = $"Certificate: {cert.DisplayName}";
            }
        }

        private void OnCertificateDetailSelected()
        {
            var row = SelectedCertificateDetail;
            if (row == null) return;
            DetailValueText = $"{row.Column1}: {row.Column2}";
        }

        // ── Group Policy scrollable document ────────────────────────────────

        private void DisplayGroupPolicy(AnalysisSection section)
        {
            CurrentMode = ContentMode.GroupPolicy;

            _allGpDocRows.Clear();
            GroupPolicyDocRows.Clear();
            GpSearchText = "";

            if (section.Tag is List<GroupPolicyCategory> categoryList)
            {
                int totalSettings = CountPolicySettings(categoryList);
                GroupPolicyHeader = $"Policies > Administrative Templates ({totalSettings} settings)";

                // Flatten the category tree into the cache
                FlattenGpCategories(categoryList, _allGpDocRows);

                // Copy to the bound collection
                foreach (var row in _allGpDocRows)
                    GroupPolicyDocRows.Add(row);
            }

            DetailValueText = "";
            DetailRegistryPath = "";
        }

        /// <summary>
        /// Recursively flattens the GroupPolicyCategory tree into a list of document rows.
        /// Each leaf category (one with Items) becomes a section header + policy rows.
        /// </summary>
        private void FlattenGpCategories(List<GroupPolicyCategory> categories,
            List<GpDocumentRow> rows, string parentPath = "")
        {
            foreach (var cat in categories.OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase))
            {
                string fullPath = string.IsNullOrEmpty(parentPath)
                    ? cat.Name
                    : $"{parentPath}/{cat.Name}";

                if (cat.Items.Count > 0)
                {
                    // Leaf category with items — emit section header + policy rows
                    rows.Add(new GpDocumentRow
                    {
                        IsSectionHeader = true,
                        CategoryPath = fullPath,
                        SettingCount = cat.Items.Count
                    });

                    foreach (var item in cat.Items)
                    {
                        string displayName;
                        if (string.IsNullOrEmpty(item.AdmxCategoryPath))
                        {
                            // Extra Registry Settings — show full path like GPResult:
                            // Software\Policies\Microsoft\...\ValueName
                            displayName = $@"Software\{item.RegistryPath}\{item.Name}";
                        }
                        else if (item.IsListChild)
                        {
                            // enabledList child — show the value name (e.g., ".bak", "%windir%\spool")
                            displayName = item.Name;
                        }
                        else
                        {
                            displayName = !string.IsNullOrEmpty(item.FriendlyName)
                                ? item.FriendlyName
                                : item.Name;
                        }

                        string displayValue = !string.IsNullOrEmpty(item.FriendlyValue)
                            ? item.FriendlyValue
                            : item.Value;

                        rows.Add(new GpDocumentRow
                        {
                            IsSectionHeader = false,
                            IsListChild = item.IsListChild,
                            PolicyName = displayName,
                            SettingValue = displayValue,
                            PolicyItem = item
                        });
                    }
                }

                // Recurse into subcategories
                if (cat.SubCategories.Count > 0)
                    FlattenGpCategories(cat.SubCategories, rows, fullPath);
            }
        }

        /// <summary>
        /// Filters the GP document rows based on the search text.
        /// Section headers are included only if at least one of their policy rows matches.
        /// </summary>
        private void FilterGpDocRows()
        {
            GroupPolicyDocRows.Clear();
            var search = _gpSearchText.Trim();

            if (string.IsNullOrEmpty(search))
            {
                // No filter — show everything
                foreach (var row in _allGpDocRows)
                    GroupPolicyDocRows.Add(row);

                // Restore original header
                int total = _allGpDocRows.Count(r => !r.IsSectionHeader);
                GroupPolicyHeader = $"Policies > Administrative Templates ({total} settings)";
                return;
            }

            // Walk the flat list: for each section header, collect its policy rows,
            // check if any match, and if so include the header + matching rows.
            GpDocumentRow? currentHeader = null;
            var pendingRows = new List<GpDocumentRow>();
            int matchCount = 0;

            foreach (var row in _allGpDocRows)
            {
                if (row.IsSectionHeader)
                {
                    // Flush previous section if it had matches
                    if (currentHeader != null && pendingRows.Count > 0)
                    {
                        GroupPolicyDocRows.Add(new GpDocumentRow
                        {
                            IsSectionHeader = true,
                            CategoryPath = currentHeader.CategoryPath,
                            SettingCount = pendingRows.Count
                        });
                        foreach (var r in pendingRows)
                            GroupPolicyDocRows.Add(r);
                    }

                    currentHeader = row;
                    pendingRows.Clear();
                }
                else
                {
                    // Check if this policy row matches the search
                    if (MatchesGpSearch(row, search))
                    {
                        pendingRows.Add(row);
                        matchCount++;
                    }
                }
            }

            // Flush last section
            if (currentHeader != null && pendingRows.Count > 0)
            {
                GroupPolicyDocRows.Add(new GpDocumentRow
                {
                    IsSectionHeader = true,
                    CategoryPath = currentHeader.CategoryPath,
                    SettingCount = pendingRows.Count
                });
                foreach (var r in pendingRows)
                    GroupPolicyDocRows.Add(r);
            }

            GroupPolicyHeader = $"Policies > Administrative Templates ({matchCount} of {_allGpDocRows.Count(r => !r.IsSectionHeader)} settings)";
        }

        private static bool MatchesGpSearch(GpDocumentRow row, string search)
        {
            if (row.PolicyName.Contains(search, StringComparison.OrdinalIgnoreCase))
                return true;
            if (row.SettingValue.Contains(search, StringComparison.OrdinalIgnoreCase))
                return true;

            var item = row.PolicyItem;
            if (item == null) return false;

            if (item.RegistryPath.Contains(search, StringComparison.OrdinalIgnoreCase))
                return true;
            if (item.RegistryValueName.Contains(search, StringComparison.OrdinalIgnoreCase))
                return true;
            if (!string.IsNullOrEmpty(item.Description) &&
                item.Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static int CountPolicySettings(List<GroupPolicyCategory> categories)
        {
            int count = 0;
            foreach (var cat in categories)
            {
                count += cat.Items.Count;
                count += cat.SubCategories.Sum(sc => CountPolicySettingsRecursive(sc));
            }
            return count;
        }

        private static int CountPolicySettingsRecursive(GroupPolicyCategory category)
        {
            int count = category.Items.Count;
            foreach (var sub in category.SubCategories)
                count += CountPolicySettingsRecursive(sub);
            return count;
        }

        private void OnGroupPolicyRowSelected()
        {
            var row = SelectedGroupPolicyRow;
            if (row == null || row.IsSectionHeader) return;

            var item = row.PolicyItem;
            if (item == null) return;

            // Show registry path in the path bar
            DetailRegistryPath = item.RegistryPath;

            // Build rich detail text for the bottom panel
            var sb = new System.Text.StringBuilder();

            // Policy name
            string policyName = !string.IsNullOrEmpty(item.FriendlyName)
                ? item.FriendlyName
                : item.Name;
            sb.AppendLine(policyName);
            sb.AppendLine();

            // ADMX description (can be multi-paragraph)
            // Strip leading policy name from description to avoid duplication
            if (!string.IsNullOrEmpty(item.Description))
            {
                var desc = item.Description;
                if (desc.StartsWith(policyName, StringComparison.OrdinalIgnoreCase))
                {
                    desc = desc.Substring(policyName.Length).TrimStart('\r', '\n', ' ', '.', '-', ':');
                }
                if (!string.IsNullOrWhiteSpace(desc))
                {
                    sb.AppendLine(desc);
                    sb.AppendLine();
                }
            }

            // Registry details
            sb.AppendLine($"Registry Path: {item.RegistryPath}");
            sb.AppendLine($"Value Name: {item.RegistryValueName}");
            if (!string.IsNullOrEmpty(item.ValueType))
                sb.AppendLine($"Value Type: {item.ValueType}");
            sb.AppendLine($"Raw Value: {item.Value}");

            if (!string.IsNullOrEmpty(item.FriendlyValue))
                sb.AppendLine($"Interpreted: {item.FriendlyValue}");

            if (!string.IsNullOrEmpty(item.SupportedOn))
                sb.AppendLine($"Supported On: {item.SupportedOn}");

            DetailValueText = sb.ToString().TrimEnd();
        }

        private void DisplayRolesFeatures()
        {
            CurrentMode = ContentMode.RolesFeatures;

            RolesFeatureNodes.Clear();
            RoleFeatureDetails.Clear();

            List<RoleFeatureItem> allItems;
            try
            {
                allItems = _activeExtractor.GetRolesAndFeaturesData();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetRolesAndFeaturesData error: {ex.Message}");
                allItems = new List<RoleFeatureItem>();
            }

            if (allItems.Count == 0)
            {
                RolesTreeHeader = "Roles & Features (none found)";
                return;
            }

            // Build parent-child lookup
            var childrenOf = new Dictionary<string, List<RoleFeatureItem>>(StringComparer.OrdinalIgnoreCase);
            var topLevel = new List<RoleFeatureItem>();

            foreach (var item in allItems)
            {
                if (string.IsNullOrEmpty(item.ParentName))
                {
                    topLevel.Add(item);
                }
                else
                {
                    if (!childrenOf.ContainsKey(item.ParentName))
                        childrenOf[item.ParentName] = new List<RoleFeatureItem>();
                    childrenOf[item.ParentName].Add(item);
                }
            }

            // Sort: installed first, then alphabetically
            Comparison<RoleFeatureItem> sortComparison = (a, b) =>
            {
                int aInstalled = a.InstallState == 1 ? 1 : 0;
                int bInstalled = b.InstallState == 1 ? 1 : 0;
                int installCompare = bInstalled.CompareTo(aInstalled);
                if (installCompare != 0) return installCompare;
                return string.Compare(a.DisplayName, b.DisplayName, StringComparison.OrdinalIgnoreCase);
            };

            topLevel.Sort(sortComparison);

            // Recursive tree builder
            void AddChildren(DeviceTreeNode parentNode, string parentKeyName)
            {
                if (!childrenOf.TryGetValue(parentKeyName, out var children)) return;
                children.Sort(sortComparison);
                foreach (var child in children)
                {
                    var childNode = new DeviceTreeNode
                    {
                        DisplayName = child.DisplayName,
                        Tag = child,
                        ForegroundBrushKey = child.InstallState != 1 ? "TextDisabledBrush" : ""
                    };
                    parentNode.Children.Add(childNode);
                    AddChildren(childNode, child.KeyName);
                }
            }

            int installedCount = allItems.Count(i => i.InstallState == 1);

            foreach (var item in topLevel)
            {
                var node = new DeviceTreeNode
                {
                    DisplayName = item.DisplayName,
                    Tag = item,
                    ForegroundBrushKey = item.InstallState != 1 ? "TextDisabledBrush" : ""
                };
                AddChildren(node, item.KeyName);
                RolesFeatureNodes.Add(node);
            }

            RolesTreeHeader = $"Roles & Features ({installedCount} installed / {allItems.Count} total)";

            // Auto-select first node
            if (RolesFeatureNodes.Count > 0)
                RolesFeatureNodes[0].IsSelected = true;
        }

        private void OnRolesFeatureNodeSelected()
        {
            var node = SelectedRolesFeatureNode;
            if (node?.Tag is not RoleFeatureItem roleItem) return;

            RoleFeatureDetails.Clear();
            RoleFeatureDetailHeader = $"Details — {roleItem.DisplayName}";

            RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Display Name", Column2 = roleItem.DisplayName });
            RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Internal Name", Column2 = roleItem.KeyName });
            RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Type", Column2 = roleItem.ComponentTypeName });
            RoleFeatureDetails.Add(new AnalyzeGridRow
            {
                Column1 = "Install State",
                Column2 = roleItem.InstallStateName,
                IsHealthy = roleItem.InstallState == 1
            });

            if (!string.IsNullOrEmpty(roleItem.Description))
                RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Description", Column2 = roleItem.Description });

            if (!string.IsNullOrEmpty(roleItem.ParentName))
                RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Parent", Column2 = roleItem.ParentName });

            if (roleItem.MajorVersion > 0 || roleItem.MinorVersion > 0)
                RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Version", Column2 = $"{roleItem.MajorVersion}.{roleItem.MinorVersion}" });

            if (!string.IsNullOrEmpty(roleItem.SystemServices))
                RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "System Services", Column2 = roleItem.SystemServices });

            if (!string.IsNullOrEmpty(roleItem.Dependencies))
                RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Dependencies", Column2 = roleItem.Dependencies });

            RoleFeatureDetails.Add(new AnalyzeGridRow { Column1 = "Numeric ID", Column2 = roleItem.NumericId.ToString() });

            DetailRegistryPath = roleItem.RegistryPath;
            DetailValueText = $"{roleItem.DisplayName} ({roleItem.ComponentTypeName}) | {roleItem.InstallStateName}";
        }

        private void OnRoleFeatureDetailSelected()
        {
            var row = SelectedRoleFeatureDetail;
            if (row == null) return;

            // Keep the registry path from the selected tree node
            if (SelectedRolesFeatureNode?.Tag is RoleFeatureItem roleItem)
                DetailRegistryPath = roleItem.RegistryPath;

            DetailValueText = $"{row.Column1} = {row.Column2}";
        }

        private void DisplayDiskPartitions()
        {
            CurrentMode = ContentMode.DiskPartitions;

            MountedDevices.Clear();
            MountedDeviceDetails.Clear();

            List<MountedDeviceEntry> mountedDevices;
            try
            {
                mountedDevices = _activeExtractor.GetMountedDevices();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetMountedDevices error: {ex.Message}");
                mountedDevices = new List<MountedDeviceEntry>();
            }

            foreach (var device in mountedDevices)
            {
                // Pre-compute display-friendly FriendlyName for the grid column
                var displayName = !string.IsNullOrEmpty(device.FriendlyName)
                    ? device.FriendlyName
                    : !string.IsNullOrEmpty(device.DeviceClass)
                        ? device.DeviceClass
                        : "";
                if (device.StaleStatus == "Stale")
                    displayName = !string.IsNullOrEmpty(displayName) ? $"{displayName} [Stale]" : "[Stale]";
                device.FriendlyName = displayName;

                MountedDevices.Add(device);
            }

            // Auto-select first device
            if (MountedDevices.Count > 0)
                SelectedMountedDevice = MountedDevices[0];
        }

        private void OnMountedDeviceSelected()
        {
            var device = SelectedMountedDevice;
            if (device == null)
            {
                MountedDeviceDetails.Clear();
                DetailRegistryPath = "";
                DetailValueText = "";
                return;
            }

            MountedDeviceDetails.Clear();
            MountedDeviceDetailHeader = $"Details — {device.MountPoint}";

            // Update bottom detail pane
            var pathText = $"Registry Path: {device.RegistryPath}";
            if (!string.IsNullOrEmpty(device.EnumPath))
                pathText += $"  |  Enum: {device.EnumPath}";
            DetailRegistryPath = pathText;
            DetailValueText = $"Mount: {device.MountPoint} | Style: {device.PartitionStyle} | {device.Identifier}";

            // Helper to add non-empty rows (section headers always added)
            void AddRow(string prop, string val, bool isSubSection = false, bool isWarning = false)
            {
                if (isSubSection || !string.IsNullOrEmpty(val))
                    MountedDeviceDetails.Add(new AnalyzeGridRow { Column1 = prop, Column2 = val, IsSubSection = isSubSection, IsWarning = isWarning });
            }

            // Mount Info
            AddRow("Value Name", device.RegistryValueName);
            AddRow("Mount Point", device.MountPoint);
            AddRow("Type", device.MountType);

            // Partition
            AddRow("Partition Style", device.PartitionStyle);
            AddRow("Disk Signature", device.DiskSignature);
            AddRow("Partition Offset", device.PartitionOffset);
            AddRow("Partition GUID", device.PartitionGuid);
            AddRow("Disk ID", device.DiskId);

            // Device Path
            AddRow("Bus Type", device.BusType);
            AddRow("Vendor", device.Vendor);
            AddRow("Product", device.Product);
            AddRow("Serial", device.Serial);
            AddRow("Device Path", device.DevicePath);

            // Enum Device Info
            AddRow("Friendly Name", device.FriendlyName);
            AddRow("Device Class", device.DeviceClass);
            AddRow("Service", device.DeviceService);
            AddRow("Manufacturer", device.Manufacturer);
            AddRow("Location", device.LocationInfo);
            AddRow("Status", device.DeviceStatus);
            AddRow("Enum Path", device.EnumPath);

            // Disk Status
            if (!string.IsNullOrEmpty(device.StaleStatus))
            {
                var statusDesc = device.StaleStatus switch
                {
                    "Active" => "Active — disk has current STORAGE\\Volume registrations",
                    "Stale" => "Stale — no active STORAGE\\Volume registrations found (disk may have been detached)",
                    "Unknown" => "Unknown — could not determine parent disk for this partition",
                    _ => device.StaleStatus
                };
                bool isStaleWarning = device.StaleStatus == "Stale";
                MountedDeviceDetails.Add(new AnalyzeGridRow { Column1 = "Volume Status", Column2 = statusDesc, IsWarning = isStaleWarning });
            }

            // Raw Data
            AddRow("Data Length", $"{device.DataLength} bytes");
        }

        private void OnMountedDeviceDetailSelected()
        {
            var row = SelectedMountedDeviceDetail;
            if (row == null || row.IsSubSection) return;
            DetailValueText = $"{row.Column1}: {row.Column2}";
        }

        private void DisplayPhysicalDisks()
        {
            CurrentMode = ContentMode.PhysicalDisks;

            PhysicalDisksList.Clear();
            PhysicalDiskDetails.Clear();

            List<PhysicalDiskEntry> physicalDisks;
            try
            {
                physicalDisks = _activeExtractor.GetPhysicalDisks();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetPhysicalDisks error: {ex.Message}");
                physicalDisks = new List<PhysicalDiskEntry>();
            }

            foreach (var disk in physicalDisks)
            {
                // Pre-compute display-friendly values for grid columns
                if (string.IsNullOrEmpty(disk.FriendlyName))
                    disk.FriendlyName = disk.DeviceId;
                if (string.IsNullOrEmpty(disk.PoolStatus))
                    disk.PoolStatus = disk.DeviceStatus;

                PhysicalDisksList.Add(disk);
            }

            // Auto-select first disk
            if (PhysicalDisksList.Count > 0)
                SelectedPhysicalDisk = PhysicalDisksList[0];
        }

        private void OnPhysicalDiskSelected()
        {
            var disk = SelectedPhysicalDisk;
            if (disk == null)
            {
                PhysicalDiskDetails.Clear();
                DetailRegistryPath = "";
                DetailValueText = "";
                return;
            }

            PhysicalDiskDetails.Clear();
            PhysicalDiskDetailHeader = $"Details — {disk.FriendlyName}";

            // Update bottom pane
            DetailRegistryPath = $"Registry Path: {disk.EnumPath}";
            DetailValueText = $"Disk: {disk.FriendlyName} | Bus: {disk.BusType} | {disk.PoolStatus}";

            // Helper to add non-empty rows (section headers always added)
            void AddRow(string prop, string val, bool isSubSection = false)
            {
                if (isSubSection || !string.IsNullOrEmpty(val))
                    PhysicalDiskDetails.Add(new AnalyzeGridRow { Column1 = prop, Column2 = val, IsSubSection = isSubSection });
            }

            // Identity
            AddRow("Friendly Name", disk.FriendlyName);
            AddRow("Description", disk.DeviceDesc);
            AddRow("Device ID", disk.DeviceId);

            // Bus/Location
            AddRow("Bus Type", disk.BusType);
            AddRow("Location", disk.LocationInfo);

            // Hardware
            AddRow("Hardware ID", disk.HardwareId);
            AddRow("Manufacturer", disk.Manufacturer);
            AddRow("Service", disk.Service);
            AddRow("Status", disk.DeviceStatus);

            // Partmgr / Storage
            AddRow("Disk ID", disk.DiskId);
            if (!string.IsNullOrEmpty(disk.DiskId))
                AddRow("Volume Count", disk.VolumeCount.ToString());
            AddRow("Drive Letters", disk.DriveLetters);

            // Pool detection
            if (disk.PoolStatus == "Probable Pool Member")
            {
                PhysicalDiskDetails.Add(new AnalyzeGridRow
                {
                    Column1 = "Storage Pool",
                    Column2 = "Probable Pool Member — this disk has a DiskId but no STORAGE\\Volume entries, indicating it may be claimed by Storage Spaces",
                    IsWarning = true
                });
            }
            else if (!string.IsNullOrEmpty(disk.PoolStatus))
            {
                AddRow("Storage Pool", disk.PoolStatus);
            }

            // Registry
            AddRow("Enum Path", disk.EnumPath);
        }

        private void OnPhysicalDiskDetailSelected()
        {
            var row = SelectedPhysicalDiskDetail;
            if (row == null || row.IsSubSection) return;
            DetailValueText = $"{row.Column1}: {row.Column2}";
        }

        private void DisplayAppxPackages(AnalysisSection section)
        {
            // Show overview first
            DisplayDefaultGrid(section);

            // Determine which hive owns this section (SOFTWARE)
            var hive = _sectionHiveMap.GetValueOrDefault(section)
                       ?? _loadedHives.FirstOrDefault(h => h.HiveType == HiveType.SOFTWARE);
            if (hive == null) return;

            var ext = hive.InfoExtractor;
            var inboxPkgs = ext.GetAppxPackages("InBox");
            var userPkgs = ext.GetAppxPackages("UserInstalled");

            // Helper to display a filtered list of Appx packages
            void ShowAppxFilter(string filterKey, string tabTitle)
            {
                try
                {
                    var packages = ext.GetAppxPackages(filterKey);
                    var filtered = new AnalysisSection
                    {
                        Title = tabTitle
                    };
                    foreach (var pkg in packages)
                    {
                        filtered.Items.Add(new AnalysisItem
                        {
                            Name = pkg.PackageName,
                            IsSubSection = false,
                            SubItems = new List<AnalysisItem>
                            {
                                new() { Name = "Version", Value = pkg.Version },
                                new() { Name = "Architecture", Value = pkg.Architecture }
                            }
                        });
                    }

                    GridColumn1Header = "Package Name";
                    GridColumn2Header = "Version";
                    GridColumn3Header = "Architecture";
                    GridColumnCount = 3;
                    ContentHeader = tabTitle;

                    CurrentMode = ContentMode.DefaultGrid;
                    DefaultGridRows.Clear();
                    foreach (var item in filtered.Items)
                    {
                        DefaultGridRows.Add(new AnalyzeGridRow
                        {
                            Column1 = item.Name,
                            Column2 = item.SubItems?.ElementAtOrDefault(0)?.Value ?? "",
                            Column3 = item.SubItems?.ElementAtOrDefault(1)?.Value ?? "",
                            SourceItem = item
                        });
                    }

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Appx filter error: {ex.Message}");
                }
            }

            // Build sub-tabs
            SubTabs.Add(new SubcategoryItem
            {
                Title = $"\U0001f4e6 InBox Preinstalled ({inboxPkgs.Count})",
                IsEnabled = true,
                Tooltip = "Show InBox preinstalled packages",
                SelectCommand = new RelayCommand(() => ShowAppxFilter("InBox", $"\U0001f4e6 InBox Preinstalled ({inboxPkgs.Count})"))
            });
            SubTabs.Add(new SubcategoryItem
            {
                Title = $"\U0001f4f2 User Installed ({userPkgs.Count})",
                IsEnabled = true,
                Tooltip = "Show user-installed packages",
                SelectCommand = new RelayCommand(() => ShowAppxFilter("UserInstalled", $"\U0001f4f2 User Installed ({userPkgs.Count})"))
            });
            ShowSubTabs = true;

            // Auto-select InBox Preinstalled by default
            SubTabs[0].SelectCommand.Execute(null);
        }

        private void DisplayCbsPackages(AnalysisSection section)
        {
            // Determine which hive owns this section (SOFTWARE)
            var hive = _sectionHiveMap.GetValueOrDefault(section)
                       ?? _loadedHives.FirstOrDefault(h => h.HiveType == HiveType.SOFTWARE);
            if (hive == null)
            {
                CurrentMode = ContentMode.CbsPackages;
                return;
            }

            var ext = hive.InfoExtractor;

            // Helper: display All Packages in a 5-column grid with search/DISM filter
            void ShowAllPackages(string tabTitle)
            {
                try
                {
                    var packagesSections = ext.GetPackagesAnalysis();

                    // Cache all parsed rows for filtering
                    _allCbsPackagesData.Clear();

                    foreach (var sec in packagesSections.Where(s => !s.Title.Contains("Summary")))
                    {
                        var groupName = sec.Title;
                        if (groupName.StartsWith("\U0001f4e6 "))
                            groupName = groupName.Substring(3);
                        var parenIndex = groupName.LastIndexOf(" (");
                        if (parenIndex > 0)
                            groupName = groupName.Substring(0, parenIndex);

                        foreach (var item in sec.Items)
                        {
                            var valueParts = item.Value?.Split('|') ?? Array.Empty<string>();
                            var state = "";
                            var installed = "";
                            var user = "";
                            int visibility = 0;
                            foreach (var part in valueParts)
                            {
                                var trimmed = part.Trim();
                                if (trimmed.StartsWith("State:"))
                                    state = trimmed.Substring(6).Trim();
                                else if (trimmed.StartsWith("Installed:"))
                                    installed = trimmed.Substring(10).Trim();
                                else if (trimmed.StartsWith("User:"))
                                    user = trimmed.Substring(5).Trim();
                                else if (trimmed.StartsWith("Visibility:"))
                                    int.TryParse(trimmed.Substring(11).Trim(), out visibility);
                            }

                            _allCbsPackagesData.Add((groupName, item.Name, state, installed, user, visibility, item));
                        }
                    }

                    // Set up grid columns
                    GridColumn1Header = "Package Group";
                    GridColumn2Header = "Package Version";
                    GridColumn3Header = "State";
                    GridColumn4Header = "Installed";
                    GridColumn5Header = "User";
                    GridColumnCount = 5;
                    ContentHeader = tabTitle;
                    CurrentMode = ContentMode.DefaultGrid;

                    // Reset search state and apply filter
                    CbsSearchText = "";
                    CbsDismFilter = false;
                    ShowCbsSearch = true;
                    FilterCbsPackages();

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex) { Debug.WriteLine($"CBS AllPackages error: {ex.Message}"); }
            }

            // Helper: display Pending Sessions in a 3-column grid
            void ShowPendingSessions(string tabTitle)
            {
                try
                {
                    ShowCbsSearch = false;
                    var sessions = ext.GetCbsPendingSessionsAnalysis();
                    CurrentMode = ContentMode.DefaultGrid;
                    DefaultGridRows.Clear();

                    GridColumn1Header = "Session";
                    GridColumn2Header = "Status";
                    GridColumn3Header = "Details";
                    GridColumnCount = 3;
                    ContentHeader = tabTitle;

                    if (sessions.Count == 0 || sessions.All(s => s.Items.Count == 0))
                    {
                        DefaultGridRows.Add(new AnalyzeGridRow { Column1 = "No pending sessions found" });
                        UpdateSubTabActiveState(tabTitle);
                        return;
                    }

                    foreach (var sec in sessions)
                    {
                        foreach (var item in sec.Items)
                        {
                            DefaultGridRows.Add(new AnalyzeGridRow
                            {
                                Column1 = item.Name,
                                Column2 = item.Value,
                                Column3 = sec.Title,
                                SourceItem = item
                            });
                        }
                    }

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex) { Debug.WriteLine($"CBS PendingSessions error: {ex.Message}"); }
            }

            // Helper: display Pending Packages in a 3-column grid
            void ShowPendingPackages(string tabTitle)
            {
                try
                {
                    ShowCbsSearch = false;
                    var packages = ext.GetCbsPendingPackagesAnalysis();
                    CurrentMode = ContentMode.DefaultGrid;
                    DefaultGridRows.Clear();

                    GridColumn1Header = "Package Name";
                    GridColumn2Header = "Status";
                    GridColumn3Header = "Details";
                    GridColumnCount = 3;
                    ContentHeader = tabTitle;

                    if (packages.Count == 0 || packages.All(s => s.Items.Count == 0))
                    {
                        DefaultGridRows.Add(new AnalyzeGridRow { Column1 = "No pending packages found" });
                        UpdateSubTabActiveState(tabTitle);
                        return;
                    }

                    foreach (var sec in packages)
                    {
                        foreach (var item in sec.Items)
                        {
                            DefaultGridRows.Add(new AnalyzeGridRow
                            {
                                Column1 = item.Name,
                                Column2 = item.Value,
                                Column3 = sec.Title,
                                SourceItem = item
                            });
                        }
                    }

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex) { Debug.WriteLine($"CBS PendingPackages error: {ex.Message}"); }
            }

            // Helper: display Reboot Status in a 3-column grid
            void ShowRebootStatus(string tabTitle)
            {
                try
                {
                    ShowCbsSearch = false;
                    var status = ext.GetCbsRebootStatusAnalysis();
                    CurrentMode = ContentMode.DefaultGrid;
                    DefaultGridRows.Clear();

                    GridColumn1Header = "Property";
                    GridColumn2Header = "Value";
                    GridColumn3Header = "Details";
                    GridColumnCount = 3;
                    ContentHeader = tabTitle;

                    if (status.Count == 0 || status.All(s => s.Items.Count == 0))
                    {
                        DefaultGridRows.Add(new AnalyzeGridRow { Column1 = "No reboot status information found" });
                        UpdateSubTabActiveState(tabTitle);
                        return;
                    }

                    foreach (var sec in status)
                    {
                        foreach (var item in sec.Items)
                        {
                            var details = item.RegistryValue ?? "";
                            if (details.Contains('\n'))
                            {
                                var lines = details.Split('\n');
                                if (lines.Length > 1)
                                {
                                    details = string.Join(" ", lines.Where(l =>
                                        !l.Trim().StartsWith("ServicingInProgress =") &&
                                        !l.Trim().StartsWith("RebootPending =") &&
                                        !l.Trim().StartsWith("RebootInProgress =") &&
                                        !l.Trim().StartsWith("SessionsPendingExclusive =") &&
                                        !l.Trim().StartsWith("LastTrustTime =") &&
                                        !string.IsNullOrWhiteSpace(l)
                                    ).Take(2)).Trim();
                                }
                            }
                            if (details.Length > 150)
                                details = details.Substring(0, 147) + "...";

                            DefaultGridRows.Add(new AnalyzeGridRow
                            {
                                Column1 = item.Name,
                                Column2 = item.Value,
                                Column3 = details,
                                SourceItem = item
                            });
                        }
                    }

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex) { Debug.WriteLine($"CBS RebootStatus error: {ex.Message}"); }
            }

            // Show All Packages as default view
            const string allPkgTitle = "All Packages";
            const string pendSessTitle = "Pending Sessions";
            const string pendPkgTitle = "Pending Packages";
            const string rebootTitle = "Reboot Status";

            ShowAllPackages(allPkgTitle);

            // Build sub-tabs
            SubTabs.Add(new SubcategoryItem
            {
                Title = allPkgTitle,
                IsEnabled = true,
                IsActive = true,
                Tooltip = "View all CBS packages",
                SelectCommand = new RelayCommand(() => ShowAllPackages(allPkgTitle))
            });
            SubTabs.Add(new SubcategoryItem
            {
                Title = pendSessTitle,
                IsEnabled = true,
                Tooltip = "View pending CBS sessions",
                SelectCommand = new RelayCommand(() => ShowPendingSessions(pendSessTitle))
            });
            SubTabs.Add(new SubcategoryItem
            {
                Title = pendPkgTitle,
                IsEnabled = true,
                Tooltip = "View pending CBS packages",
                SelectCommand = new RelayCommand(() => ShowPendingPackages(pendPkgTitle))
            });
            SubTabs.Add(new SubcategoryItem
            {
                Title = rebootTitle,
                IsEnabled = true,
                Tooltip = "View CBS reboot status",
                SelectCommand = new RelayCommand(() => ShowRebootStatus(rebootTitle))
            });
            ShowSubTabs = true;
        }

        private void FilterCbsPackages()
        {
            DefaultGridRows.Clear();

            var search = _cbsSearchText.Trim().ToLowerInvariant();
            var dismFilter = _cbsDismFilter;

            foreach (var (group, package, state, installed, user, visibility, item) in _allCbsPackagesData)
            {
                // DISM /Get-Package filter: only show packages with Visibility == 1
                if (dismFilter && visibility != 1)
                    continue;

                // Search filter: match against Package Group and Package Version
                if (!string.IsNullOrEmpty(search) &&
                    !group.ToLowerInvariant().Contains(search) &&
                    !package.ToLowerInvariant().Contains(search))
                    continue;

                DefaultGridRows.Add(new AnalyzeGridRow
                {
                    Column1 = group,
                    Column2 = package,
                    Column3 = state,
                    Column4 = installed,
                    Column5 = user,
                    SourceItem = item
                });
            }

            if (DefaultGridRows.Count == 0)
                DefaultGridRows.Add(new AnalyzeGridRow { Column1 = "No packages found" });
        }

        private void DisplayUserProfiles(AnalysisSection section)
        {
            // Cache the profiles section for filtering
            var cachedSection = section;

            // Helper to display profiles with a filter
            void ShowProfilesFiltered(string filter, string tabTitle)
            {
                try
                {
                    CurrentMode = ContentMode.DefaultGrid;
                    DefaultGridRows.Clear();

                    GridColumn1Header = "Name";
                    GridColumn2Header = "SID";
                    GridColumn3Header = "Path";
                    GridColumn4Header = "Last Logon";
                    GridColumnCount = 4;
                    ContentHeader = tabTitle;

                    foreach (var item in cachedSection.Items)
                    {
                        if (!item.IsSubSection || item.SubItems == null || item.SubItems.Count == 0)
                            continue;

                        var sid = item.SubItems[0].Value ?? "";

                        // Filter: "Temp" shows only SIDs ending with .bak
                        if (filter == "Temp" && !sid.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
                            continue;

                        var path = item.SubItems.Count > 1 ? item.SubItems[1].Value ?? "" : "";
                        var lastLogon = item.SubItems.Count > 2 ? item.SubItems[2].Value ?? "" : "";

                        DefaultGridRows.Add(new AnalyzeGridRow
                        {
                            Column1 = item.Name,
                            Column2 = sid,
                            Column3 = path,
                            Column4 = lastLogon,
                            IsSubSection = false,
                            SourceItem = item
                        });
                    }

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex) { Debug.WriteLine($"UserProfiles filter error: {ex.Message}"); }
            }

            // Show All Profiles as default
            const string allTitle = "All Profiles";
            const string tempTitle = "Temp Profiles";

            ShowProfilesFiltered("All", allTitle);

            // Build sub-tabs
            SubTabs.Add(new SubcategoryItem
            {
                Title = allTitle,
                IsEnabled = true,
                IsActive = true,
                Tooltip = "Show all user profiles",
                SelectCommand = new RelayCommand(() => ShowProfilesFiltered("All", allTitle))
            });
            SubTabs.Add(new SubcategoryItem
            {
                Title = tempTitle,
                IsEnabled = true,
                Tooltip = "Show temporary (.bak) profiles only",
                SelectCommand = new RelayCommand(() => ShowProfilesFiltered("Temp", tempTitle))
            });
            ShowSubTabs = true;
        }

        private void DisplayGuestAgent(AnalysisSection section)
        {
            // Show the main Guest Agent overview first
            DisplayDefaultGrid(section);

            // Build sub-tabs
            var isSoftware = _loadedHives.Any(h => h.HiveType == HiveType.SOFTWARE);

            var extensionsTab = new SubcategoryItem
            {
                Title = "🔌 Extensions",
                IsEnabled = isSoftware,
                Tooltip = isSoftware
                    ? "View Azure VM Extensions installed on this system"
                    : "Requires SOFTWARE hive to be loaded",
                SelectCommand = new RelayCommand(() =>
                {
                    if (!isSoftware) return;

                    try
                    {
                        var softwareHive = _loadedHives.First(h => h.HiveType == HiveType.SOFTWARE);
                        var extensionsSection = softwareHive.InfoExtractor.GetAzureExtensionsAnalysis();
                        ContentHeader = extensionsSection.Title;

                        // Custom grid layout with wide Extension column
                        CurrentMode = ContentMode.DefaultGrid;
                        DefaultGridRows.Clear();
                        GridColumn1Header = "Property";
                        GridColumn2Header = "Value";
                        GridColumnCount = 2;
                        GridColumn1Star = 4;
                        GridColumn2Star = 1;

                        foreach (var item in extensionsSection.Items)
                        {
                            DefaultGridRows.Add(new AnalyzeGridRow
                            {
                                Column1 = item.Name,
                                Column2 = item.Value,
                                SourceItem = item
                            });
                        }

                        UpdateSubTabActiveState("🔌 Extensions");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Extensions error: {ex.Message}");
                    }
                })
            };

            SubTabs.Add(extensionsTab);
            ShowSubTabs = true;
        }

        private void DisplayStorageFilters(AnalysisSection section)
        {
            // Determine which hive owns this section (SYSTEM)
            var hive = _sectionHiveMap.GetValueOrDefault(section)
                       ?? _loadedHives.FirstOrDefault(h => h.HiveType == HiveType.SYSTEM);
            if (hive == null)
            {
                DisplayDefaultGrid(section);
                return;
            }

            var ext = hive.InfoExtractor;

            // Helper to display filter items in a 2-column grid
            void ShowStorageFilter(string filterKey, string tabTitle)
            {
                try
                {
                    var items = filterKey == "Disk"
                        ? ext.GetDiskFilters()
                        : ext.GetVolumeFilters();

                    CurrentMode = ContentMode.DefaultGrid;
                    DefaultGridRows.Clear();

                    GridColumn1Header = "Property";
                    GridColumn2Header = "Value";
                    GridColumnCount = 2;
                    ContentHeader = tabTitle;

                    foreach (var item in items)
                    {
                        DefaultGridRows.Add(new AnalyzeGridRow
                        {
                            Column1 = item.Name,
                            Column2 = item.Value,
                            SourceItem = item
                        });
                    }

                    UpdateSubTabActiveState(tabTitle);
                }
                catch (Exception ex) { Debug.WriteLine($"StorageFilters error: {ex.Message}"); }
            }

            // Show Disk Filters as default
            const string diskTitle = "\U0001f4be Disk Filters";
            const string volumeTitle = "\U0001f4c0 Volume Filters";

            ShowStorageFilter("Disk", diskTitle);

            // Build sub-tabs
            SubTabs.Add(new SubcategoryItem
            {
                Title = diskTitle,
                IsEnabled = true,
                IsActive = true,
                Tooltip = "Show disk filter drivers",
                SelectCommand = new RelayCommand(() => ShowStorageFilter("Disk", diskTitle))
            });
            SubTabs.Add(new SubcategoryItem
            {
                Title = volumeTitle,
                IsEnabled = true,
                Tooltip = "Show volume filter drivers",
                SelectCommand = new RelayCommand(() => ShowStorageFilter("Volume", volumeTitle))
            });
            ShowSubTabs = true;
        }
    }
}
