using System.Windows;

namespace RegistryExpert.Wpf.Helpers;

/// <summary>
/// A Freezable-based proxy that allows non-visual-tree elements (e.g. DataGridColumn)
/// to bind to the window's DataContext. Freezable objects inherit the ambient DataContext
/// when declared in a ResourceDictionary, bridging the gap for elements like
/// DataGridTemplateColumn that cannot use RelativeSource AncestorType bindings.
/// </summary>
public class BindingProxy : Freezable
{
    public static readonly DependencyProperty DataProperty =
        DependencyProperty.Register(
            nameof(Data),
            typeof(object),
            typeof(BindingProxy),
            new UIPropertyMetadata(null));

    public object Data
    {
        get => GetValue(DataProperty);
        set => SetValue(DataProperty, value);
    }

    protected override Freezable CreateInstanceCore() => new BindingProxy();
}
