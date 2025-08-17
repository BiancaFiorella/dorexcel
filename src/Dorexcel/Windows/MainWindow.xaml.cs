using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;

namespace Dorexcel.Windows;

public sealed partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        ExtendsContentIntoTitleBar = true;
    }

    private void OnWindowActivated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState == WindowActivationState.Deactivated) return;

        OverlappedPresenter presenter = (OverlappedPresenter)AppWindow.Presenter;
        presenter.Maximize();
    }
}
