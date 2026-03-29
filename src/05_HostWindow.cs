using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PsToolbox
{
    public class HostWindow : Window
    {
        ItemsControl _toolList;

        static readonly Brush AccentBrush = new SolidColorBrush(Color.FromRgb(55, 120, 200));
        static readonly Brush BorderBrushThin = new SolidColorBrush(Color.FromRgb(225, 225, 225));

        public HostWindow()
        {
            Title = "toolbox - active";
            Width = 300;
            Height = 212;
            MinWidth = 300;
            MinHeight = 212;
            MaxWidth = 300;
            MaxHeight = 212;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = SystemParameters.WorkArea.Right - 320;
            Top = SystemParameters.WorkArea.Top + 40;
            Background = Brushes.White;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 12;

            var dock = new DockPanel();

            var footer = new Border
            {
                BorderThickness = new Thickness(0, 1, 0, 0),
                BorderBrush = BorderBrushThin,
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                Padding = new Thickness(12, 10, 12, 10)
            };
            DockPanel.SetDock(footer, Dock.Bottom);

            var footerGrid = new Grid();
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var btnSettings = new Button
            {
                Content = "Settings",
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(10, 5, 10, 5),
                Cursor = System.Windows.Input.Cursors.Hand,
                MinWidth = 72
            };
            btnSettings.Click += (s, e) =>
            {
                var settings = new SettingsWindow();
                settings.Owner = this;
                settings.ShowDialog();
                RefreshState();
            };
            Grid.SetColumn(btnSettings, 1);
            footerGrid.Children.Add(btnSettings);

            var btnExit = new Button
            {
                Content = "Exit",
                Padding = new Thickness(10, 5, 10, 5),
                Cursor = System.Windows.Input.Cursors.Hand,
                Background = AccentBrush,
                Foreground = Brushes.White,
                BorderBrush = AccentBrush,
                MinWidth = 60
            };
            btnExit.Click += (s, e) => Close();
            Grid.SetColumn(btnExit, 2);
            footerGrid.Children.Add(btnExit);

            footer.Child = footerGrid;
            dock.Children.Add(footer);

            var listBorder = new Border
            {
                BorderThickness = new Thickness(1),
                BorderBrush = BorderBrushThin,
                Background = Brushes.White,
                Margin = new Thickness(12, 12, 12, 12),
                Padding = new Thickness(0)
            };
            dock.Children.Add(listBorder);

            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(0)
            };
            listBorder.Child = scroll;

            _toolList = new ItemsControl();
            _toolList.ItemTemplate = BuildItemTemplate();
            scroll.Content = _toolList;

            Content = dock;

            Loaded += (s, e) => RefreshState();
            Closing += (s, e) =>
            {
                try { ContextMenuManager.Cleanup(App.Tools); }
                catch { }
            };
        }

        public void RefreshState()
        {
            _toolList.ItemsSource = App.EnabledTools().ToList();
        }

        static DataTemplate BuildItemTemplate()
        {
            var template = new DataTemplate(typeof(ToolManifest));

            var rowBorder = new FrameworkElementFactory(typeof(Border));
            rowBorder.SetValue(Border.BorderThicknessProperty, new Thickness(0, 0, 0, 1));
            rowBorder.SetValue(Border.BorderBrushProperty, BorderBrushThin);
            rowBorder.SetValue(Border.PaddingProperty, new Thickness(10, 7, 10, 7));

            var panel = new FrameworkElementFactory(typeof(DockPanel));

            var marker = new FrameworkElementFactory(typeof(Border));
            marker.SetValue(DockPanel.DockProperty, Dock.Left);
            marker.SetValue(FrameworkElement.WidthProperty, 3.0);
            marker.SetValue(FrameworkElement.HeightProperty, 14.0);
            marker.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 1, 8, 0));
            marker.SetValue(Border.BackgroundProperty, AccentBrush);
            panel.AppendChild(marker);

            var text = new FrameworkElementFactory(typeof(TextBlock));
            text.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("DisplayName"));
            text.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            panel.AppendChild(text);

            rowBorder.AppendChild(panel);
            template.VisualTree = rowBorder;
            return template;
        }
    }
}
