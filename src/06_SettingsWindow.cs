using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PsToolbox
{
    public class SettingsWindow : Window
    {
        TextBox _txtMenuRoot;
        ListBox _lstTools;
        CheckBox _chkEnabled;
        TextBlock _txtDescription;
        StackPanel _settingsPanel;
        readonly Dictionary<string, FrameworkElement> _editors = new Dictionary<string, FrameworkElement>(StringComparer.OrdinalIgnoreCase);
        ToolManifest _currentTool;
        bool _loading;

        static readonly Brush AccentBrush = new SolidColorBrush(Color.FromRgb(55, 120, 200));
        static readonly Brush CardBorderBrush = new SolidColorBrush(Color.FromRgb(225, 225, 225));
        static readonly Brush MutedBrush = new SolidColorBrush(Color.FromRgb(110, 110, 110));

        public SettingsWindow()
        {
            Title = "Settings";
            Width = 820;
            Height = 580;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = Brushes.White;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 13;

            var dock = new DockPanel();

            var footer = new Border
            {
                BorderThickness = new Thickness(0, 1, 0, 0),
                BorderBrush = CardBorderBrush,
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                Padding = new Thickness(16, 12, 16, 12)
            };
            DockPanel.SetDock(footer, Dock.Bottom);

            var footerPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var btnSave = new Button
            {
                Content = "Save",
                Padding = new Thickness(14, 7, 14, 7),
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true,
                Background = AccentBrush,
                Foreground = Brushes.White,
                BorderBrush = AccentBrush,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            btnSave.Click += OnSave;
            footerPanel.Children.Add(btnSave);

            var btnCancel = new Button
            {
                Content = "Cancel",
                Padding = new Thickness(14, 7, 14, 7),
                IsCancel = true,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            btnCancel.Click += (s, e) =>
            {
                App.ResetConfigFromDisk();
                DialogResult = false;
                Close();
            };
            footerPanel.Children.Add(btnCancel);
            footer.Child = footerPanel;
            dock.Children.Add(footer);

            var root = new Grid { Margin = new Thickness(18, 16, 18, 16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(230) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            dock.Children.Add(root);

            var globalCard = new Border
            {
                BorderThickness = new Thickness(1),
                BorderBrush = CardBorderBrush,
                Background = new SolidColorBrush(Color.FromRgb(251, 251, 251)),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(14),
                Margin = new Thickness(0, 0, 0, 14)
            };
            Grid.SetRow(globalCard, 0);
            Grid.SetColumnSpan(globalCard, 2);
            root.Children.Add(globalCard);

            var globalStack = new StackPanel();
            globalCard.Child = globalStack;
            globalStack.Children.Add(new TextBlock
            {
                Text = "Global",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = AccentBrush,
                Margin = new Thickness(0, 0, 0, 8)
            });
            globalStack.Children.Add(FieldRow("Parent Menu", _txtMenuRoot = new TextBox { Text = Config.MenuRootText }));
            globalStack.Children.Add(new TextBlock
            {
                Text = "Changes apply when you click Save. Closing the host window removes all registered menu entries.",
                Foreground = MutedBrush,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap
            });

            var leftCard = new Border
            {
                BorderThickness = new Thickness(1),
                BorderBrush = CardBorderBrush,
                Background = Brushes.White,
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 14, 0)
            };
            Grid.SetRow(leftCard, 1);
            Grid.SetColumn(leftCard, 0);
            root.Children.Add(leftCard);

            var leftStack = new StackPanel();
            leftCard.Child = leftStack;
            leftStack.Children.Add(new TextBlock
            {
                Text = "Tools",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = AccentBrush,
                Margin = new Thickness(0, 0, 0, 8)
            });
            _lstTools = new ListBox();
            _lstTools.SelectionChanged += OnToolSelectionChanged;
            leftStack.Children.Add(_lstTools);

            var rightCard = new Border
            {
                BorderThickness = new Thickness(1),
                BorderBrush = CardBorderBrush,
                Background = Brushes.White,
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(16)
            };
            Grid.SetRow(rightCard, 1);
            Grid.SetColumn(rightCard, 1);
            root.Children.Add(rightCard);

            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            rightCard.Child = scroll;

            var rightStack = new StackPanel();
            scroll.Content = rightStack;

            _chkEnabled = new CheckBox
            {
                Content = "Enable this tool",
                Margin = new Thickness(0, 0, 0, 10)
            };
            rightStack.Children.Add(_chkEnabled);

            _txtDescription = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 14),
                Foreground = MutedBrush,
                TextWrapping = TextWrapping.Wrap
            };
            rightStack.Children.Add(_txtDescription);

            rightStack.Children.Add(new TextBlock
            {
                Text = "Defaults",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = AccentBrush,
                Margin = new Thickness(0, 0, 0, 10)
            });

            _settingsPanel = new StackPanel();
            rightStack.Children.Add(_settingsPanel);

            Content = dock;

            LoadToolList();
        }

        void LoadToolList()
        {
            _lstTools.Items.Clear();
            foreach (var tool in App.Tools)
                _lstTools.Items.Add(tool);

            if (_lstTools.Items.Count > 0)
                _lstTools.SelectedIndex = 0;
        }

        void OnToolSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading) return;
            CommitCurrentTool();
            LoadTool((ToolManifest)_lstTools.SelectedItem);
        }

        void LoadTool(ToolManifest tool)
        {
            _loading = true;
            _currentTool = tool;
            _editors.Clear();
            _settingsPanel.Children.Clear();

            if (tool == null)
            {
                _chkEnabled.IsChecked = false;
                _txtDescription.Text = string.Empty;
                _loading = false;
                return;
            }

            _chkEnabled.IsChecked = Config.ToolEnabled(tool);
            _txtDescription.Text = tool.Description ?? string.Empty;

            foreach (var setting in tool.Settings ?? new List<ToolSetting>())
            {
                var editor = CreateEditor(tool, setting);
                _editors[setting.Key] = editor;
                _settingsPanel.Children.Add(FieldRow(setting.Label, editor, setting.Hint));
            }

            if (!_editors.Any())
            {
                _settingsPanel.Children.Add(new TextBlock
                {
                    Text = "No configurable defaults.",
                    Foreground = MutedBrush
                });
            }

            _loading = false;
        }

        FrameworkElement CreateEditor(ToolManifest tool, ToolSetting setting)
        {
            var key = Config.ToolSettingKey(tool.Id, setting.Key);
            var value = Config.Get(key, setting.Default ?? string.Empty);
            var type = (setting.Type ?? "text").ToLowerInvariant();

            if (type == "bool")
            {
                return new CheckBox { IsChecked = value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase) };
            }

            if (type == "choice" || type == "combo")
            {
                var combo = new ComboBox { IsEditable = type == "combo", Width = type == "combo" ? 260 : 220 };
                foreach (var option in setting.Options ?? new string[0])
                    combo.Items.Add(option);

                var matched = combo.Items.Cast<object>().FirstOrDefault(i => string.Equals((string)i, value, StringComparison.OrdinalIgnoreCase));
                if (matched != null)
                {
                    combo.SelectedItem = matched;
                }
                else if (type == "choice")
                {
                    if (combo.Items.Count > 0)
                        combo.SelectedIndex = 0;
                }
                else
                {
                    combo.Text = value ?? string.Empty;
                }

                return combo;
            }

            return new TextBox
            {
                Text = value,
                Width = type == "number" ? 120 : 260
            };
        }

        void CommitCurrentTool()
        {
            if (_currentTool == null) return;

            Config.SetToolEnabled(_currentTool, _chkEnabled.IsChecked == true);
            foreach (var setting in _currentTool.Settings ?? new List<ToolSetting>())
            {
                FrameworkElement editor;
                if (!_editors.TryGetValue(setting.Key, out editor))
                    continue;
                Config.Set(Config.ToolSettingKey(_currentTool.Id, setting.Key), ReadValue(editor, setting));
            }
        }

        string ReadValue(FrameworkElement editor, ToolSetting setting)
        {
            var type = (setting.Type ?? "text").ToLowerInvariant();
            if (type == "bool")
                return (((CheckBox)editor).IsChecked == true) ? "1" : "0";
            if (type == "choice")
                return (string)((ComboBox)editor).SelectedItem ?? (setting.Default ?? string.Empty);
            if (type == "combo")
                return ((ComboBox)editor).Text ?? string.Empty;
            return ((TextBox)editor).Text ?? string.Empty;
        }

        Grid FieldRow(string label, FrameworkElement input, string hint = null)
        {
            var grid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var labelBlock = new TextBlock
            {
                Text = label,
                Margin = new Thickness(0, 4, 10, 0),
                Foreground = new SolidColorBrush(Color.FromRgb(80, 80, 80))
            };
            Grid.SetColumn(labelBlock, 0);
            grid.Children.Add(labelBlock);

            var right = new StackPanel();
            right.Children.Add(input);
            if (!string.IsNullOrWhiteSpace(hint))
            {
                right.Children.Add(new TextBlock
                {
                    Text = hint,
                    Margin = new Thickness(0, 4, 0, 0),
                    Foreground = MutedBrush,
                    FontSize = 11,
                    TextWrapping = TextWrapping.Wrap
                });
            }
            Grid.SetColumn(right, 1);
            grid.Children.Add(right);
            return grid;
        }

        void OnSave(object sender, RoutedEventArgs e)
        {
            CommitCurrentTool();
            Config.MenuRootText = _txtMenuRoot.Text;
            Config.Save();
            App.ApplyContextMenus();
            DialogResult = true;
            Close();
        }
    }
}


