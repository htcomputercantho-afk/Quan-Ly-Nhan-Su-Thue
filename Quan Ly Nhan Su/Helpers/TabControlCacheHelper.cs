using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;

namespace TaxPersonnelManagement.Helpers
{
    /// <summary>
    /// Helper lưu cache visual tree cho TabControl trong WPF, giúp chuyển tab tức thì (0ms)
    /// bằng cách giữ lại ContentPresenter của các tab trong bộ nhớ thay vì hủy và dựng lại mỗi lần chuyển tab.
    /// </summary>
    public static class TabControlCacheHelper
    {
        public static readonly DependencyProperty IsCachedProperty =
            DependencyProperty.RegisterAttached(
                "IsCached",
                typeof(bool),
                typeof(TabControlCacheHelper),
                new PropertyMetadata(false, OnIsCachedChanged));

        public static bool GetIsCached(DependencyObject obj) => (bool)obj.GetValue(IsCachedProperty);
        public static void SetIsCached(DependencyObject obj, bool value) => obj.SetValue(IsCachedProperty, value);

        private static readonly DependencyProperty ItemsHolderProperty =
            DependencyProperty.RegisterAttached(
                "ItemsHolder",
                typeof(Grid),
                typeof(TabControlCacheHelper));

        private static void OnIsCachedChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TabControl tc)
            {
                if ((bool)e.NewValue)
                {
                    tc.Loaded += Tc_Loaded;
                    tc.SelectionChanged += Tc_SelectionChanged;
                }
                else
                {
                    tc.Loaded -= Tc_Loaded;
                    tc.SelectionChanged -= Tc_SelectionChanged;
                }
            }
        }

        private static void Tc_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is TabControl tc)
            {
                EnsureGridHolder(tc);
                UpdateSelectedTab(tc);

                // Tải trước (pre-warm) các tab còn lại ở luồng nền DispatcherPriority.Background
                tc.Dispatcher.BeginInvoke(new Action(() =>
                {
                    PrewarmAllTabs(tc);
                }), DispatcherPriority.Background);
            }
        }

        private static void PrewarmAllTabs(TabControl tc)
        {
            var holder = EnsureGridHolder(tc);
            if (holder == null) return;

            var selectedTab = tc.SelectedItem as TabItem;

            foreach (var item in tc.Items)
            {
                if (item is TabItem tab && tab != selectedTab)
                {
                    bool exists = false;
                    foreach (UIElement child in holder.Children)
                    {
                        if (child is ContentPresenter cp && cp.Tag == tab)
                        {
                            exists = true;
                            break;
                        }
                    }

                    if (!exists)
                    {
                        var cp = new ContentPresenter
                        {
                            Content = tab.Content,
                            ContentTemplate = tab.ContentTemplate,
                            ContentTemplateSelector = tab.ContentTemplateSelector,
                            ContentStringFormat = tab.ContentStringFormat,
                            Tag = tab,
                            Visibility = Visibility.Collapsed
                        };
                        holder.Children.Add(cp);
                    }
                }
            }
        }

        private static void Tc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.OriginalSource == sender && sender is TabControl tc)
            {
                UpdateSelectedTab(tc);
            }
        }

        private static Grid? EnsureGridHolder(TabControl tc)
        {
            var holder = (Grid?)tc.GetValue(ItemsHolderProperty);
            if (holder != null) return holder;

            // Tìm PART_SelectedContentHost trong template
            var cp = tc.Template?.FindName("PART_SelectedContentHost", tc) as ContentPresenter;
            if (cp != null && cp.Parent is Border border)
            {
                var fg = (System.Windows.Media.Brush)cp.GetValue(TextElement.ForegroundProperty);
                border.Child = null; // Gỡ ContentPresenter đơn mặc định của WPF
                holder = new Grid();
                if (fg != null)
                {
                    holder.SetValue(TextElement.ForegroundProperty, fg);
                }
                border.Child = holder;
                tc.SetValue(ItemsHolderProperty, holder);
                return holder;
            }
            else if (cp != null && cp.Parent is Panel panel)
            {
                var fg = (System.Windows.Media.Brush)cp.GetValue(TextElement.ForegroundProperty);
                int index = panel.Children.IndexOf(cp);
                panel.Children.RemoveAt(index);
                holder = new Grid();
                if (fg != null)
                {
                    holder.SetValue(TextElement.ForegroundProperty, fg);
                }
                panel.Children.Insert(index, holder);
                tc.SetValue(ItemsHolderProperty, holder);
                return holder;
            }
            return null;
        }

        private static void UpdateSelectedTab(TabControl tc)
        {
            var holder = EnsureGridHolder(tc);
            if (holder == null) return;

            var selectedTab = tc.SelectedItem as TabItem;
            if (selectedTab == null) return;

            ContentPresenter? selectedCp = null;
            foreach (UIElement child in holder.Children)
            {
                if (child is ContentPresenter cp && cp.Tag == selectedTab)
                {
                    selectedCp = cp;
                    break;
                }
            }

            if (selectedCp == null)
            {
                selectedCp = new ContentPresenter
                {
                    Content = selectedTab.Content,
                    ContentTemplate = selectedTab.ContentTemplate,
                    ContentTemplateSelector = selectedTab.ContentTemplateSelector,
                    ContentStringFormat = selectedTab.ContentStringFormat,
                    Tag = selectedTab
                };
                holder.Children.Add(selectedCp);
            }

            foreach (UIElement child in holder.Children)
            {
                if (child is ContentPresenter cp)
                {
                    cp.Visibility = (cp == selectedCp) ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }
    }
}
