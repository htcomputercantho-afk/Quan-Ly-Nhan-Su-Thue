using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace TaxPersonnelManagement.Helpers
{
    /// <summary>
    /// Attached property cắt (clip) nội dung của một phần tử theo góc bo tròn,
    /// đồng thời tự động xử lý triệt để phần khoảng trắng phía trên thanh cuộn dọc (Vertical ScrollBar)
    /// của DataGrid để thanh tiêu đề màu xanh kéo dài hết chiều ngang bảng và bo tròn góc trên phải hoàn hảo.
    /// </summary>
    public static class ClipHelper
    {
        public static readonly DependencyProperty CornerRadiusProperty =
            DependencyProperty.RegisterAttached(
                "CornerRadius",
                typeof(CornerRadius),
                typeof(ClipHelper),
                new PropertyMetadata(new CornerRadius(0), OnCornerRadiusChanged));

        public static CornerRadius GetCornerRadius(DependencyObject obj) => (CornerRadius)obj.GetValue(CornerRadiusProperty);
        public static void SetCornerRadius(DependencyObject obj, CornerRadius value) => obj.SetValue(CornerRadiusProperty, value);

        private static readonly DependencyProperty InternalRadiusProperty =
            DependencyProperty.RegisterAttached("InternalRadius", typeof(CornerRadius), typeof(ClipHelper), new PropertyMetadata(new CornerRadius(0)));

        private static void SetCornerRadiusInternal(DependencyObject obj, CornerRadius r) => obj.SetValue(InternalRadiusProperty, r);
        private static CornerRadius GetCornerRadiusInternal(DependencyObject obj) => (CornerRadius)obj.GetValue(InternalRadiusProperty);

        private static void OnCornerRadiusChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not FrameworkElement fe) return;

            fe.SizeChanged -= OnSizeChanged;
            fe.Loaded -= OnLoaded;
            fe.SizeChanged += OnSizeChanged;
            fe.Loaded += OnLoaded;
            UpdateClip(fe);
            FixDataGridHeaderCorners(fe);
        }

        private static void OnLoaded(object sender, RoutedEventArgs e)
        {
            var fe = (FrameworkElement)sender;
            UpdateClip(fe);
            FixDataGridHeaderCorners(fe);
        }

        private static void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            var fe = (FrameworkElement)sender;
            UpdateClip(fe);
            FixDataGridHeaderCorners(fe);
        }

        /// <summary>
        /// Xử lý triệt để khoảng trắng phía trên thanh cuộn dọc của DataGrid:
        /// 1. Tìm hoặc tạo ô nền xanh (Border) tại Column 2, Row 0 ngay phía trên thanh cuộn dọc.
        /// 2. Kéo dài DataGridColumnHeadersPresenter phủ sang cột cuộn dọc (Grid.ColumnSpan = 2).
        /// 3. Gán màu nền xanh đồng bộ cho Presenter và ô góc để không bao giờ lộ khoảng trắng.
        /// 4. Bo tròn trực tiếp thanh Header xanh để cả góc trái và góc phải đều bo tròn 12px đều đặn.
        /// </summary>
        public static void FixDataGridHeaderCorners(FrameworkElement root)
        {
            var r = GetCornerRadius(root);
            if (r.TopLeft == 0 && r.TopRight == 0 && r.BottomLeft == 0 && r.BottomRight == 0)
            {
                r = new CornerRadius(12);
            }
            var headerRadius = new CornerRadius(r.TopLeft, r.TopRight, 0, 0);

            var dgs = new List<DataGrid>();
            if (root is DataGrid selfDg)
            {
                dgs.Add(selfDg);
            }
            dgs.AddRange(FindVisualChildren<DataGrid>(root));

            foreach (var dg in dgs)
            {
                ApplyDataGridFix(dg, headerRadius);
            }
        }

        private static void ApplyDataGridFix(DataGrid dg, CornerRadius headerRadius)
        {
            var presenter = FindVisualChild<DataGridColumnHeadersPresenter>(dg);
            if (presenter == null)
            {
                // Nếu DataGrid chưa load xong template thì đăng ký Loaded & LayoutUpdated
                RoutedEventHandler? loadedHandler = null;
                loadedHandler = (s, e) =>
                {
                    dg.Loaded -= loadedHandler;
                    ApplyDataGridFix(dg, headerRadius);
                };
                dg.Loaded += loadedHandler;

                EventHandler? layoutHandler = null;
                layoutHandler = (s, e) =>
                {
                    var p = FindVisualChild<DataGridColumnHeadersPresenter>(dg);
                    if (p != null)
                    {
                        dg.LayoutUpdated -= layoutHandler;
                        ApplyDataGridFix(dg, headerRadius);
                    }
                };
                dg.LayoutUpdated += layoutHandler;
                return;
            }

            // Lấy màu nền chuẩn của Header
            Brush? headerBrush = presenter.Background;
            if (headerBrush == null || headerBrush == Brushes.Transparent)
            {
                headerBrush = Application.Current?.TryFindResource("PrimaryHueMidBrush") as Brush
                              ?? new SolidColorBrush(Color.FromRgb(0, 120, 212));
                presenter.Background = headerBrush;
            }

            if (VisualTreeHelper.GetParent(presenter) is Grid parentGrid && parentGrid.ColumnDefinitions.Count >= 3)
            {
                // Đảm bảo Presenter bắt đầu từ cột 1 và phủ sang cả cột 2 (cột thanh cuộn)
                Grid.SetColumn(presenter, 1);
                Grid.SetColumnSpan(presenter, 2);

                // Thêm ô nền bù màu xanh (HeaderCornerFiller) vào Column 2, Row 0
                // để chắc chắn góc trên phải phía trên ScrollBar luôn luôn có màu xanh hoàn hảo
                Border? cornerFiller = null;
                foreach (UIElement child in parentGrid.Children)
                {
                    if (child is Border b && (string?)b.Tag == "HeaderCornerFiller")
                    {
                        cornerFiller = b;
                        break;
                    }
                }

                if (cornerFiller == null)
                {
                    cornerFiller = new Border
                    {
                        Tag = "HeaderCornerFiller",
                        Background = headerBrush,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Stretch
                    };
                    Grid.SetRow(cornerFiller, 0);
                    Grid.SetColumn(cornerFiller, 2);
                    Panel.SetZIndex(cornerFiller, 0);
                    Panel.SetZIndex(presenter, 1);
                    parentGrid.Children.Add(cornerFiller);
                }
                else
                {
                    cornerFiller.Background = headerBrush;
                }
            }

            // Bo tròn trực tiếp thanh Header xanh
            presenter.SizeChanged -= OnElementSizeChanged;
            presenter.SizeChanged += OnElementSizeChanged;
            SetCornerRadiusInternal(presenter, headerRadius);
            ApplyGeometryClip(presenter, headerRadius);

            // Hook ScrollViewer để cập nhật khi thanh cuộn xuất hiện hoặc ẩn đi
            var scrollViewer = FindVisualChild<ScrollViewer>(dg);
            if (scrollViewer != null)
            {
                scrollViewer.ScrollChanged -= OnScrollViewerScrollChanged;
                scrollViewer.ScrollChanged += OnScrollViewerScrollChanged;
            }
        }

        private static void OnScrollViewerScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer sv && VisualTreeHelper.GetParent(sv) is DependencyObject parent)
            {
                var dg = FindVisualAncestor<DataGrid>(sv);
                if (dg != null)
                {
                    var presenter = FindVisualChild<DataGridColumnHeadersPresenter>(dg);
                    if (presenter != null)
                    {
                        ApplyGeometryClip(presenter, GetCornerRadiusInternal(presenter));
                    }
                }
            }
        }

        private static void OnElementSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (sender is FrameworkElement fe)
            {
                ApplyGeometryClip(fe, GetCornerRadiusInternal(fe));
            }
        }

        private static void UpdateClip(FrameworkElement fe)
        {
            ApplyGeometryClip(fe, GetCornerRadius(fe));
        }

        private static void ApplyGeometryClip(FrameworkElement fe, CornerRadius r)
        {
            double w = fe.ActualWidth, h = fe.ActualHeight;
            if (w <= 0 || h <= 0)
            {
                return;
            }
            if (r.TopLeft == 0 && r.TopRight == 0 && r.BottomRight == 0 && r.BottomLeft == 0)
            {
                fe.Clip = null;
                return;
            }

            double max = Math.Min(w, h) / 2;
            double tl = Math.Min(r.TopLeft, max), tr = Math.Min(r.TopRight, max);
            double br = Math.Min(r.BottomRight, max), bl = Math.Min(r.BottomLeft, max);

            var geo = new StreamGeometry();
            using (var ctx = geo.Open())
            {
                ctx.BeginFigure(new Point(tl, 0), true, true);
                ctx.LineTo(new Point(w - tr, 0), false, false);
                if (tr > 0) ctx.ArcTo(new Point(w, tr), new Size(tr, tr), 0, false, SweepDirection.Clockwise, false, false);
                ctx.LineTo(new Point(w, h - br), false, false);
                if (br > 0) ctx.ArcTo(new Point(w - br, h), new Size(br, br), 0, false, SweepDirection.Clockwise, false, false);
                ctx.LineTo(new Point(bl, h), false, false);
                if (bl > 0) ctx.ArcTo(new Point(0, h - bl), new Size(bl, bl), 0, false, SweepDirection.Clockwise, false, false);
                ctx.LineTo(new Point(0, tl), false, false);
                if (tl > 0) ctx.ArcTo(new Point(tl, 0), new Size(tl, tl), 0, false, SweepDirection.Clockwise, false, false);
            }
            geo.Freeze();
            fe.Clip = geo;
        }

        public static T? FindVisualChild<T>(DependencyObject? parent) where T : DependencyObject
        {
            if (parent == null) return null;
            if (parent is T matched) return matched;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                var result = FindVisualChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject? parent) where T : DependencyObject
        {
            if (parent == null) yield break;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T matched) yield return matched;
                foreach (var item in FindVisualChildren<T>(child))
                    yield return item;
            }
        }

        public static T? FindVisualAncestor<T>(DependencyObject? obj) where T : DependencyObject
        {
            while (obj != null)
            {
                if (obj is T matched) return matched;
                obj = VisualTreeHelper.GetParent(obj);
            }
            return null;
        }
    }
}
