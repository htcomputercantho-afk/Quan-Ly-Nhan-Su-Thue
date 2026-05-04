using Microsoft.EntityFrameworkCore;
using System.Linq;
using System.Windows.Controls;
using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using Microsoft.Win32;
using ClosedXML.Excel;

namespace TaxPersonnelManagement.Views
{
    public partial class SalaryListView : UserControl
    {
        // Pagination
        private List<Personnel> _fullSalaryList = new List<Personnel>();
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;

        public SalaryListView()
        {
            InitializeComponent();
            ApplyAuthorization();
            LoadFilterOptions();
            LoadData();
        }

        private void ApplyAuthorization()
        {
        }

        private void AdminOnly_Loaded(object sender, RoutedEventArgs e)
        {
            if (App.CurrentUser?.Role == UserRole.Staff && sender is FrameworkElement element)
            {
                element.Visibility = Visibility.Collapsed;
            }
        }

        public class FilterItem
        {
            public string Label { get; set; } = string.Empty;
            public int Value { get; set; }
            public FilterType Type { get; set; }
            public bool IsHeader { get; set; }

            public override string ToString() => Label;
        }

        public enum FilterType
        {
            None,
            Month,
            Quarter,
            HalfYear,
            Header
        }

        private void LoadFilterOptions()
        {
            var years = new List<FilterItem>();
            years.Add(new FilterItem { Label = "-- Tất cả các năm --", Value = 0 });
            int currentYear = DateTime.Now.Year;
            for (int i = 2025; i <= currentYear + 5; i++)
            {
                years.Add(new FilterItem { Label = $"Năm {i}", Value = i });
            }
            cbYear.ItemsSource = years;
            cbYear.SelectedIndex = 0;

            var periods = new List<FilterItem>();
            periods.Add(new FilterItem { Label = "-- Cả năm --", Value = 0, Type = FilterType.None });

            periods.Add(new FilterItem { Label = "Theo Quý", IsHeader = true, Type = FilterType.Header });
            periods.Add(new FilterItem { Label = "Quý I (Tháng 1 - 3)", Value = 1, Type = FilterType.Quarter });
            periods.Add(new FilterItem { Label = "Quý II (Tháng 4 - 6)", Value = 2, Type = FilterType.Quarter });
            periods.Add(new FilterItem { Label = "Quý III (Tháng 7 - 9)", Value = 3, Type = FilterType.Quarter });
            periods.Add(new FilterItem { Label = "Quý IV (Tháng 10 - 12)", Value = 4, Type = FilterType.Quarter });

            periods.Add(new FilterItem { Label = "Theo Bán niên", IsHeader = true, Type = FilterType.Header });
            periods.Add(new FilterItem { Label = "6 tháng đầu năm", Value = 1, Type = FilterType.HalfYear });
            periods.Add(new FilterItem { Label = "6 tháng cuối năm", Value = 2, Type = FilterType.HalfYear });

            periods.Add(new FilterItem { Label = "Chi tiết theo Tháng", IsHeader = true, Type = FilterType.Header });
            for (int i = 1; i <= 12; i++)
            {
                periods.Add(new FilterItem { Label = $"Tháng {i}", Value = i, Type = FilterType.Month });
            }

            cbPeriod.ItemsSource = periods;
            cbPeriod.SelectedIndex = 0;
        }

        private void cbFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    var query = context.Personnel.AsQueryable();

                    if (cbYear.SelectedItem is FilterItem selectedYear && selectedYear.Value > 0)
                    {
                        query = query.Where(p => p.ExpectedSalaryIncreaseDate.HasValue && p.ExpectedSalaryIncreaseDate.Value.Year == selectedYear.Value);
                    }

                    if (cbPeriod.SelectedItem is FilterItem selectedPeriod && selectedPeriod.Value > 0)
                    {
                        query = query.Where(p => p.ExpectedSalaryIncreaseDate.HasValue);

                        switch (selectedPeriod.Type)
                        {
                            case FilterType.Month:
                                query = query.Where(p => p.ExpectedSalaryIncreaseDate!.Value.Month == selectedPeriod.Value);
                                break;
                            case FilterType.Quarter:
                                int startMonth = (selectedPeriod.Value - 1) * 3 + 1;
                                int endMonth = startMonth + 2;
                                query = query.Where(p => p.ExpectedSalaryIncreaseDate!.Value.Month >= startMonth && p.ExpectedSalaryIncreaseDate!.Value.Month <= endMonth);
                                break;
                            case FilterType.HalfYear:
                                if (selectedPeriod.Value == 1)
                                    query = query.Where(p => p.ExpectedSalaryIncreaseDate!.Value.Month >= 1 && p.ExpectedSalaryIncreaseDate!.Value.Month <= 6);
                                else
                                    query = query.Where(p => p.ExpectedSalaryIncreaseDate!.Value.Month >= 7 && p.ExpectedSalaryIncreaseDate!.Value.Month <= 12);
                                break;
                        }
                    }

                    _fullSalaryList = query.OrderBy(p => p.FullName).ToList();
                    _currentPage = 1;
                    ApplyPagination();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải dữ liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyPagination()
        {
            int totalItems = _fullSalaryList.Count;
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));

            if (_currentPage > _totalPages) _currentPage = _totalPages;
            if (_currentPage < 1) _currentPage = 1;

            var pageItems = _fullSalaryList
                .Skip((_currentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            dgSalary.ItemsSource = pageItems;

            // Update footer
            if (totalItems == 0)
            {
                txtPagingInfo.Text = "Không có dữ liệu";
                txtPageInfo.Text = "0 / 0";
            }
            else
            {
                int startIndex = (_currentPage - 1) * PageSize;
                int from = startIndex + 1;
                int to = startIndex + pageItems.Count;
                txtPagingInfo.Text = $"Hiển thị {from} - {to} trên {totalItems} nhân sự";
                txtPageInfo.Text = $"{_currentPage} / {_totalPages}";
            }

            btnPrevPage.IsEnabled = _currentPage > 1;
            btnNextPage.IsEnabled = _currentPage < _totalPages;
        }

        private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage > 1)
            {
                _currentPage--;
                ApplyPagination();
            }
        }

        private void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage < _totalPages)
            {
                _currentPage++;
                ApplyPagination();
            }
        }

        private void dgSalary_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void btnEditSalary_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is Personnel p)
            {
                if (Application.Current.MainWindow is MainWindow mw)
                {
                    // Tab 6 (Salary Info) is index 5
                    mw.NavigateToPersonnelDetail(p, 5);
                }
            }
        }

        /// <summary>
        /// Tạo tên file Excel động dựa trên bộ lọc Năm và Kỳ xem đang chọn.
        /// VD: DanhSachNangLuong_Nam2026_QuýI_20260314.xlsx
        /// Nếu không chọn bộ lọc: DanhSachNangLuong_20260314.xlsx
        /// </summary>
        private string GenerateExportFileName()
        {
            string yearPart = "";
            if (cbYear.SelectedItem is FilterItem yi && yi.Value > 0)
                yearPart = $"_Nam{yi.Value}";

            string periodPart = "";
            if (cbPeriod.SelectedItem is FilterItem pi && pi.Value > 0)
            {
                string label = pi.Label.Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "");
                periodPart = $"_{label}";
            }

            return $"DanhSachNangLuong{yearPart}{periodPart}_{DateTime.Now:yyyyMMdd}.xlsx";
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _fullSalaryList;
                if (data == null || !data.Any())
                {
                    MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = GenerateExportFileName()
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Danh sách lương");

                        // 1. Tiêu đề file Excel - thay đổi theo bộ lọc năm/kỳ
                        string yearText = "Năm " + (cbYear.SelectedItem is FilterItem y ? y.Value.ToString() : DateTime.Now.Year.ToString());
                        if (cbYear.SelectedItem is FilterItem yi && yi.Value == 0) yearText = ""; 
                        
                        string periodText = cbPeriod.SelectedItem?.ToString() ?? "";
                        bool isAllPeriod = (cbPeriod.SelectedItem is FilterItem pi && pi.Value == 0);
                        
                        string title;
                        if (string.IsNullOrEmpty(yearText) && isAllPeriod)
                        {
                            title = "DANH SÁCH NÂNG LƯƠNG";
                        }
                        else if (isAllPeriod)
                        {
                            title = $"DANH SÁCH NHÂN SỰ ĐẾN HẠN NÂNG LƯƠNG - {yearText}".ToUpper();
                        }
                        else
                        {
                            title = $"DANH SÁCH NHÂN SỰ ĐẾN HẠN NÂNG LƯƠNG - {yearText} ({periodText})".ToUpper();
                        }
                        
                        var titleRange = worksheet.Range("A1:K1");
                        titleRange.Merge();
                        titleRange.Value = title;
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 14;
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                        titleRange.Style.Border.OutsideBorderColor = XLColor.DarkGreen;
                        titleRange.Style.Font.FontColor = XLColor.DarkGreen;

                        // 2. Tiêu đề các cột
                        string[] headers = { "STT", "Họ và tên", "Ngày sinh", "Mã ngạch", "Bậc lương", "Hệ số lương", "% Vượt khung", "Thời điểm tính bậc lương lần sau", "Tháng", "Dự kiến lên lương", "Ghi chú" };
                        
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(3, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Alignment.WrapText = true;
                        }

                        // 3. Ghi dữ liệu từng dòng
                        int row = 4;
                        int stt = 1;
                        foreach (var item in data)
                        {
                            worksheet.Cell(row, 1).Value = stt++;
                            worksheet.Cell(row, 2).Value = item.FullName;
                            worksheet.Cell(row, 3).Value = item.DateOfBirth; 
                            worksheet.Cell(row, 4).Value = item.RankCode;
                            worksheet.Cell(row, 5).Value = item.CurrentSalaryStep;
                            worksheet.Cell(row, 6).Value = item.CurrentSalaryCoefficient;
                            worksheet.Cell(row, 7).Value = item.ExceedFramePercent > 0 ? $"{item.ExceedFramePercent}%" : "";
                            
                             if (item.NextSalaryStepDate.HasValue)
                                worksheet.Cell(row, 8).Value = item.NextSalaryStepDate.Value;
                             
                             if (item.ExpectedSalaryIncreaseDate.HasValue)
                             {
                                worksheet.Cell(row, 9).Value = item.ExpectedSalaryIncreaseDate.Value.Month;
                                var nextDateCell = worksheet.Cell(row, 10);
                                nextDateCell.Value = item.ExpectedSalaryIncreaseDate.Value;
                                nextDateCell.Style.Font.FontColor = XLColor.Red;
                             }

                            worksheet.Cell(row, 11).Value = ""; // No Note property in Entity

                            // Borders
                            for (int col = 1; col <= 11; col++)
                            {
                                worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            }

                            row++;
                        }

                        // 4. Định dạng: Tự động điều chỉnh độ rộng cột, căn giữa dữ liệu
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(8).Width = 20; 
                        worksheet.Column(10).Width = 20;

                        // Alignment: Center all data cells by default, except Name (Col 2)
                        var dataRange = worksheet.Range(4, 1, row - 1, 11);
                        dataRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        
                        // Left align Names
                        var nameRange = worksheet.Range(4, 2, row - 1, 2);
                        nameRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất danh sách lương thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi khi xuất Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }


}
