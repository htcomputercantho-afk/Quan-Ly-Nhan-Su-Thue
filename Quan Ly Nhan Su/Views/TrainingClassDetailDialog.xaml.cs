using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using ClosedXML.Excel;
using Microsoft.Win32;


namespace TaxPersonnelManagement.Views
{
    public partial class TrainingClassDetailDialog : Window
    {
        private int _classId;
        private TrainingClass? _trainingClass;
        private List<Personnel> _participants = new();
        private bool _isDataChanged = false;

        public bool IsDataChanged => _isDataChanged;

        public TrainingClassDetailDialog(int classId)
        {
            InitializeComponent();
            _classId = classId;
            LoadClassInfo();
            LoadParticipants();
            LoadPersonnelComboBox();
        }

        private void LoadClassInfo()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    _trainingClass = db.TrainingClasses.Find(_classId);
                    if (_trainingClass != null)
                    {
                        lblClassName.Text = _trainingClass.ClassName;
                        lblParticipationDate.Text = _trainingClass.ParticipationDate?.ToString("dd/MM/yyyy") ?? "--/--/----";
                        lblDecisionNumber.Text = string.IsNullOrEmpty(_trainingClass.DecisionNumber) ? "Không có" : _trainingClass.DecisionNumber;
                        lblDecisionDate.Text = _trainingClass.DecisionDate?.ToString("dd/MM/yyyy") ?? "--/--/----";
                        lblDecisionUnit.Text = string.IsNullOrEmpty(_trainingClass.DecisionUnit) ? "Không có" : _trainingClass.DecisionUnit;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải thông tin lớp học: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadParticipants()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var query = db.PersonnelTrainings
                                  .Where(pt => pt.TrainingClassId == _classId)
                                  .Include(pt => pt.Personnel)
                                  .Select(pt => pt.Personnel)
                                  .ToList();

                    _participants = query;

                    // Set STT
                    int stt = 1;
                    foreach (var p in _participants)
                    {
                        p.STT = stt++;
                    }

                    dgParticipants.ItemsSource = null;
                    dgParticipants.ItemsSource = _participants;
                    lblCount.Text = $"Số lượng: {_participants.Count} học viên";
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh sách học viên: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private bool _isFiltering = false;

        private void LoadPersonnelComboBox()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    // Load all personnel and order by FullName
                    var allPersonnel = db.Personnel
                                         .OrderBy(p => p.FullName)
                                         .ToList();

                    var comboSource = allPersonnel.Select(p => new PersonnelComboItem
                    {
                        Id = p.Id,
                        FullName = p.FullName,
                        Department = p.Department ?? "Không bộ phận",
                        IdentityCardNumber = p.IdentityCardNumber ?? "Không có số CCCD"
                    }).ToList();

                    cboAddPersonnel.ItemsSource = comboSource;
                    cboAddPersonnel.SelectedValuePath = "Id";
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh mục cán bộ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void cboAddPersonnel_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_isFiltering) return;

            var textBox = (TextBox)e.OriginalSource;
            if (textBox == null || !textBox.IsFocused) return;

            string filterText = textBox.Text;

            _isFiltering = true;
            try
            {
                var view = System.Windows.Data.CollectionViewSource.GetDefaultView(cboAddPersonnel.ItemsSource);
                if (view != null)
                {
                    if (string.IsNullOrWhiteSpace(filterText))
                    {
                        view.Filter = null;
                    }
                    else
                    {
                        view.Filter = item =>
                        {
                            if (item is PersonnelComboItem comboItem)
                            {
                                return comboItem.FullName.Contains(filterText, StringComparison.OrdinalIgnoreCase) ||
                                       comboItem.Department.Contains(filterText, StringComparison.OrdinalIgnoreCase) ||
                                       comboItem.IdentityCardNumber.Contains(filterText, StringComparison.OrdinalIgnoreCase);
                            }
                            return false;
                        };
                    }

                    cboAddPersonnel.IsDropDownOpen = true;

                    // Restore text and caret position if WPF cleared it
                    if (textBox.Text != filterText)
                    {
                        textBox.Text = filterText;
                        textBox.SelectionStart = filterText.Length;
                    }
                }
            }
            finally
            {
                _isFiltering = false;
            }
        }

        private void cboAddPersonnel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Clear filter on selection so all items are available when reopening
            var view = System.Windows.Data.CollectionViewSource.GetDefaultView(cboAddPersonnel.ItemsSource);
            if (view != null)
            {
                view.Filter = null;
            }
        }

        private void btnAddParticipant_Click(object sender, RoutedEventArgs e)
        {
            if (cboAddPersonnel.SelectedValue == null)
            {
                var warning = new WarningWindow("Vui lòng chọn một cán bộ từ danh sách!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            int personnelId = (int)cboAddPersonnel.SelectedValue;

            try
            {
                using (var db = new AppDbContext())
                {
                    // Check if already in this class
                    bool exists = db.PersonnelTrainings.Any(pt => pt.TrainingClassId == _classId && pt.PersonnelId == personnelId);
                    if (exists)
                    {
                        var warning = new WarningWindow("Cán bộ này đã có tên trong danh sách học viên!", "Thông báo");
                        warning.Owner = this;
                        warning.ShowDialog();
                        return;
                    }

                    var pt = new PersonnelTraining
                    {
                        TrainingClassId = _classId,
                        PersonnelId = personnelId
                    };
                    db.PersonnelTrainings.Add(pt);
                    db.SaveChanges();

                    _isDataChanged = true;
                    LoadParticipants();
                    cboAddPersonnel.SelectedIndex = -1;
                    cboAddPersonnel.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi thêm học viên: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnRemoveParticipant_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is int personnelId)
            {
                var confirm = new ConfirmWindow("Bạn có chắc chắn muốn xóa học viên này khỏi lớp học?", "Xác nhận");
                confirm.Owner = this;
                if (confirm.ShowDialog() == true)
                {
                    try
                    {
                        using (var db = new AppDbContext())
                        {
                            var pt = db.PersonnelTrainings.FirstOrDefault(pt => pt.TrainingClassId == _classId && pt.PersonnelId == personnelId);
                            if (pt != null)
                            {
                                db.PersonnelTrainings.Remove(pt);
                                db.SaveChanges();

                                _isDataChanged = true;
                                LoadParticipants();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        var warning = new WarningWindow($"Lỗi xóa học viên: {ex.Message}", "Lỗi");
                        warning.Owner = this;
                        warning.ShowDialog();
                    }
                }
            }
        }

        private void btnEditClass_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddTrainingClassDialog(_classId);
            dialog.Owner = this;
            if (dialog.ShowDialog() == true)
            {
                _isDataChanged = true;
                LoadClassInfo();
            }
        }

        private void btnDeleteClass_Click(object sender, RoutedEventArgs e)
        {
            var confirm = new ConfirmWindow("Bạn có chắc chắn muốn xóa lớp học/hội nghị này? Toàn bộ lịch sử tham gia của học viên sẽ bị xóa và không thể khôi phục!", "Cảnh báo");
            confirm.Owner = this;
            if (confirm.ShowDialog() == true)
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var tc = db.TrainingClasses.Find(_classId);
                        if (tc != null)
                        {
                            db.TrainingClasses.Remove(tc);
                            db.SaveChanges();

                            _isDataChanged = true;
                            DialogResult = true; // Tell list view to reload
                            Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Lỗi xóa lớp học: {ex.Message}", "Lỗi");
                    warning.Owner = this;
                    warning.ShowDialog();
                }
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (!_participants.Any())
            {
                var warning = new WarningWindow("Không có học viên nào trong lớp để xuất!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"DanhSachHocVien_{lblClassName.Text.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Học viên");

                        // 1. Title & Info
                        worksheet.Cell(1, 1).Value = "DANH SÁCH HỌC VIÊN THAM GIA LỚP HỌC/HỘI NGHỊ";
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 15;
                        worksheet.Range("A1:D1").Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        worksheet.Cell(2, 1).Value = $"Lớp/Hội nghị: {lblClassName.Text}";
                        worksheet.Cell(2, 1).Style.Font.Bold = true;
                        worksheet.Range("A2:D2").Merge();

                        worksheet.Cell(3, 1).Value = $"Thời gian tham gia: {lblParticipationDate.Text}   |   Số QĐ: {lblDecisionNumber.Text}   |   Đơn vị QĐ: {lblDecisionUnit.Text}";
                        worksheet.Range("A3:D3").Merge();

                        // 2. Table Headers
                        string[] headers = { "STT", "Họ và tên", "Ngày tháng năm sinh", "Bộ phận công tác" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(5, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                        worksheet.Row(5).Height = 25;

                        // 3. Data Rows
                        int startRow = 6;
                        int stt = 1;
                        foreach (var p in _participants)
                        {
                            worksheet.Cell(startRow, 1).Value = stt++;
                            worksheet.Cell(startRow, 2).Value = p.FullName;
                            worksheet.Cell(startRow, 3).Value = p.DateOfBirth?.ToString("dd/MM/yyyy") ?? "";
                            worksheet.Cell(startRow, 4).Value = p.Department;

                            // Formats and borders
                            for (int i = 1; i <= 4; i++)
                            {
                                var cell = worksheet.Cell(startRow, i);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                if (i == 1 || i == 3) cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            startRow++;
                        }

                        // 4. Widths
                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 30;
                        worksheet.Column(3).Width = 25;
                        worksheet.Column(4).Width = 30;

                        workbook.SaveAs(saveFileDialog.FileName);

                        var success = new SuccessWindow("Xuất danh sách học viên thành công!", null, saveFileDialog.FileName, true);
                        success.Owner = this;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Lỗi xuất Excel: {ex.Message}", "Lỗi");
                    warning.Owner = this;
                    warning.ShowDialog();
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = _isDataChanged;
            Close();
        }

        public class PersonnelComboItem
        {
            public int Id { get; set; }
            public string FullName { get; set; } = string.Empty;
            public string Department { get; set; } = string.Empty;
            public string IdentityCardNumber { get; set; } = string.Empty;
        }
    }
}
