using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class PlanningRecordDialog : Window
    {
        private Personnel? _selectedPersonnel;
        public PlanningRecord? ResultRecord { get; private set; }
        private readonly string _planningType;
        private readonly bool _isEditMode;
        private readonly int _recordId;

        // Constructor cho chế độ thêm mới
        public PlanningRecordDialog(string planningType)
        {
            InitializeComponent();
            _planningType = planningType;
            _isEditMode = false;
            txtTitle.Text = $"Thêm Quy hoạch {_planningType}";
            
            // Đặt mặc định ngày ký là hôm nay
            dpDecisionDate.SelectedDate = DateTime.Today;

            // Load danh sách nhiệm kỳ quy hoạch
            LoadPlanningTerms();
            // Load danh sách chức vụ
            LoadPositions();
        }

        // Constructor cho chế độ chỉnh sửa
        public PlanningRecordDialog(PlanningRecord record)
        {
            InitializeComponent();
            _planningType = record.PlanningType;
            _isEditMode = true;
            _recordId = record.Id;
            txtTitle.Text = $"Cập nhật Quy hoạch {_planningType}";

            // Tải thông tin bản ghi lên giao diện
            _selectedPersonnel = record.Personnel;
            if (_selectedPersonnel != null)
            {
                BindPersonnel(_selectedPersonnel);
            }

            // Bind các trường nhập liệu
            SetComboBoxValue(cboStatus, record.Status);
            LoadPlanningTerms(record.PlanningTerm);
            LoadPositions(record.PlannedPosition, record.PlannedTransitionPosition);
            
            // Hiển thị trình độ lưu trong bản ghi, nếu trống thì lấy từ Personnel
            txtTrainingLevel.Text = !string.IsNullOrEmpty(record.TrainingLevel) ? record.TrainingLevel : (_selectedPersonnel?.EducationLevel ?? "");
            txtPoliticalTheoryLevel.Text = !string.IsNullOrEmpty(record.PoliticalTheoryLevel) ? record.PoliticalTheoryLevel : (_selectedPersonnel?.PoliticalTheoryLevel ?? "");
            
            txtDecisionNumber.Text = record.DecisionNumber;
            dpDecisionDate.SelectedDate = record.DecisionDate;
            txtDecisionUnit.Text = record.DecisionUnit;
            if (!string.IsNullOrEmpty(record.Evaluation3Years))
            {
                txtEvaluation3Years.Text = System.Text.RegularExpressions.Regex.Replace(record.Evaluation3Years, @",\s*(?=\d{4}:)", Environment.NewLine);
            }
            else
            {
                txtEvaluation3Years.Text = GetEvaluation3YearsText(record.PersonnelId);
            }
            txtNote.Text = record.Note;

            // Chế độ chỉnh sửa thì không cho phép thay đổi cán bộ công chức để tránh nhầm lẫn
            btnSelectPersonnel.IsEnabled = false;
        }

        private void BindPersonnel(Personnel p)
        {
            txtNoPersonnelSelected.Visibility = Visibility.Collapsed;
            pnlSelectedPersonnel.Visibility = Visibility.Visible;
            txtPersonnelName.Text = p.FullName;
            txtPersonnelCardNumber.Text = p.IdentityCardNumber ?? "Không có";
            txtPersonnelDepartment.Text = p.Department ?? "Không có bộ phận";
            txtPersonnelPosition.Text = p.Position ?? "Không có chức vụ";

            // Nếu là thêm mới, tự động điền các thông tin trình độ từ Personnel
            if (!_isEditMode)
            {
                txtTrainingLevel.Text = p.EducationLevel ?? "";
                txtPoliticalTheoryLevel.Text = p.PoliticalTheoryLevel ?? "";
                txtEvaluation3Years.Text = GetEvaluation3YearsText(p.Id);
            }
        }

        private void SetComboBoxValue(ComboBox cbo, string value)
        {
            foreach (ComboBoxItem item in cbo.Items)
            {
                if (item.Content.ToString() == value)
                {
                    cbo.SelectedItem = item;
                    break;
                }
            }
        }

        private void LoadPlanningTerms(string? selectedTerm = null)
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var terms = db.PlanningTerms
                                  .OrderByDescending(t => t.TermName)
                                  .Select(t => t.TermName)
                                  .ToList();

                    if (!string.IsNullOrEmpty(selectedTerm) && !terms.Contains(selectedTerm))
                    {
                        terms.Add(selectedTerm);
                    }

                    cboPlanningTerm.ItemsSource = terms;
                    if (!string.IsNullOrEmpty(selectedTerm))
                    {
                        cboPlanningTerm.SelectedItem = selectedTerm;
                    }
                    else
                    {
                        cboPlanningTerm.SelectedIndex = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh mục nhiệm kỳ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnAddPlanningTerm_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var termDialog = new PlanningTermManagementDialog();
                termDialog.Owner = this;
                if (termDialog.ShowDialog() == true)
                {
                    // Sau khi đóng form, load lại danh sách từ DB
                    string? selected = termDialog.SelectedTermName;
                    if (!string.IsNullOrEmpty(selected))
                    {
                        LoadPlanningTerms(selected);
                    }
                    else
                    {
                        string? currentVal = cboPlanningTerm.SelectedItem as string;
                        LoadPlanningTerms(currentVal);
                    }
                }
                else
                {
                    // Reload lại kể cả cancel vì có thể người dùng đã sửa/xóa các nhiệm kỳ khác
                    string? currentVal = cboPlanningTerm.SelectedItem as string;
                    LoadPlanningTerms(currentVal);
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi quản lý nhiệm kỳ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void LoadPositions(string? selectedPlanned = null, string? selectedTransition = null)
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var posOrder = new System.Collections.Generic.List<string> {
                        "Chi cục trưởng",
                        "Quyền Chi cục trưởng",
                        "Phó Chi cục trưởng",
                        "Trưởng Thuế cơ sở",
                        "Quyền Trưởng Thuế cơ sở",
                        "Phó Trưởng Thuế cơ sở",
                        "Đội trưởng",
                        "Trưởng phòng",
                        "Phó Đội trưởng",
                        "Phó Trưởng phòng",
                        "Tổ trưởng",
                        "Phó Tổ trưởng",
                        "Công chức",
                        "Nhân viên"
                    };

                    var positions = db.Positions
                                      .Select(p => p.Name)
                                      .Distinct()
                                      .ToList()
                                      .OrderBy(name =>
                                      {
                                          int idx = posOrder.FindIndex(p => p.Equals(name, System.StringComparison.OrdinalIgnoreCase));
                                          return idx == -1 ? 999 : idx;
                                      })
                                      .ThenBy(name => name)
                                      .ToList();

                    cboPlannedPosition.ItemsSource = new List<string>(positions);
                    cboPlannedTransitionPosition.ItemsSource = new List<string>(positions);

                    if (!string.IsNullOrEmpty(selectedPlanned))
                    {
                        cboPlannedPosition.Text = selectedPlanned;
                    }
                    if (!string.IsNullOrEmpty(selectedTransition))
                    {
                        cboPlannedTransitionPosition.Text = selectedTransition;
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh mục chức vụ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnAddPlannedPosition_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new AddPositionDialog();
                dialog.Owner = this;
                if (dialog.ShowDialog() == true)
                {
                    string? selected = dialog.SelectedPosition;
                    string? transitionVal = cboPlannedTransitionPosition.SelectedItem as string ?? cboPlannedTransitionPosition.Text.Trim();
                    LoadPositions(selected, transitionVal);
                }
                else
                {
                    string? plannedVal = cboPlannedPosition.SelectedItem as string ?? cboPlannedPosition.Text.Trim();
                    string? transitionVal = cboPlannedTransitionPosition.SelectedItem as string ?? cboPlannedTransitionPosition.Text.Trim();
                    LoadPositions(plannedVal, transitionVal);
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi quản lý chức vụ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnAddPlannedTransitionPosition_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new AddPositionDialog();
                dialog.Owner = this;
                if (dialog.ShowDialog() == true)
                {
                    string? selected = dialog.SelectedPosition;
                    string? plannedVal = cboPlannedPosition.SelectedItem as string ?? cboPlannedPosition.Text.Trim();
                    LoadPositions(plannedVal, selected);
                }
                else
                {
                    string? plannedVal = cboPlannedPosition.SelectedItem as string ?? cboPlannedPosition.Text.Trim();
                    string? transitionVal = cboPlannedTransitionPosition.SelectedItem as string ?? cboPlannedTransitionPosition.Text.Trim();
                    LoadPositions(plannedVal, transitionVal);
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi quản lý chức vụ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnSelectPersonnel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Mở dialog chọn cán bộ (không loại trừ ai nên truyền danh sách trống)
                var selector = new PersonnelSelectorDialog(new List<int>(), isSingleSelect: true);
                selector.Owner = this;
                if (selector.ShowDialog() == true && selector.SelectedPersonnelIds.Any())
                {
                    int selectedId = selector.SelectedPersonnelIds.First();
                    using (var db = new AppDbContext())
                    {
                        var p = db.Personnel.Find(selectedId);
                        if (p != null)
                        {
                            _selectedPersonnel = p;
                            BindPersonnel(_selectedPersonnel);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi chọn cán bộ công chức: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            // 1. Kiểm tra tính hợp lệ của dữ liệu (Validation)
            if (_selectedPersonnel == null)
            {
                var warning = new WarningWindow("Vui lòng chọn một cán bộ công chức!", "Cảnh báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            string status = (cboStatus.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Tiếp tục quy hoạch";
            string term = (cboPlanningTerm.SelectedItem as string) ?? cboPlanningTerm.Text.Trim();

            if (string.IsNullOrWhiteSpace(term))
            {
                var warning = new WarningWindow("Vui lòng chọn hoặc thêm nhiệm kỳ quy hoạch!", "Cảnh báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            // 2. Tạo đối tượng kết quả
            ResultRecord = new PlanningRecord
            {
                Id = _isEditMode ? _recordId : 0,
                PersonnelId = _selectedPersonnel.Id,
                PlanningType = _planningType,
                Status = status,
                PlanningTerm = term,
                CurrentPosition = _selectedPersonnel.Position,
                PlannedPosition = (cboPlannedPosition.SelectedItem as string) ?? cboPlannedPosition.Text.Trim(),
                PlannedTransitionPosition = (cboPlannedTransitionPosition.SelectedItem as string) ?? cboPlannedTransitionPosition.Text.Trim(),
                TrainingLevel = txtTrainingLevel.Text.Trim(),
                PoliticalTheoryLevel = txtPoliticalTheoryLevel.Text.Trim(),
                DecisionNumber = txtDecisionNumber.Text.Trim(),
                DecisionDate = dpDecisionDate.SelectedDate,
                DecisionUnit = txtDecisionUnit.Text.Trim(),
                Evaluation3Years = txtEvaluation3Years.Text.Trim(),
                Note = txtNote.Text.Trim()
            };

            DialogResult = true;
            Close();
        }

        private string GetEvaluation3YearsText(int personnelId)
        {
            try
            {
                int currentYear = DateTime.Today.Year;
                int y1 = currentYear - 3;
                int y2 = currentYear - 2;
                int y3 = currentYear - 1;

                using (var db = new AppDbContext())
                {
                    var evals = db.EvaluationRecords
                                  .Where(e => e.PersonnelId == personnelId && e.Year >= y1 && e.Year <= y3)
                                  .ToList();

                    var listResults = new List<string>();
                    for (int y = y1; y <= y3; y++)
                    {
                        var ev = evals.FirstOrDefault(e => e.Year == y);
                        string ratingText = ev != null ? ShortenRating(ev.Rating) : "Chưa đánh giá";
                        listResults.Add($"{y}: {ratingText}");
                    }

                    return string.Join(Environment.NewLine, listResults);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi lấy xếp loại 3 năm gần nhất: {ex.Message}");
                return string.Empty;
            }
        }

        private string ShortenRating(string rating)
        {
            if (string.IsNullOrEmpty(rating)) return "Chưa đánh giá";
            if (rating.Equals("Hoàn thành xuất sắc nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                return "Hoàn thành xuất sắc";
            if (rating.Equals("Hoàn thành tốt nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                return "Hoàn thành tốt";
            if (rating.Equals("Hoàn thành nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                return "Hoàn thành nhiệm vụ";
            if (rating.Equals("Không hoàn thành nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                return "Không hoàn thành";
            return rating;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Cho phép di chuyển cửa sổ bằng cách kéo Header
        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
    }
}
