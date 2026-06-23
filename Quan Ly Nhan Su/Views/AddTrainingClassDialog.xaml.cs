using System;
using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class AddTrainingClassDialog : Window
    {
        private int? _classId;
        private TrainingClass? _trainingClass;

        public AddTrainingClassDialog(int? classId = null)
        {
            InitializeComponent();
            _classId = classId;
            LoadData();
        }

        private void LoadData()
        {
            if (_classId.HasValue)
            {
                txtTitle.Text = "Sửa Lớp học / Hội nghị";
                try
                {
                    using (var db = new AppDbContext())
                    {
                        _trainingClass = db.TrainingClasses.Find(_classId.Value);
                        if (_trainingClass != null)
                        {
                            txtClassName.Text = _trainingClass.ClassName;
                            dpParticipationDate.SelectedDate = _trainingClass.ParticipationDate;
                            txtDecisionNumber.Text = _trainingClass.DecisionNumber;
                            dpDecisionDate.SelectedDate = _trainingClass.DecisionDate;
                            txtDecisionUnit.Text = _trainingClass.DecisionUnit;
                        }
                    }
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Lỗi tải thông tin lớp học: {ex.Message}", "Lỗi");
                    warning.Owner = this;
                    warning.ShowDialog();
                }
            }
            else
            {
                txtTitle.Text = "Thêm Lớp học / Hội nghị";
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string className = txtClassName.Text.Trim();
            if (string.IsNullOrEmpty(className))
            {
                var warning = new WarningWindow("Vui lòng nhập tên lớp học hoặc hội nghị!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            try
            {
                using (var db = new AppDbContext())
                {
                    if (_classId.HasValue)
                    {
                        var tc = db.TrainingClasses.Find(_classId.Value);
                        if (tc != null)
                        {
                            tc.ClassName = className;
                            tc.ParticipationDate = dpParticipationDate.SelectedDate;
                            tc.DecisionNumber = txtDecisionNumber.Text.Trim();
                            tc.DecisionDate = dpDecisionDate.SelectedDate;
                            tc.DecisionUnit = txtDecisionUnit.Text.Trim();
                            db.Entry(tc).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                        }
                    }
                    else
                    {
                        var tc = new TrainingClass
                        {
                            ClassName = className,
                            ParticipationDate = dpParticipationDate.SelectedDate,
                            DecisionNumber = txtDecisionNumber.Text.Trim(),
                            DecisionDate = dpDecisionDate.SelectedDate,
                            DecisionUnit = txtDecisionUnit.Text.Trim()
                        };
                        db.TrainingClasses.Add(tc);
                    }

                    db.SaveChanges();
                    DialogResult = true;
                    Close();
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi lưu thông tin lớp học: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
