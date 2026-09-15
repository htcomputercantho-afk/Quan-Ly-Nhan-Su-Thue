using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class PersonnelDegreesDialog : Window
    {
        private readonly Personnel _personnel;
        private List<PersonnelDegree> _allDegrees = new();
        private PersonnelDegree? _editingDegree = null;
        private readonly string? _initialTab;

        /// <summary>
        /// Đánh dấu đã có thay đổi dữ liệu để form cha biết và làm mới giao diện ngoài.
        /// </summary>
        public bool HasChanges { get; private set; } = false;

        public PersonnelDegreesDialog(Personnel personnel, string? initialTab = null)
        {
            InitializeComponent();
            _personnel = personnel ?? throw new ArgumentNullException(nameof(personnel));
            _initialTab = initialTab;
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtPersonnelName.Text = string.IsNullOrWhiteSpace(_personnel.FullName) ? "Cán bộ mới" : _personnel.FullName;
            txtPersonnelDept.Text = string.IsNullOrEmpty(_personnel.Department) ? "---" : _personnel.Department;
            txtPersonnelPosition.Text = string.IsNullOrEmpty(_personnel.Position) ? "---" : _personnel.Position;
            txtDialogTitle.Text = $"DANH SÁCH VĂN BẰNG & CHỨNG CHỈ — {(string.IsNullOrWhiteSpace(_personnel.FullName) ? "CÁN BỘ MỚI" : _personnel.FullName.ToUpper())}";

            EnsureSeedFromExistingData();
            PopulateDegreeNameSuggestions();

            // Kích hoạt tab ban đầu theo yêu cầu mở từ form cha
            if (!string.IsNullOrEmpty(_initialTab))
            {
                switch (_initialTab)
                {
                    case "Tin học":
                        rbTabIT.IsChecked = true;
                        break;
                    case "Ngoại ngữ":
                        rbTabLang.IsChecked = true;
                        break;
                    case "Chuyên môn":
                        rbTabMajor.IsChecked = true;
                        break;
                    case "Quản lý Nhà nước":
                        rbTabState.IsChecked = true;
                        break;
                    case "Lý luận chính trị":
                        rbTabPolTheory.IsChecked = true;
                        break;
                    case "Chứng chỉ khác":
                        rbTabOther.IsChecked = true;
                        break;
                    default:
                        rbTabAll.IsChecked = true;
                        break;
                }
            }
            else
            {
                rbTabAll.IsChecked = true;
            }

            LoadDegrees();
        }

        /// <summary>
        /// Tự động kế thừa các bằng cấp đã có sẵn trong hồ sơ cán bộ (nếu chưa có bản ghi nào trong PersonnelDegrees).
        /// Hỗ trợ cả cán bộ đã lưu trong DB và cán bộ mới tạo trong bộ nhớ.
        /// </summary>
        private void EnsureSeedFromExistingData()
        {
            if (_personnel.Id <= 0)
            {
                _personnel.PersonnelDegrees ??= new List<PersonnelDegree>();
                if (!_personnel.PersonnelDegrees.Any())
                {
                    int tempId = -1;

                    if (!string.IsNullOrWhiteSpace(_personnel.EducationLevel))
                    {
                        _personnel.PersonnelDegrees.Add(new PersonnelDegree
                        {
                            Id = tempId--,
                            DegreeType = "Chuyên môn",
                            DegreeName = _personnel.EducationLevel.Trim(),
                            Major = _personnel.Major,
                            Institution = _personnel.University,
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.ITSkillLevel))
                    {
                        _personnel.PersonnelDegrees.Add(new PersonnelDegree
                        {
                            Id = tempId--,
                            DegreeType = "Tin học",
                            DegreeName = _personnel.ITSkillLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.LanguageSkillLevel))
                    {
                        _personnel.PersonnelDegrees.Add(new PersonnelDegree
                        {
                            Id = tempId--,
                            DegreeType = "Ngoại ngữ",
                            DegreeName = _personnel.LanguageSkillLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.StateManagementLevel))
                    {
                        _personnel.PersonnelDegrees.Add(new PersonnelDegree
                        {
                            Id = tempId--,
                            DegreeType = "Quản lý Nhà nước",
                            DegreeName = _personnel.StateManagementLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.PoliticalTheoryLevel))
                    {
                        _personnel.PersonnelDegrees.Add(new PersonnelDegree
                        {
                            Id = tempId--,
                            DegreeType = "Lý luận chính trị",
                            DegreeName = _personnel.PoliticalTheoryLevel.Trim(),
                            IsPrimary = true
                        });
                    }
                }
                return;
            }

            try
            {
                using var db = new AppDbContext();
                bool hasAny = db.PersonnelDegrees.Any(d => d.PersonnelId == _personnel.Id);
                if (!hasAny)
                {
                    var seeded = new List<PersonnelDegree>();

                    if (!string.IsNullOrWhiteSpace(_personnel.EducationLevel))
                    {
                        seeded.Add(new PersonnelDegree
                        {
                            PersonnelId = _personnel.Id,
                            DegreeType = "Chuyên môn",
                            DegreeName = _personnel.EducationLevel.Trim(),
                            Major = _personnel.Major,
                            Institution = _personnel.University,
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.ITSkillLevel))
                    {
                        seeded.Add(new PersonnelDegree
                        {
                            PersonnelId = _personnel.Id,
                            DegreeType = "Tin học",
                            DegreeName = _personnel.ITSkillLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.LanguageSkillLevel))
                    {
                        seeded.Add(new PersonnelDegree
                        {
                            PersonnelId = _personnel.Id,
                            DegreeType = "Ngoại ngữ",
                            DegreeName = _personnel.LanguageSkillLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.StateManagementLevel))
                    {
                        seeded.Add(new PersonnelDegree
                        {
                            PersonnelId = _personnel.Id,
                            DegreeType = "Quản lý Nhà nước",
                            DegreeName = _personnel.StateManagementLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(_personnel.PoliticalTheoryLevel))
                    {
                        seeded.Add(new PersonnelDegree
                        {
                            PersonnelId = _personnel.Id,
                            DegreeType = "Lý luận chính trị",
                            DegreeName = _personnel.PoliticalTheoryLevel.Trim(),
                            IsPrimary = true
                        });
                    }

                    if (seeded.Any())
                    {
                        db.PersonnelDegrees.AddRange(seeded);
                        db.SaveChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error seeding degrees: " + ex.Message);
            }
        }

        /// <summary>
        /// Thứ tự ưu tiên hiển thị theo nhóm Loại bằng:
        /// 1. Chuyên môn -> 2. Tin học -> 3. Ngoại ngữ -> 4. Quản lý Nhà nước -> 5. Lý luận chính trị -> 6. Chứng chỉ khác.
        /// </summary>
        private static int GetCategoryRank(string? degreeType)
        {
            if (string.IsNullOrWhiteSpace(degreeType)) return 99;
            string type = degreeType.Trim().ToLowerInvariant();

            if (type.Contains("chuyên môn")) return 1;
            if (type.Contains("tin học")) return 2;
            if (type.Contains("ngoại ngữ")) return 3;
            if (type.Contains("quản lý nhà nước") || type.Contains("ql nhà nước") || type.Contains("qlnn")) return 4;
            if (type.Contains("lý luận chính trị") || type.Contains("llct")) return 5;
            if (type.Contains("chứng chỉ khác") || type.Contains("khác")) return 6;

            return 7;
        }

        /// <summary>
        /// Phân cấp trình độ chuyên môn từ cao xuống thấp:
        /// 1. Tiến sĩ (Tiến sĩ, Tiến sỹ, TS, TSKH, PhD...)
        /// 2. Thạc sĩ (Thạc sĩ, Thạc sỹ, ThS, Master, Cao học...)
        /// 3. Đại học (Đại học, ĐH, Cử nhân, Kỹ sư, Bác sĩ, Dược sĩ, Kiến trúc sư...)
        /// 4. Cao đẳng (Cao đẳng, CĐ...)
        /// 5. Trung cấp (Trung cấp, TC, Trung học chuyên nghiệp...)
        /// 6. Sơ cấp (Sơ cấp, Chứng chỉ nghề, Đào tạo nghề...)
        /// 7. Khác
        /// </summary>
        private static int GetEducationLevelRank(string? degreeName)
        {
            if (string.IsNullOrWhiteSpace(degreeName)) return 99;
            string name = degreeName.Trim().ToLowerInvariant();

            // 1. Tiến sĩ
            if (name.Contains("tiến sĩ") || name.Contains("tiến sỹ") || name.Contains("tien si") || name.Contains("tien sy") ||
                name.Contains("tskh") || name == "ts" || name.StartsWith("ts.") || name.StartsWith("ts ") || name.EndsWith(" ts") ||
                name.Contains("phd") || name.Contains("doctor"))
            {
                return 1;
            }

            // 2. Thạc sĩ
            if (name.Contains("thạc sĩ") || name.Contains("thạc sỹ") || name.Contains("thac si") || name.Contains("thac sy") ||
                name.Contains("cao học") || name.Contains("cao hoc") || name == "ths" || name.StartsWith("ths.") ||
                name.StartsWith("ths ") || name.EndsWith(" ths") || name.Contains("master"))
            {
                return 2;
            }

            // 3. Đại học / Cử nhân / Kỹ sư
            if (name.Contains("đại học") || name.Contains("dai hoc") || name.Contains("cử nhân") || name.Contains("cu nhan") ||
                name.Contains("kỹ sư") || name.Contains("ky su") || name.Contains("bác sĩ") || name.Contains("bác sỹ") ||
                name.Contains("dược sĩ") || name.Contains("dược sỹ") || name.Contains("kiến trúc sư") ||
                name == "đh" || name.StartsWith("đh ") || name.StartsWith("đh-") || name.StartsWith("đh.") || name.EndsWith(" đh"))
            {
                return 3;
            }

            // 4. Cao đẳng
            if (name.Contains("cao đẳng") || name.Contains("cao dang") ||
                name == "cđ" || name.StartsWith("cđ ") || name.StartsWith("cđ-") || name.StartsWith("cđ.") || name.EndsWith(" cđ"))
            {
                return 4;
            }

            // 5. Trung cấp
            if (name.Contains("trung cấp") || name.Contains("trung cap") || name.Contains("trung học chuyên nghiệp") ||
                name == "tc" || name.StartsWith("tc ") || name.StartsWith("tc-") || name.StartsWith("tc.") || name.EndsWith(" tc"))
            {
                return 5;
            }

            // 6. Sơ cấp
            if (name.Contains("sơ cấp") || name.Contains("so cap") || name.Contains("chứng chỉ nghề") || name.Contains("đào tạo nghề"))
            {
                return 6;
            }

            return 7;
        }

        /// <summary>
        /// Trích xuất năm cấp / tốt nghiệp từ GraduationYear hoặc IssueDate.
        /// </summary>
        private static int ParseGraduationYear(string? yearStr, DateTime? issueDate = null)
        {
            if (!string.IsNullOrWhiteSpace(yearStr))
            {
                string s = yearStr.Trim();
                if (int.TryParse(s, out int exactYear) && exactYear >= 1900 && exactYear <= 2100)
                {
                    return exactYear;
                }

                var match = System.Text.RegularExpressions.Regex.Match(s, @"\b(19\d\d|20\d\d)\b");
                if (match.Success && int.TryParse(match.Value, out int regexYear))
                {
                    return regexYear;
                }
            }

            if (issueDate.HasValue && issueDate.Value.Year >= 1900 && issueDate.Value.Year <= 2100)
            {
                return issueDate.Value.Year;
            }

            return 0;
        }

        private static bool IsMajorCategory(string? degreeType)
        {
            if (string.IsNullOrWhiteSpace(degreeType)) return false;
            return degreeType.Trim().ToLowerInvariant().Contains("chuyên môn");
        }

        /// <summary>
        /// Sắp xếp danh sách văn bằng theo đúng quy tắc:
        /// 1. Nhóm Loại bằng: Chuyên môn -> Tin học -> Ngoại ngữ -> QL Nhà nước -> Lý luận chính trị -> Chứng chỉ khác.
        /// 2. Trong chuyên môn: Sắp xếp theo trình độ từ cao xuống thấp (Tiến sĩ -> Thạc sĩ -> Đại học -> Cao đẳng -> Trung cấp -> Sơ cấp).
        ///    Nếu cùng trình độ: Năm cấp gần nhất trước.
        /// 3. Các loại bằng còn lại (Tin học, Ngoại ngữ, QLNN, LLCT, Chứng chỉ khác): Sắp xếp theo Năm cấp gần nhất trước.
        /// </summary>
        private static IEnumerable<PersonnelDegree> SortDegrees(IEnumerable<PersonnelDegree> degrees)
        {
            return degrees
                .OrderBy(d => GetCategoryRank(d.DegreeType))
                .ThenBy(d => IsMajorCategory(d.DegreeType) ? GetEducationLevelRank(d.DegreeName) : 0)
                .ThenByDescending(d => ParseGraduationYear(d.GraduationYear, d.IssueDate))
                .ThenByDescending(d => d.IsPrimary)
                .ThenBy(d => d.Id);
        }

        /// <summary>
        /// Tải toàn bộ danh sách văn bằng từ CSDL hoặc bộ nhớ và làm mới bảng hiển thị.
        /// </summary>
        private void LoadDegrees()
        {
            if (_personnel.Id <= 0)
            {
                var list = _personnel.PersonnelDegrees?.ToList() ?? new List<PersonnelDegree>();
                _allDegrees = SortDegrees(list).ToList();
            }
            else
            {
                using var db = new AppDbContext();
                var list = db.PersonnelDegrees
                    .AsNoTracking()
                    .Where(d => d.PersonnelId == _personnel.Id)
                    .ToList();
                _allDegrees = SortDegrees(list).ToList();
            }

            ApplyFilterAndBind();
            UpdateSummary();
        }

        /// <summary>
        /// Lọc dữ liệu theo Tab hiện hành và gán STT.
        /// </summary>
        private void ApplyFilterAndBind()
        {
            IEnumerable<PersonnelDegree> filtered = _allDegrees;

            if (rbTabIT.IsChecked == true)
            {
                filtered = filtered.Where(d => d.DegreeType == "Tin học");
            }
            else if (rbTabLang.IsChecked == true)
            {
                filtered = filtered.Where(d => d.DegreeType == "Ngoại ngữ");
            }
            else if (rbTabMajor.IsChecked == true)
            {
                filtered = filtered.Where(d => d.DegreeType == "Chuyên môn");
            }
            else if (rbTabState.IsChecked == true)
            {
                filtered = filtered.Where(d => d.DegreeType == "Quản lý Nhà nước");
            }
            else if (rbTabPolTheory.IsChecked == true)
            {
                filtered = filtered.Where(d => d.DegreeType == "Lý luận chính trị");
            }
            else if (rbTabOther.IsChecked == true)
            {
                filtered = filtered.Where(d => d.DegreeType == "Chứng chỉ khác" ||
                                              (!new[] { "Tin học", "Ngoại ngữ", "Chuyên môn", "Quản lý Nhà nước", "Lý luận chính trị" }.Contains(d.DegreeType)));
            }

            var sortedList = SortDegrees(filtered).ToList();
            var viewList = sortedList.Select((d, idx) => new DegreeViewModel(d, idx + 1)).ToList();
            dgDegrees.ItemsSource = viewList;
        }

        private void UpdateSummary()
        {
            int total = _allDegrees.Count;
            int majorCount = _allDegrees.Count(d => d.DegreeType == "Chuyên môn");
            int itCount = _allDegrees.Count(d => d.DegreeType == "Tin học");
            int langCount = _allDegrees.Count(d => d.DegreeType == "Ngoại ngữ");
            int stateCount = _allDegrees.Count(d => d.DegreeType == "Quản lý Nhà nước");
            int polCount = _allDegrees.Count(d => d.DegreeType == "Lý luận chính trị");
            int otherCount = _allDegrees.Count(d => d.DegreeType == "Chứng chỉ khác" ||
                                                   (!new[] { "Tin học", "Ngoại ngữ", "Chuyên môn", "Quản lý Nhà nước", "Lý luận chính trị" }.Contains(d.DegreeType)));

            txtDegreeSummary.Text = otherCount > 0
                ? $"Tổng số: {total} văn bằng, chứng chỉ ({majorCount} Chuyên môn, {itCount} Tin học, {langCount} Ngoại ngữ, {stateCount} QLNN, {polCount} LLCT, {otherCount} CC khác)"
                : $"Tổng số: {total} văn bằng, chứng chỉ ({majorCount} Chuyên môn, {itCount} Tin học, {langCount} Ngoại ngữ, {stateCount} QLNN, {polCount} LLCT)";
        }

        private void TabFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;

            // Đồng bộ loại bằng mặc định trên form nhập liệu theo tab đang chọn
            if (rbTabMajor.IsChecked == true)
                SelectComboBoxItemByText(cboDegreeType, "Chuyên môn");
            else if (rbTabIT.IsChecked == true)
                SelectComboBoxItemByText(cboDegreeType, "Tin học");
            else if (rbTabLang.IsChecked == true)
                SelectComboBoxItemByText(cboDegreeType, "Ngoại ngữ");
            else if (rbTabState.IsChecked == true)
                SelectComboBoxItemByText(cboDegreeType, "Quản lý Nhà nước");
            else if (rbTabPolTheory.IsChecked == true)
                SelectComboBoxItemByText(cboDegreeType, "Lý luận chính trị");
            else if (rbTabOther.IsChecked == true)
                SelectComboBoxItemByText(cboDegreeType, "Chứng chỉ khác");

            ApplyFilterAndBind();
        }

        private void cboDegreeType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PopulateDegreeNameSuggestions();
        }

        private void dgDegrees_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgDegrees != null && dgDegrees.SelectedItem != null)
            {
                dgDegrees.SelectedItem = null;
            }
        }

        private void dgDegrees_RequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
        {
            // Ngăn chặn các phần tử con hoặc thay đổi focus tự động kéo thanh cuộn nhảy lên đầu
            e.Handled = true;
        }

        /// <summary>
        /// Nạp các gợi ý bằng cấp/chứng chỉ tương ứng theo Loại bằng được chọn.
        /// </summary>
        private void PopulateDegreeNameSuggestions()
        {
            if (cboDegreeName == null || cboDegreeType == null) return;

            string selectedType = (cboDegreeType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Chuyên môn";
            string currentText = cboDegreeName.Text;

            cboDegreeName.Items.Clear();

            string[] suggestions = selectedType switch
            {
                "Chuyên môn" => new[] { "Tiến sĩ", "Thạc sĩ", "Cử nhân", "Kỹ sư", "Cao đẳng", "Trung cấp", "Sơ cấp", "Đại học văn bằng 2" },
                "Tin học" => new[] { "Bằng A", "Bằng B", "Bằng C", "Ứng dụng CNTT cơ bản", "Ứng dụng CNTT nâng cao", "IC3", "MOS", "Tin học văn phòng", "Cử nhân CNTT" },
                "Ngoại ngữ" => new[] { "Bằng A", "Bằng B", "Bằng C", "Aptis B1", "Aptis B2", "IELTS 5.5", "IELTS 6.0", "IELTS 6.5", "IELTS 7.0", "TOEIC 500", "TOEIC 650", "TOEIC 750", "Khung 6 bậc B1", "Khung 6 bậc B2", "Tiếng Pháp B", "Tiếng Trung B" },
                "Quản lý Nhà nước" => new[] { "Chuyên viên", "Chuyên viên chính", "Chuyên viên cao cấp", "Lãnh đạo cấp phòng", "Bồi dưỡng ngạch kiểm soát viên" },
                "Lý luận chính trị" => new[] { "Sơ cấp", "Trung cấp", "Cao cấp", "Cử nhân chính trị" },
                _ => new[] { "Chứng chỉ đào tạo nghề", "Chứng chỉ bồi dưỡng kỹ năng", "Bằng khen / Giấy chứng nhận" }
            };

            foreach (var s in suggestions)
            {
                cboDegreeName.Items.Add(s);
            }

            cboDegreeName.Text = currentText;
        }

        /// <summary>
        /// Lưu văn bằng (Thêm mới hoặc Cập nhật bản ghi đang sửa).
        /// </summary>
        private void btnSaveDegree_Click(object sender, RoutedEventArgs e)
        {
            string type = (cboDegreeType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Chuyên môn";
            string name = cboDegreeName.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("Vui lòng nhập hoặc chọn Tên bằng / Chứng chỉ!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                cboDegreeName.Focus();
                return;
            }

            string? major = string.IsNullOrWhiteSpace(txtDegreeMajor.Text) ? null : txtDegreeMajor.Text.Trim();
            string? inst = string.IsNullOrWhiteSpace(txtInstitution.Text) ? null : txtInstitution.Text.Trim();
            string? classif = string.IsNullOrWhiteSpace(txtClassification.Text) ? null : txtClassification.Text.Trim();
            string? gradYear = string.IsNullOrWhiteSpace(txtGraduationYear.Text) ? null : txtGraduationYear.Text.Trim();
            bool isPrimary = chkIsPrimary.IsChecked == true;

            // Xử lý chế độ In-Memory cho cán bộ chưa lưu vào CSDL
            if (_personnel.Id <= 0)
            {
                _personnel.PersonnelDegrees ??= new List<PersonnelDegree>();

                if (_editingDegree == null)
                {
                    // THÊM MỚI IN-MEMORY
                    bool hasSameType = _allDegrees.Any(d => d.DegreeType == type);
                    if (!hasSameType)
                    {
                        isPrimary = true;
                    }

                    if (isPrimary)
                    {
                        foreach (var d in _allDegrees.Where(d => d.DegreeType == type))
                        {
                            d.IsPrimary = false;
                        }
                    }

                    int nextTempId = _allDegrees.Any() ? Math.Min(_allDegrees.Min(d => d.Id) - 1, -1) : -1;
                    var newDegree = new PersonnelDegree
                    {
                        Id = nextTempId,
                        PersonnelId = 0,
                        DegreeType = type,
                        DegreeName = name,
                        Major = major,
                        Institution = inst,
                        Classification = classif,
                        GraduationYear = gradYear,
                        IsPrimary = isPrimary
                    };

                    _allDegrees.Add(newDegree);
                    _personnel.PersonnelDegrees.Add(newDegree);

                    if (isPrimary)
                    {
                        SyncToPersonnelFields(type, name, major, inst);
                    }
                }
                else
                {
                    // CẬP NHẬT IN-MEMORY
                    var target = _allDegrees.FirstOrDefault(d => d.Id == _editingDegree.Id);
                    if (target != null)
                    {
                        if (isPrimary && !target.IsPrimary)
                        {
                            foreach (var d in _allDegrees.Where(d => d.DegreeType == type && d.Id != target.Id))
                            {
                                d.IsPrimary = false;
                            }
                        }

                        target.DegreeType = type;
                        target.DegreeName = name;
                        target.Major = major;
                        target.Institution = inst;
                        target.Classification = classif;
                        target.GraduationYear = gradYear;
                        target.IsPrimary = isPrimary;

                        // Đồng bộ lại vào list trong personnel
                        var pDegree = _personnel.PersonnelDegrees.FirstOrDefault(d => d.Id == target.Id);
                        if (pDegree != null)
                        {
                            pDegree.DegreeType = type;
                            pDegree.DegreeName = name;
                            pDegree.Major = major;
                            pDegree.Institution = inst;
                            pDegree.Classification = classif;
                            pDegree.GraduationYear = gradYear;
                            pDegree.IsPrimary = isPrimary;
                        }

                        if (isPrimary)
                        {
                            SyncToPersonnelFields(type, name, major, inst);
                        }
                    }
                }

                HasChanges = true;
                ResetForm();
                LoadDegrees();
                return;
            }

            // Xử lý chế độ CSDL trực tiếp cho cán bộ đã tồn tại
            try
            {
                using var db = new AppDbContext();

                if (_editingDegree == null)
                {
                    // THÊM MỚI VÀO CSDL
                    bool hasSameType = db.PersonnelDegrees.Any(d => d.PersonnelId == _personnel.Id && d.DegreeType == type);
                    if (!hasSameType)
                    {
                        isPrimary = true;
                    }

                    if (isPrimary)
                    {
                        var existingSameType = db.PersonnelDegrees
                            .Where(d => d.PersonnelId == _personnel.Id && d.DegreeType == type)
                            .ToList();
                        foreach (var d in existingSameType)
                        {
                            d.IsPrimary = false;
                        }
                    }

                    var newDegree = new PersonnelDegree
                    {
                        PersonnelId = _personnel.Id,
                        DegreeType = type,
                        DegreeName = name,
                        Major = major,
                        Institution = inst,
                        Classification = classif,
                        GraduationYear = gradYear,
                        IsPrimary = isPrimary
                    };

                    db.PersonnelDegrees.Add(newDegree);
                    db.SaveChanges();

                    if (isPrimary)
                    {
                        SyncToPersonnelFields(type, name, major, inst);
                    }
                }
                else
                {
                    // CẬP NHẬT TRONG CSDL
                    var entity = db.PersonnelDegrees.Find(_editingDegree.Id);
                    if (entity != null)
                    {
                        if (isPrimary && !entity.IsPrimary)
                        {
                            var existingSameType = db.PersonnelDegrees
                                .Where(d => d.PersonnelId == _personnel.Id && d.DegreeType == type && d.Id != entity.Id)
                                .ToList();
                            foreach (var d in existingSameType)
                            {
                                d.IsPrimary = false;
                            }
                        }

                        entity.DegreeType = type;
                        entity.DegreeName = name;
                        entity.Major = major;
                        entity.Institution = inst;
                        entity.Classification = classif;
                        entity.GraduationYear = gradYear;
                        entity.IsPrimary = isPrimary;

                        db.SaveChanges();

                        if (isPrimary)
                        {
                            SyncToPersonnelFields(type, name, major, inst);
                        }
                    }
                }

                HasChanges = true;
                ResetForm();
                LoadDegrees();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lưu văn bằng: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Đặt một bằng làm bằng chính (cao nhất) để hiển thị ngoài form.
        /// </summary>
        private void btnSetPrimary_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is DegreeViewModel vm)
            {
                // In-memory mode
                if (_personnel.Id <= 0)
                {
                    var target = _allDegrees.FirstOrDefault(d => d.Id == vm.Id);
                    if (target == null) return;

                    foreach (var d in _allDegrees.Where(d => d.DegreeType == target.DegreeType))
                    {
                        d.IsPrimary = (d.Id == target.Id);
                    }

                    if (_personnel.PersonnelDegrees != null)
                    {
                        foreach (var d in _personnel.PersonnelDegrees.Where(d => d.DegreeType == target.DegreeType))
                        {
                            d.IsPrimary = (d.Id == target.Id);
                        }
                    }

                    SyncToPersonnelFields(target.DegreeType, target.DegreeName, target.Major, target.Institution);
                    HasChanges = true;
                    LoadDegrees();
                    return;
                }

                try
                {
                    using var db = new AppDbContext();
                    var entity = db.PersonnelDegrees.Find(vm.Id);
                    if (entity == null) return;

                    var sameTypes = db.PersonnelDegrees
                        .Where(d => d.PersonnelId == _personnel.Id && d.DegreeType == entity.DegreeType)
                        .ToList();

                    foreach (var d in sameTypes)
                    {
                        d.IsPrimary = (d.Id == entity.Id);
                    }

                    db.SaveChanges();

                    SyncToPersonnelFields(entity.DegreeType, entity.DegreeName, entity.Major, entity.Institution);

                    HasChanges = true;
                    LoadDegrees();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi khi cập nhật bằng chính: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Chuyển form sang trạng thái sửa bản ghi được chọn.
        /// </summary>
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is DegreeViewModel vm)
            {
                var degree = _allDegrees.FirstOrDefault(d => d.Id == vm.Id);
                if (degree == null) return;

                _editingDegree = degree;

                SelectComboBoxItemByText(cboDegreeType, degree.DegreeType);
                cboDegreeName.Text = degree.DegreeName;
                txtDegreeMajor.Text = degree.Major ?? "";
                txtInstitution.Text = degree.Institution ?? "";
                txtClassification.Text = degree.Classification ?? "";
                txtGraduationYear.Text = degree.GraduationYear ?? "";
                chkIsPrimary.IsChecked = degree.IsPrimary;

                txtFormModeTitle.Text = $"Sửa văn bằng: {degree.DegreeName}";
                txtBtnSaveText.Text = "Cập nhật";
                iconFormMode.Kind = MaterialDesignThemes.Wpf.PackIconKind.Pencil;
                btnCancelEdit.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// Xóa văn bằng khỏi CSDL hoặc danh sách bộ nhớ.
        /// </summary>
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is DegreeViewModel vm)
            {
                var confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa văn bằng '{vm.DegreeName}' ({vm.DegreeType}) không?",
                    "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirm != MessageBoxResult.Yes) return;

                // In-memory mode
                if (_personnel.Id <= 0)
                {
                    var target = _allDegrees.FirstOrDefault(d => d.Id == vm.Id);
                    if (target != null)
                    {
                        bool wasPrimary = target.IsPrimary;
                        string degreeType = target.DegreeType;

                        _allDegrees.Remove(target);
                        _personnel.PersonnelDegrees?.Remove(target);

                        if (wasPrimary)
                        {
                            var remaining = _allDegrees
                                .Where(d => d.DegreeType == degreeType)
                                .LastOrDefault();

                            if (remaining != null)
                            {
                                remaining.IsPrimary = true;
                                SyncToPersonnelFields(remaining.DegreeType, remaining.DegreeName, remaining.Major, remaining.Institution);
                            }
                            else
                            {
                                SyncToPersonnelFields(degreeType, "", "", "");
                            }
                        }

                        HasChanges = true;
                        if (_editingDegree?.Id == vm.Id)
                        {
                            ResetForm();
                        }
                        LoadDegrees();
                    }
                    return;
                }

                try
                {
                    using var db = new AppDbContext();
                    var entity = db.PersonnelDegrees.Find(vm.Id);
                    if (entity != null)
                    {
                        bool wasPrimary = entity.IsPrimary;
                        string degreeType = entity.DegreeType;

                        db.PersonnelDegrees.Remove(entity);
                        db.SaveChanges();

                        if (wasPrimary)
                        {
                            var remaining = db.PersonnelDegrees
                                .Where(d => d.PersonnelId == _personnel.Id && d.DegreeType == degreeType)
                                .OrderByDescending(d => d.Id)
                                .FirstOrDefault();

                            if (remaining != null)
                            {
                                remaining.IsPrimary = true;
                                db.SaveChanges();
                                SyncToPersonnelFields(remaining.DegreeType, remaining.DegreeName, remaining.Major, remaining.Institution);
                            }
                            else
                            {
                                SyncToPersonnelFields(degreeType, "", "", "");
                            }
                        }

                        HasChanges = true;
                        if (_editingDegree?.Id == vm.Id)
                        {
                            ResetForm();
                        }
                        LoadDegrees();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi khi xóa văn bằng: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnCancelEdit_Click(object sender, RoutedEventArgs e)
        {
            ResetForm();
        }

        private void ResetForm()
        {
            _editingDegree = null;
            cboDegreeName.Text = "";
            txtDegreeMajor.Clear();
            txtInstitution.Clear();
            txtClassification.Clear();
            txtGraduationYear.Clear();
            chkIsPrimary.IsChecked = false;

            txtFormModeTitle.Text = "Thêm văn bằng, chứng chỉ mới";
            txtBtnSaveText.Text = "Thêm bằng";
            iconFormMode.Kind = MaterialDesignThemes.Wpf.PackIconKind.PlusCircleOutline;
            btnCancelEdit.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Đồng bộ thông tin bằng chính vào thực thể Personnel để cập nhật ra form cha.
        /// </summary>
        private void SyncToPersonnelFields(string degreeType, string degreeName, string? major, string? inst)
        {
            switch (degreeType)
            {
                case "Tin học":
                    _personnel.ITSkillLevel = degreeName;
                    break;
                case "Ngoại ngữ":
                    _personnel.LanguageSkillLevel = degreeName;
                    break;
                case "Chuyên môn":
                    _personnel.EducationLevel = degreeName;
                    if (!string.IsNullOrEmpty(major)) _personnel.Major = major;
                    if (!string.IsNullOrEmpty(inst)) _personnel.University = inst;
                    break;
                case "Quản lý Nhà nước":
                    _personnel.StateManagementLevel = degreeName;
                    break;
                case "Lý luận chính trị":
                    _personnel.PoliticalTheoryLevel = degreeName;
                    break;
            }

            // Đồng bộ trực tiếp vào database nếu Personnel đã có trong CSDL
            if (_personnel.Id > 0)
            {
                try
                {
                    using var db = new AppDbContext();
                    var p = db.Personnel.Find(_personnel.Id);
                    if (p != null)
                    {
                        switch (degreeType)
                        {
                            case "Tin học":
                                p.ITSkillLevel = degreeName;
                                break;
                            case "Ngoại ngữ":
                                p.LanguageSkillLevel = degreeName;
                                break;
                            case "Chuyên môn":
                                p.EducationLevel = degreeName;
                                if (!string.IsNullOrEmpty(major)) p.Major = major;
                                if (!string.IsNullOrEmpty(inst)) p.University = inst;
                                break;
                            case "Quản lý Nhà nước":
                                p.StateManagementLevel = degreeName;
                                break;
                            case "Lý luận chính trị":
                                p.PoliticalTheoryLevel = degreeName;
                                break;
                        }
                        db.SaveChanges();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error syncing to personnel db: " + ex.Message);
                }
            }
        }

        private static void SelectComboBoxItemByText(ComboBox cbo, string text)
        {
            foreach (var item in cbo.Items)
            {
                if (item is ComboBoxItem cbi && string.Equals(cbi.Content?.ToString(), text, StringComparison.OrdinalIgnoreCase))
                {
                    cbo.SelectedItem = cbi;
                    return;
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }

    /// <summary>
    /// ViewModel hỗ trợ hiển thị trên DataGrid có kèm STT, màu danh mục và hiển thị thân thiện.
    /// </summary>
    public class DegreeViewModel
    {
        private static readonly BrushConverter _brushConverter = new();

        private static readonly Brush BlueBg = (Brush)_brushConverter.ConvertFrom("#EFF6FF")!;
        private static readonly Brush BlueBorder = (Brush)_brushConverter.ConvertFrom("#BFDBFE")!;
        private static readonly Brush BlueText = (Brush)_brushConverter.ConvertFrom("#1D4ED8")!;

        private static readonly Brush PurpleBg = (Brush)_brushConverter.ConvertFrom("#F5F3FF")!;
        private static readonly Brush PurpleBorder = (Brush)_brushConverter.ConvertFrom("#DDD6FE")!;
        private static readonly Brush PurpleText = (Brush)_brushConverter.ConvertFrom("#6D28D9")!;

        private static readonly Brush GreenBg = (Brush)_brushConverter.ConvertFrom("#ECFDF5")!;
        private static readonly Brush GreenBorder = (Brush)_brushConverter.ConvertFrom("#A7F3D0")!;
        private static readonly Brush GreenText = (Brush)_brushConverter.ConvertFrom("#047857")!;

        private static readonly Brush AmberBg = (Brush)_brushConverter.ConvertFrom("#FFFBEB")!;
        private static readonly Brush AmberBorder = (Brush)_brushConverter.ConvertFrom("#FDE68A")!;
        private static readonly Brush AmberText = (Brush)_brushConverter.ConvertFrom("#B45309")!;

        private static readonly Brush RoseBg = (Brush)_brushConverter.ConvertFrom("#FFF1F2")!;
        private static readonly Brush RoseBorder = (Brush)_brushConverter.ConvertFrom("#FECDD3")!;
        private static readonly Brush RoseText = (Brush)_brushConverter.ConvertFrom("#BE123C")!;

        private static readonly Brush TealBg = (Brush)_brushConverter.ConvertFrom("#F0FDFA")!;
        private static readonly Brush TealBorder = (Brush)_brushConverter.ConvertFrom("#99F6E4")!;
        private static readonly Brush TealText = (Brush)_brushConverter.ConvertFrom("#0D9488")!;

        private static readonly Brush GrayBg = (Brush)_brushConverter.ConvertFrom("#F1F5F9")!;
        private static readonly Brush GrayBorder = (Brush)_brushConverter.ConvertFrom("#E2E8F0")!;
        private static readonly Brush GrayText = (Brush)_brushConverter.ConvertFrom("#64748B")!;
        private static readonly Brush DarkText = (Brush)_brushConverter.ConvertFrom("#334155")!;
        private static readonly Brush DashText = (Brush)_brushConverter.ConvertFrom("#94A3B8")!;

        private static readonly Brush StarActiveBg = (Brush)_brushConverter.ConvertFrom("#FEF3C7")!;
        private static readonly Brush StarActiveBorder = (Brush)_brushConverter.ConvertFrom("#FDE68A")!;
        private static readonly Brush StarActiveFg = (Brush)_brushConverter.ConvertFrom("#D97706")!;

        private static readonly Brush StarInactiveBg = (Brush)_brushConverter.ConvertFrom("#F8FAFC")!;
        private static readonly Brush StarInactiveBorder = (Brush)_brushConverter.ConvertFrom("#E2E8F0")!;
        private static readonly Brush StarInactiveFg = (Brush)_brushConverter.ConvertFrom("#94A3B8")!;

        public int STT { get; set; }
        public int Id { get; set; }
        public string DegreeType { get; set; }
        public string DegreeName { get; set; }
        public string? Major { get; set; }
        public string? Institution { get; set; }
        public string? Classification { get; set; }
        public string? GraduationYear { get; set; }
        public bool IsPrimary { get; set; }

        public Brush CategoryBgBrush => DegreeType switch
        {
            "Tin học" => BlueBg,
            "Ngoại ngữ" => PurpleBg,
            "Chuyên môn" => GreenBg,
            "Quản lý Nhà nước" => AmberBg,
            "Lý luận chính trị" => RoseBg,
            "Chứng chỉ khác" => TealBg,
            _ => GrayBg
        };

        public Brush CategoryBorderBrush => DegreeType switch
        {
            "Tin học" => BlueBorder,
            "Ngoại ngữ" => PurpleBorder,
            "Chuyên môn" => GreenBorder,
            "Quản lý Nhà nước" => AmberBorder,
            "Lý luận chính trị" => RoseBorder,
            "Chứng chỉ khác" => TealBorder,
            _ => GrayBorder
        };

        public Brush CategoryTextBrush => DegreeType switch
        {
            "Tin học" => BlueText,
            "Ngoại ngữ" => PurpleText,
            "Chuyên môn" => GreenText,
            "Quản lý Nhà nước" => AmberText,
            "Lý luận chính trị" => RoseText,
            "Chứng chỉ khác" => TealText,
            _ => GrayText
        };

        public PackIconKind CategoryIcon => DegreeType switch
        {
            "Tin học" => PackIconKind.Laptop,
            "Ngoại ngữ" => PackIconKind.Translate,
            "Chuyên môn" => PackIconKind.School,
            "Quản lý Nhà nước" => PackIconKind.TownHall,
            "Lý luận chính trị" => PackIconKind.BookOpenOutline,
            "Chứng chỉ khác" => PackIconKind.CertificateOutline,
            _ => PackIconKind.FileDocumentOutline
        };

        public string DisplayMajor => string.IsNullOrWhiteSpace(Major) ? "—" : Major.Trim();
        public string DisplayInstitution => string.IsNullOrWhiteSpace(Institution) ? "—" : Institution.Trim();
        public string DisplayClassification => string.IsNullOrWhiteSpace(Classification) ? "—" : Classification.Trim();
        public string DisplayGraduationYear => string.IsNullOrWhiteSpace(GraduationYear) ? "—" : GraduationYear.Trim();

        public Brush MajorForeground => string.IsNullOrWhiteSpace(Major) ? DashText : DarkText;
        public Brush InstitutionForeground => string.IsNullOrWhiteSpace(Institution) ? DashText : DarkText;
        public Brush ClassificationForeground => string.IsNullOrWhiteSpace(Classification) ? DashText : DarkText;
        public Brush GraduationYearForeground => string.IsNullOrWhiteSpace(GraduationYear) ? DashText : DarkText;

        public Brush StarButtonBg => IsPrimary ? StarActiveBg : StarInactiveBg;
        public Brush StarButtonBorder => IsPrimary ? StarActiveBorder : StarInactiveBorder;
        public Brush StarButtonFg => IsPrimary ? StarActiveFg : StarInactiveFg;
        public PackIconKind StarButtonIcon => IsPrimary ? PackIconKind.Star : PackIconKind.StarOutline;
        public string StarButtonTooltip => IsPrimary ? "⭐ Bằng chính (đang hiển thị đại diện trên hồ sơ)" : "Nhấp để đặt làm bằng chính (hiển thị đại diện)";

        public DegreeViewModel(PersonnelDegree d, int stt)
        {
            STT = stt;
            Id = d.Id;
            DegreeType = d.DegreeType;
            DegreeName = d.DegreeName;
            Major = d.Major;
            Institution = d.Institution;
            Classification = d.Classification;
            GraduationYear = d.GraduationYear;
            IsPrimary = d.IsPrimary;
        }
    }
}
