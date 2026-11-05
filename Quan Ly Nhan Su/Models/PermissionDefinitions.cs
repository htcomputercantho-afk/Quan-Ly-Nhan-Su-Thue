using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace TaxPersonnelManagement.Models
{
    /// <summary>
    /// Đại diện cho 1 quyền hạn cụ thể trong hệ thống.
    /// Có hỗ trợ thông báo thay đổi (INotifyPropertyChanged) phục vụ data binding trong WPF.
    /// </summary>
    public class PermissionItem : INotifyPropertyChanged
    {
        private string _code = string.Empty;
        private string _name = string.Empty;
        private string _description = string.Empty;
        private bool _isSelected = false;

        public string Code
        {
            get => _code;
            set { _code = value; OnPropertyChanged(nameof(Code)); }
        }

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public string Description
        {
            get => _description;
            set { _description = value; OnPropertyChanged(nameof(Description)); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                    SelectionChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public event EventHandler? SelectionChanged;
        public event PropertyChangedEventHandler? PropertyChanged;

        public PermissionItem() { }

        public PermissionItem(string code, string name, string description = "")
        {
            Code = code;
            Name = name;
            Description = description;
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public PermissionItem Clone()
        {
            return new PermissionItem
            {
                Code = this.Code,
                Name = this.Name,
                Description = this.Description,
                IsSelected = this.IsSelected
            };
        }
    }

    /// <summary>
    /// Nhóm các quyền hạn theo Module / Chức năng hoặc Danh mục Select.
    /// Có checkbox chọn nhanh toàn bộ quyền trong nhóm (Tri-state).
    /// </summary>
    public class PermissionGroup : INotifyPropertyChanged
    {
        private string _groupId = string.Empty;
        private string _groupName = string.Empty;
        private string _groupIcon = "FolderOutline";
        private string _description = string.Empty;
        private ObservableCollection<PermissionItem> _permissions = new();
        private bool _isUpdatingGroupSelection = false;

        public string GroupId
        {
            get => _groupId;
            set { _groupId = value; OnPropertyChanged(nameof(GroupId)); }
        }

        public string GroupName
        {
            get => _groupName;
            set { _groupName = value; OnPropertyChanged(nameof(GroupName)); }
        }

        public string GroupIcon
        {
            get => _groupIcon;
            set { _groupIcon = value; OnPropertyChanged(nameof(GroupIcon)); }
        }

        public string Description
        {
            get => _description;
            set { _description = value; OnPropertyChanged(nameof(Description)); }
        }

        public ObservableCollection<PermissionItem> Permissions
        {
            get => _permissions;
            set
            {
                _permissions = value;
                HookItemEvents();
                OnPropertyChanged(nameof(Permissions));
                OnPropertyChanged(nameof(IsAllSelected));
            }
        }

        /// <summary>
        /// Trạng thái checkbox nhóm: true = chọn hết, false = bỏ hết, null = chọn một phần
        /// </summary>
        public bool? IsAllSelected
        {
            get
            {
                if (Permissions.Count == 0) return false;
                int count = Permissions.Count(p => p.IsSelected);
                if (count == Permissions.Count) return true;
                if (count == 0) return false;
                return null;
            }
            set
            {
                if (_isUpdatingGroupSelection || !value.HasValue) return;

                _isUpdatingGroupSelection = true;
                bool target = value.Value;
                foreach (var item in Permissions)
                {
                    item.IsSelected = target;
                }
                _isUpdatingGroupSelection = false;

                OnPropertyChanged(nameof(IsAllSelected));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void HookItemEvents()
        {
            foreach (var item in _permissions)
            {
                item.SelectionChanged -= Item_SelectionChanged;
                item.SelectionChanged += Item_SelectionChanged;
            }
        }

        private void Item_SelectionChanged(object? sender, EventArgs e)
        {
            if (!_isUpdatingGroupSelection)
            {
                OnPropertyChanged(nameof(IsAllSelected));
            }
        }

        public PermissionGroup Clone()
        {
            var grp = new PermissionGroup
            {
                GroupId = this.GroupId,
                GroupName = this.GroupName,
                GroupIcon = this.GroupIcon,
                Description = this.Description
            };
            foreach (var p in this.Permissions)
            {
                grp.Permissions.Add(p.Clone());
            }
            grp.HookItemEvents();
            return grp;
        }
    }

    /// <summary>
    /// Bộ đăng ký tập trung tất cả các quyền hạn trong toàn bộ ứng dụng (Dynamic Permission Registry).
    /// Cho phép tự động sinh danh sách ô tích phân quyền và tự thích ứng khi lập trình viên thêm chức năng mới.
    /// </summary>
    public static class PermissionRegistry
    {
        // 1. Quản lý Danh mục & Popup Select (Thêm data nhỏ cho combobox)
        public const string Catalog_Department = "Catalog_Department";
        public const string Catalog_Position = "Catalog_Position";
        public const string Catalog_Rank = "Catalog_Rank";
        public const string Catalog_SalaryConfig = "Catalog_SalaryConfig";
        public const string Catalog_SalaryDelay = "Catalog_SalaryDelay";
        public const string Catalog_Discipline = "Catalog_Discipline";
        public const string Catalog_PlanningTerm = "Catalog_PlanningTerm";
        public const string Catalog_TrainingClass = "Catalog_TrainingClass";

        // 2. Tổng quan & Thống kê
        public const string Menu_Dashboard = "Menu_Dashboard";
        public const string Menu_Statistics = "Menu_Statistics";
        public const string Statistics_Export = "Statistics_Export";

        // 3. Hồ sơ Cán bộ (Nhân sự)
        public const string Menu_Personnel = "Menu_Personnel";
        public const string Personnel_AddEdit = "Personnel_AddEdit";
        public const string Personnel_Delete = "Personnel_Delete";
        public const string Personnel_ImportExport = "Personnel_ImportExport";

        // 4. Lương & Thu nhập
        public const string Menu_Salary = "Menu_Salary";
        public const string Salary_Update = "Salary_Update";
        public const string Salary_Export = "Salary_Export";
        public const string Menu_AnnualIncome = "Menu_AnnualIncome";
        public const string Income_Edit = "Income_Edit";
        public const string Income_Export = "Income_Export";

        // 5. Chi tiết nghỉ & Thời gian công tác
        public const string Menu_LeaveDetail = "Menu_LeaveDetail";
        public const string Leave_Manage = "Leave_Manage";
        public const string Menu_PositionDuration = "Menu_PositionDuration";

        // 6. Thi đua khen thưởng & Đánh giá xếp loại
        public const string Menu_EmulationReward = "Menu_EmulationReward";
        public const string EmulationReward_Manage = "EmulationReward_Manage";
        public const string Menu_Evaluation = "Menu_Evaluation";
        public const string Evaluation_Manage = "Evaluation_Manage";
        public const string Evaluation_Export = "Evaluation_Export";

        // 7. Đào tạo, bồi dưỡng
        public const string Menu_Training = "Menu_Training";
        public const string Training_Manage = "Training_Manage";
        public const string Training_DeleteClass = "Training_DeleteClass";
        public const string Training_Export = "Training_Export";

        // 8. Quy hoạch, Luân chuyển & Phân công nhiệm vụ
        public const string Menu_Planning = "Menu_Planning";
        public const string Planning_Manage = "Planning_Manage";
        public const string Planning_Export = "Planning_Export";
        public const string Menu_Rotation = "Menu_Rotation";
        public const string Rotation_Manage = "Rotation_Manage";
        public const string Rotation_Export = "Rotation_Export";
        public const string Menu_DutyAssignment = "Menu_DutyAssignment";
        public const string DutyAssignment_Upload = "DutyAssignment_Upload";

        // 9. Quản trị hệ thống
        public const string RoleManagement_Manage = "Menu_RoleManagement";
        public const string UserManagement_Manage = "Menu_UserManagement";
        public const string BackupRestore_Manage = "Menu_BackupRestore";

        private static readonly List<PermissionGroup> _registeredGroups = new();

        static PermissionRegistry()
        {
            InitializeRegistry();
        }

        private static void InitializeRegistry()
        {
            _registeredGroups.Clear();

            // Nhóm 1: Quản trị Popup danh mục select
            var groupCatalog = new PermissionGroup
            {
                GroupId = "CATALOG",
                GroupName = "1. Quản lý Danh mục & Popup Thêm Select (Nút '+', '⚙️')",
                GroupIcon = "FolderCogOutline",
                Description = "Kiểm soát các popup thêm data nhỏ cho các ô chọn select/combobox"
            };
            groupCatalog.Permissions.Add(new(Catalog_Department, "Quản lý Bộ phận / Phòng ban (Popup)", "Cho phép thêm, sửa, xóa phòng ban trong popup chọn bộ phận"));
            groupCatalog.Permissions.Add(new(Catalog_Position, "Quản lý Chức vụ (Popup)", "Cho phép thêm, sửa, xóa chức vụ chính quyền và đoàn thể"));
            groupCatalog.Permissions.Add(new(Catalog_Rank, "Quản lý Ngạch công chức (Popup)", "Cho phép cấu hình danh mục ngạch, tên ngạch và niên hạn"));
            groupCatalog.Permissions.Add(new(Catalog_SalaryConfig, "Cấu hình Bảng lương & Lương cơ sở (Popup)", "Cho phép chỉnh sửa bảng hệ số ngạch/bậc và mức lương cơ sở"));
            groupCatalog.Permissions.Add(new(Catalog_SalaryDelay, "Cấu hình Kéo dài nâng bậc lương (Popup)", "Cấu hình danh mục lý do kéo dài thời hạn nâng lương"));
            groupCatalog.Permissions.Add(new(Catalog_Discipline, "Cấu hình Hình thức kỷ luật (Popup)", "Cấu hình danh mục hình thức kỷ luật công chức"));
            groupCatalog.Permissions.Add(new(Catalog_PlanningTerm, "Quản lý Nhiệm kỳ & Chức danh quy hoạch (Popup)", "Quản lý danh mục nhiệm kỳ và chức danh trong quy hoạch"));
            groupCatalog.Permissions.Add(new(Catalog_TrainingClass, "Thêm mới Lớp đào tạo & Hội nghị (Popup)", "Mở popup tạo và thiết lập thông tin lớp bồi dưỡng"));
            groupCatalog.HookItemEvents();
            _registeredGroups.Add(groupCatalog);

            // Nhóm 2: Hồ sơ Cán bộ (Nhân sự)
            var groupPersonnel = new PermissionGroup
            {
                GroupId = "PERSONNEL",
                GroupName = "2. Phân hệ Hồ sơ Cán bộ (Nhân sự)",
                GroupIcon = "AccountGroupOutline",
                Description = "Xem, thêm mới, sửa, xóa và xuất nhập Excel hồ sơ cán bộ"
            };
            groupPersonnel.Permissions.Add(new(Menu_Personnel, "Xem Menu Nhân sự", "Hiển thị menu và cho phép xem chi tiết hồ sơ cán bộ"));
            groupPersonnel.Permissions.Add(new(Personnel_AddEdit, "Thêm mới & Cập nhật hồ sơ (Nút Lưu)", "Cho phép chỉnh sửa thông tin cá nhân, công tác, văn bằng, quá trình"));
            groupPersonnel.Permissions.Add(new(Personnel_Delete, "Xóa hồ sơ cán bộ", "Cho phép xóa hồ sơ cán bộ ra khỏi hệ thống"));
            groupPersonnel.Permissions.Add(new(Personnel_ImportExport, "Xuất & Nhập Excel cán bộ", "Cho phép tải file Excel mẫu và xuất báo cáo danh sách công chức"));
            groupPersonnel.HookItemEvents();
            _registeredGroups.Add(groupPersonnel);

            // Nhóm 3: Lương & Thu nhập
            var groupSalary = new PermissionGroup
            {
                GroupId = "SALARY",
                GroupName = "3. Phân hệ Lương & Tổng thu nhập",
                GroupIcon = "CashMultiple",
                Description = "Theo dõi quá trình nâng lương và tổng hợp thu nhập hàng năm"
            };
            groupSalary.Permissions.Add(new(Menu_Salary, "Xem Menu Nâng lương", "Xem danh sách đến hạn nâng lương thường xuyên và trước hạn"));
            groupSalary.Permissions.Add(new(Salary_Update, "Cập nhật nâng lương & Nâng trước hạn", "Cập nhật ngày hưởng lương mới, quyết định và xét trước hạn"));
            groupSalary.Permissions.Add(new(Salary_Export, "Xuất Excel danh sách nâng lương", "Xuất file báo cáo danh sách nâng lương"));
            groupSalary.Permissions.Add(new(Menu_AnnualIncome, "Xem Menu Tổng thu nhập năm", "Xem bảng chi tiết thu nhập từng tháng của công chức"));
            groupSalary.Permissions.Add(new(Income_Edit, "Chỉnh sửa & Nhập hàng loạt thu nhập", "Cho phép nhập dữ liệu lương, thưởng, làm thêm giờ và nhập Excel"));
            groupSalary.Permissions.Add(new(Income_Export, "Xuất Excel bảng kê thu nhập", "Xuất file Excel tổng hợp thu nhập"));
            groupSalary.HookItemEvents();
            _registeredGroups.Add(groupSalary);

            // Nhóm 4: Nghỉ phép & Thời gian công tác
            var groupLeave = new PermissionGroup
            {
                GroupId = "LEAVE_POSITION",
                GroupName = "4. Phân hệ Nghỉ phép & Thời gian công tác",
                GroupIcon = "CalendarClockOutline",
                Description = "Quản lý ngày nghỉ, phép năm và thời gian giữ vị trí công tác"
            };
            groupLeave.Permissions.Add(new(Menu_LeaveDetail, "Xem Menu Chi tiết nghỉ", "Xem danh sách tổng hợp ngày nghỉ phép, ốm, thai sản..."));
            groupLeave.Permissions.Add(new(Leave_Manage, "Thêm & Điều chỉnh ngày nghỉ", "Thêm mới đợt nghỉ hoặc cập nhật, xóa lịch sử nghỉ của cán bộ"));
            groupLeave.Permissions.Add(new(Menu_PositionDuration, "Xem Menu Thời gian công tác", "Theo dõi thời gian giữ vị trí lãnh đạo, ngạch, công tác theo quy định"));
            groupLeave.HookItemEvents();
            _registeredGroups.Add(groupLeave);

            // Nhóm 5: Thi đua khen thưởng & Đánh giá xếp loại
            var groupReward = new PermissionGroup
            {
                GroupId = "REWARD_EVALUATION",
                GroupName = "5. Phân hệ Thi đua Khen thưởng & Xếp loại",
                GroupIcon = "CertificateOutline",
                Description = "Quản lý danh hiệu thi đua, khen thưởng và đánh giá công chức hàng năm"
            };
            groupReward.Permissions.Add(new(Menu_EmulationReward, "Xem Menu Thi đua khen thưởng", "Xem danh sách quyết định khen thưởng"));
            groupReward.Permissions.Add(new(EmulationReward_Manage, "Thêm, sửa, xóa khen thưởng", "Cập nhật danh hiệu thi đua và quyết định khen thưởng"));
            groupReward.Permissions.Add(new(Menu_Evaluation, "Xem Menu Đánh giá xếp loại chất lượng", "Xem lịch sử xếp loại chất lượng công chức"));
            groupReward.Permissions.Add(new(Evaluation_Manage, "Cập nhật xếp loại & Nhập Excel", "Lưu kết quả xếp loại, sửa trực tiếp, quyết định hàng loạt và nhập Excel"));
            groupReward.Permissions.Add(new(Evaluation_Export, "Xuất Excel danh sách xếp loại", "Xuất file báo cáo danh sách xếp loại ra Excel"));
            groupReward.HookItemEvents();
            _registeredGroups.Add(groupReward);

            // Nhóm 6: Đào tạo, bồi dưỡng
            var groupTraining = new PermissionGroup
            {
                GroupId = "TRAINING",
                GroupName = "6. Phân hệ Đào tạo, Bồi dưỡng",
                GroupIcon = "SchoolOutline",
                Description = "Theo dõi lớp tập huấn, bồi dưỡng chuyên môn và phân bổ học viên"
            };
            groupTraining.Permissions.Add(new(Menu_Training, "Xem Menu Đào tạo bồi dưỡng", "Xem danh sách lớp đào tạo và học viên tham gia"));
            groupTraining.Permissions.Add(new(Training_Manage, "Quản lý lớp & Phân công học viên", "Thêm mới, sửa thông tin lớp và phân công cán bộ đào tạo"));
            groupTraining.Permissions.Add(new(Training_DeleteClass, "Xóa lớp đào tạo", "Xóa lớp đào tạo ra khỏi hệ thống"));
            groupTraining.Permissions.Add(new(Training_Export, "Xuất Excel danh sách đào tạo", "Xuất file danh sách lớp và học viên ra Excel"));
            groupTraining.HookItemEvents();
            _registeredGroups.Add(groupTraining);

            // Nhóm 7: Quy hoạch, Luân chuyển & Phân công nhiệm vụ
            var groupPlanning = new PermissionGroup
            {
                GroupId = "PLANNING_ROTATION",
                GroupName = "7. Quy hoạch, Luân chuyển & Phân công nhiệm vụ",
                GroupIcon = "AccountSwitchOutline",
                Description = "Quy hoạch cán bộ, luân chuyển điều động và tài liệu phân công nhiệm vụ"
            };
            groupPlanning.Permissions.Add(new(Menu_Planning, "Xem Menu Quy hoạch", "Xem danh sách cán bộ quy hoạch các nhiệm kỳ"));
            groupPlanning.Permissions.Add(new(Planning_Manage, "Thêm, sửa, xóa quy hoạch", "Quản lý dữ liệu quy hoạch cán bộ"));
            groupPlanning.Permissions.Add(new(Planning_Export, "Xuất Excel dữ liệu quy hoạch", "Xuất file báo cáo quy hoạch cán bộ ra Excel"));
            groupPlanning.Permissions.Add(new(Menu_Rotation, "Xem Menu Luân chuyển, Điều động", "Xem danh sách luân chuyển theo kế hoạch và ngoài kế hoạch"));
            groupPlanning.Permissions.Add(new(Rotation_Manage, "Thêm, sửa, xóa luân chuyển điều động", "Quản lý hồ sơ luân chuyển cán bộ"));
            groupPlanning.Permissions.Add(new(Rotation_Export, "Xuất Excel danh sách luân chuyển", "Xuất file Excel danh sách cán bộ luân chuyển"));
            groupPlanning.Permissions.Add(new(Menu_DutyAssignment, "Xem Menu Phân công nhiệm vụ", "Xem tài liệu phân công nhiệm vụ theo Tổ hàng tháng"));
            groupPlanning.Permissions.Add(new(DutyAssignment_Upload, "Tải lên & Xóa file phân công", "Tải file PDF phân công nhiệm vụ lên hệ thống"));
            groupPlanning.HookItemEvents();
            _registeredGroups.Add(groupPlanning);

            // Nhóm 8: Tổng quan & Báo cáo Thống kê
            var groupDashboard = new PermissionGroup
            {
                GroupId = "DASHBOARD_STATS",
                GroupName = "8. Tổng quan & Báo cáo Thống kê",
                GroupIcon = "ChartPie",
                Description = "Xem bảng số liệu tổng quan và báo cáo thống kê nhân sự"
            };
            groupDashboard.Permissions.Add(new(Menu_Dashboard, "Xem Menu Tổng quan (Dashboard)", "Xem các chỉ số tổng quan nhanh"));
            groupDashboard.Permissions.Add(new(Menu_Statistics, "Xem Menu Thống kê", "Xem biểu đồ cơ cấu, trình độ, độ tuổi và báo cáo chi tiết"));
            groupDashboard.Permissions.Add(new(Statistics_Export, "Xuất Excel báo cáo thống kê", "Xuất file Excel báo cáo thống kê"));
            groupDashboard.HookItemEvents();
            _registeredGroups.Add(groupDashboard);

            // Nhóm 9: Quản trị Hệ thống
            var groupSystem = new PermissionGroup
            {
                GroupId = "SYSTEM",
                GroupName = "9. Quản trị Hệ thống (Dành cho Admin)",
                GroupIcon = "ShieldAccountOutline",
                Description = "Quản lý vai trò phân quyền, tài khoản và sao lưu dữ liệu"
            };
            groupSystem.Permissions.Add(new(RoleManagement_Manage, "Quản lý Vai trò & Phân quyền (Menu Vai trò)", "Xem, tạo mới, chỉnh sửa và cấu hình quyền hạn vai trò"));
            groupSystem.Permissions.Add(new(UserManagement_Manage, "Quản lý Tài khoản (Menu Tài khoản)", "Tạo tài khoản nhân viên, gán vai trò và cấp lại mật khẩu"));
            groupSystem.Permissions.Add(new(BackupRestore_Manage, "Sao lưu & Phục hồi CSDL (Menu Sao lưu)", "Sao lưu dữ liệu, đồng bộ Google Drive và phục hồi cơ sở dữ liệu"));
            groupSystem.HookItemEvents();
            _registeredGroups.Add(groupSystem);
        }

        /// <summary>
        /// Tạo bản sao các nhóm quyền để hiển thị trên giao diện tạo/sửa vai trò.
        /// </summary>
        public static List<PermissionGroup> GetFreshGroups()
        {
            return _registeredGroups.Select(g => g.Clone()).ToList();
        }

        /// <summary>
        /// Cho phép các module/plugin tương lai đăng ký thêm chức năng mới một cách tự động.
        /// Chức năng mới này sẽ ngay lập tức xuất hiện trong hộp thoại phân quyền!
        /// </summary>
        public static void RegisterCustomFeature(string groupId, string groupName, string groupIcon, PermissionItem item)
        {
            var grp = _registeredGroups.FirstOrDefault(g => g.GroupId.Equals(groupId, StringComparison.OrdinalIgnoreCase));
            if (grp == null)
            {
                grp = new PermissionGroup
                {
                    GroupId = groupId,
                    GroupName = groupName,
                    GroupIcon = groupIcon
                };
                _registeredGroups.Add(grp);
            }

            if (!grp.Permissions.Any(p => p.Code.Equals(item.Code, StringComparison.OrdinalIgnoreCase)))
            {
                grp.Permissions.Add(item);
                grp.HookItemEvents();
            }
        }
    }
}

