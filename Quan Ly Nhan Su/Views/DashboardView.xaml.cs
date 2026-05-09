using System.Windows.Controls;

namespace TaxPersonnelManagement.Views
{
    public partial class DashboardView : UserControl
    {
        public DashboardView(int? targetPersonnelId = null)
        {
            InitializeComponent();
            if (targetPersonnelId.HasValue)
            {
                PersonnelList.TargetPersonnelId = targetPersonnelId;
                PersonnelList.LoadData();
            }
        }
    }
}
