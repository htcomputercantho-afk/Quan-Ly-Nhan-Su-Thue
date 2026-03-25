using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class AddDisciplineTypeDialog : Window
    {
        public string SelectedDisciplineType { get; private set; } = string.Empty;

        public AddDisciplineTypeDialog()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên hình thức kỷ luật!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            using (var context = new AppDbContext())
            {
                var newType = new DisciplineType { Name = txtName.Text.Trim() };
                context.DisciplineTypes.Add(newType);
                context.SaveChanges();
                SelectedDisciplineType = newType.Name;
            }

            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
