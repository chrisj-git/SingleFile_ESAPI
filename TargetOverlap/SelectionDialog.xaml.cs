using System.Windows;
using System.Windows.Controls;

namespace TargetOverlap
{
    /// <summary>
    /// Interaction logic for SelectionDialog.xaml
    /// </summary>
    public partial class SelectionDialog : Window
    {
        public string SelectedOption { get; private set; }
        public SelectionDialog()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Structures_comboBox.SelectedItem is string selectedItem)
            {
                SelectedOption = selectedItem;
                DialogResult = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("Please select an option.");
            }
        }
    }
}
