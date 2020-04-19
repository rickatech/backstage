using ProLoop.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ProLoop.Views
{
    /// <summary>
    /// Interaction logic for OpenView.xaml
    /// </summary>
    public partial class OpenView : UserControl
    {
        private OpenViewModel _opeviewmodel;

        public OpenView(OpenViewModel opeviewmodel)
        {
            _opeviewmodel = opeviewmodel;
            InitializeComponent();            
        }

        private void RbtOrganization_Checked(object sender, RoutedEventArgs e)
        {
            _opeviewmodel.IsProject = !rbtOrganization.IsChecked.Value;
        }
    }
}
