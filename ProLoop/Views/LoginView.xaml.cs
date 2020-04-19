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
using ProLoop.ViewModels;

namespace ProLoop.Views
{
    /// <summary>
    /// Interaction logic for LoginView.xaml
    /// </summary>
    public partial class LoginView : UserControl
    {
        private LoginViewModel _loginViewModel;

        public LoginView(LoginViewModel loginViewModel)
        {
            InitializeComponent();
            this.DataContext = loginViewModel;
            _loginViewModel = loginViewModel;
            this.Loaded += LoginView_Loaded;
        }

        private void LoginView_Loaded(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_loginViewModel.UserPassword))
                controPassword.Password = _loginViewModel.UserPassword;
        }
    }
}
