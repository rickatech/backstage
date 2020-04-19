using System.Windows.Input;
using ProLoop.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Prism.Commands;
using System.Windows.Controls;

namespace ProLoop.ViewModels
{
    public class LoginViewModel : ViewModelBase
    {
        private IProloopClient _proloopClient;
        public LoginViewModel(IProloopClient proloopClient)
        {
            _proloopClient = proloopClient;
            LoginCommand = new DelegateCommand<object>(ProcessLogin);
            ClearCommand = new DelegateCommand(ClearLogin);
        }

        private async void ProcessLogin(object obj)
        {
            PasswordBox passwordBox = obj as PasswordBox;
            UserPassword = passwordBox.Password;
            await _proloopClient.LoginAsync(UserName, UserPassword, Url);
            Utility.ProploopUtility.ActivateOpenPanel();
        }

        private void ClearLogin()
        {
            UserName = string.Empty;
            Url = string.Empty;
        }

        public void GetClient()
        {

        }
        public void GetProject()
        {

        }
        public void GetMatter()
        {

        }
        public void GetFile()
        {

        }

        private string userPassword;

        public string UserPassword
        {
            get { return userPassword; }
            set { userPassword = value;
                OnPropertyChanganged();
            }
        }

        private string username;
        public string UserName
        {
            get { return username; }
            set { username = value;
                OnPropertyChanganged();
            }
        }
        private string  url="";      

        public string Url
        {
            get { return url; }
            set { url = value;
                OnPropertyChanganged();
            }
        }

        public DelegateCommand<object> LoginCommand { get; private set; }
        public DelegateCommand ClearCommand { get; private set; }
    }
}
