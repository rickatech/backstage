using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProLoop.Services.Entity;

namespace ProLoop.ViewModels
{
    public class OpenViewModel : ViewModelBase
    {
        public ObservableCollection<ProjectOrg> ProjectList;
        public ObservableCollection<Client> Clients;
        public ObservableCollection<Matter> Matters;
        public OpenViewModel()
        {
            ProjectList = new ObservableCollection<ProjectOrg>() { new ProjectOrg() {  } };
            Clients = new ObservableCollection<Client>();
            Matters = new ObservableCollection<Matter>();
        }
        private string content;

        public string Content
        {
            get { return content; }
            set { content = value;
                OnPropertyChanganged();
            }
        }

        private string fileName;

        public string FileName
        {
            get { return fileName; }
            set { fileName = value;
                OnPropertyChanganged();
            }
        }
        private string keywords;

        public string Keywords
        {
            get { return keywords; }
            set { keywords = value;
                OnPropertyChanganged();
            }
        }

        private string editors;

        public string Editors
        {
            get { return editors; }
            set { editors = value;
                OnPropertyChanganged();
            }
        }

        private string docId;

        public string DocId
        {
            get { return docId; }
            set { docId = value;
                OnPropertyChanganged();
            }
        }

        private bool showFileTree;

        public bool ShowFileTree
        {
            get { return showFileTree; }
            set { showFileTree = value;
                OnPropertyChanganged();
            }
        }

        private bool showPath;

        public bool ShowPath
        {
            get { return showPath; }
            set { showPath = value;
                OnPropertyChanganged();
            }
        }

        private bool isProject;

        public bool IsProject
        {
            get { return isProject; ; }
            set { isProject = value;
                OnPropertyChanganged();
            }
        }


    }
}
