using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.WordAddin
{
    public class Project
    {
        public int Id
        {
            get;
            set;
        }

        public string Title
        {
            get;
            set;
        }
    }
}


		

	





public class ProLoopFile
{
    public string path
    {
        get;
        set;
    }

    public string name
    {
        get;
        set;
    }
}

public enum Context
{
    Orgs,
    Projects
}