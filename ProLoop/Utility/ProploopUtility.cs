using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.Utility
{
    class ProploopUtility
    {
        public static void ActivateOpenPanel()
        {
            var currentInstance = AddinModule.CurrentInstance;
            currentInstance.ActivatePanel();
        }
    }
}
