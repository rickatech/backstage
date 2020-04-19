using Autofac;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProLoop.Services;

namespace ProLoop.Startup
{
   public class Bootstrapper
    {
        public IContainer Bootstrap()
        {
            var builder = new ContainerBuilder();
            builder.RegisterType<Controller.MainEntry>().AsSelf();
            builder.RegisterType<ViewModels.LoginViewModel>().AsSelf();
            builder.RegisterType<ViewModels.OpenViewModel>().AsSelf();
            builder.RegisterType<ProloopClient>().As<IProloopClient>();
            return builder.Build();
        }
    }
}
