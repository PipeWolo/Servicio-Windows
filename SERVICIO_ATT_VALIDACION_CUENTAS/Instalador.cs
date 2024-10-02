using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;

namespace ServiceMail
{
    [RunInstallerAttribute(true)]
    public class ProjectInstaller : Installer
    {
        private ServiceInstaller serviceInstaller;
        private ServiceProcessInstaller processInstaller;

        /// <summary>
        /// El constructor instala el servicio en la lista de servicios de windows
        /// y ademas reporta los eventos en el event viewer
        /// </summary>
        public ProjectInstaller()
        {

            processInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();

            processInstaller.Account = ServiceAccount.LocalSystem;

            serviceInstaller.StartType = ServiceStartMode.Manual;

            serviceInstaller.ServiceName = "SERVICIO_ATT_VALIDACION_CUENTAS";
            serviceInstaller.DisplayName = "SERVICIO_ATT_VALIDACION_CUENTAS";
            serviceInstaller.Description = "Servicio que valida cuentas de correo personales";

            Installers.Add(serviceInstaller);
            Installers.Add(processInstaller);

        }
    }
}
