using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Diagnostics;

namespace ServiceMail
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {
            #region PRODUCCION
            EventLog.WriteEntry("SERVICIO_ATT_VALIDACION_CUENTAS", DateTime.Now.ToLongTimeString() + " - Iniciando servicio...");
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new SERVICIO_ATT_VALIDACION_CUENTAS()
            };
            ServiceBase.Run(ServicesToRun);
            EventLog.WriteEntry("SERVICIO_ATT_VALIDACION_CUENTAS", DateTime.Now.ToLongTimeString() + " - Término servicio...");
            #endregion PRODUCCION

            #region DESARROLLO
            //SERVICIO_ATT_VALIDACION_CUENTAS servicio = new SERVICIO_ATT_VALIDACION_CUENTAS();
            //servicio.Iniciar();
            #endregion DESARROLLO
        }
    }
}
