using System;
using System.Diagnostics;
using System.IO;
using Entel.ECC.UtilRegistro;
using SERVICIO_ATT_VALIDACION_CUENTAS.App_Code;

namespace ServiceMail
{
    class Config
    {

        private string Seccion = "SERVICIO_ATT_VALIDACION_CUENTAS";
        public string PkgLogs { get; set; } = "LOGIN_CENTRALIZADO.PKG_LC_MS_LOGS";
        public string Udl { get; set; }
        public string ConUdl { get; set; }
        public string Pkg { get; set; }
        public string LogServicio { get; set; }
        public Int32 Intervalo { get; set; }
        public Int32 IntervaloEspera { get; set; }
        public int NivelLog { get; set; }
        public string CorreoServidor { get; set; }
        public string PuertoSMTP { get; set; }
        public string UsuarioSMTP { get; set; }
        public string PasswordSMTP { get; set; }
        public string CorreoPara { get; set; }
        public string CorreoAsunto { get; set; }
        public string CorreoMensaje { get; set; }

        public Config()
        {

            Registro registro = new Registro();

            registro.AbreSeccion(@"Software\Microsoft\Windows\" + this.Seccion);

            //Datos del Schema LOGIN_CENTRALIZADO
            this.Udl = registro.LeeValor("udl");
            this.ConUdl = this.ObtieneConexionOracle(this.Udl);
            this.Pkg = registro.LeeValor("pkg");

            this.LogServicio = registro.LeeValor("Directorio_Archivo_LOG");
            if (!this.LogServicio.EndsWith("\\")) this.LogServicio += "\\";

            string intervalo = registro.LeeValor("Timer_de_Ejecucion");

            if (!String.IsNullOrEmpty(intervalo))
                this.Intervalo = Convert.ToInt32(intervalo) * 1000;

            //string IntervaloEspera = registro.LeeValor("Timer_de_Verificacion");

            //if (!String.IsNullOrEmpty(IntervaloEspera))
            //    this.IntervaloEspera = Convert.ToInt32(IntervaloEspera) * 1000;

            string nivel = "0";

            if (string.IsNullOrEmpty(nivel)) nivel = "0";

            this.NivelLog = Convert.ToInt32(nivel);

            // SERVIDOR SMTP
            this.CorreoServidor = registro.LeeValor("correo_servidor");
            this.PuertoSMTP = registro.LeeValor("puerto_smtp");
            this.UsuarioSMTP = registro.LeeValor("correo_de");
            this.PasswordSMTP = registro.LeeValor("correo_password");
            this.CorreoPara = registro.LeeValor("correo_para");
            this.CorreoAsunto = registro.LeeValor("correo_asunto");
            this.CorreoMensaje = registro.LeeValor("correo_mensaje");

        }

        /// <summary>
        /// Valida que la informacion contenida en el registro sea valida
        /// </summary>
        /// <returns>Verdadero si los valores leidos fueron correctos</returns>
        public bool ValidaInfo()
        {

            // Conexion a LOGIN_CENTRALIZADO
            if (string.IsNullOrEmpty(this.Udl))
            {
                EventLog.WriteEntry(this.Seccion, "Debe definir archivo de conexion UDL a Oracle del Schema LOGIN_CENTRALIZADO.", EventLogEntryType.Error);
                return false;
            }

            if (!File.Exists(this.Udl))
            {
                EventLog.WriteEntry(this.Seccion, "Archivo de definicion de conexion a DB a Oracle del Schema  no existe en ruta especificada.", EventLogEntryType.Error);
                return false;
            }

            if (string.IsNullOrEmpty(this.Pkg))
            {
                EventLog.WriteEntry(this.Seccion, "Debe definir Package del Schema LOGIN_CENTRALIZADO", EventLogEntryType.Error);
                return false;
            }

            if (string.IsNullOrEmpty(this.LogServicio))
            {
                EventLog.WriteEntry(this.Seccion, "Debe definir la ruta de log del servicio.", EventLogEntryType.Error);
                return false;
            }

            return true;

        }

        /// <summary>
        /// Valida que la informacion contenida en el registro sea valida
        /// </summary>
        /// <returns>Verdadero si los valores leidos fueron correctos</returns>
        public bool ValidaInfoCorreo()
        {
            // EL Servidor SMTP
            if (string.IsNullOrEmpty(this.CorreoServidor))
            {
                EventLog.WriteEntry(this.Seccion, "Debe el servidor SMTP", EventLogEntryType.Error);
                return false;
            }


            //Validando el correo de
            if (string.IsNullOrEmpty(this.UsuarioSMTP))
            {
                EventLog.WriteEntry(this.Seccion, "Debe indicar el correo", EventLogEntryType.Error);
                return false;
            }
            else
            {
                NavigatorLibreria Libreria = new NavigatorLibreria();
                if (this.UsuarioSMTP.Contains(";"))
                {
                    EventLog.WriteEntry(this.Seccion, "Solo se admite un correo en (Correo de ) ", EventLogEntryType.Error);
                    return false;
                }
                else
                {
                    if (Libreria.ValidarEmail(this.UsuarioSMTP) == false)
                    {
                        EventLog.WriteEntry(this.Seccion, "El Correo " + this.UsuarioSMTP + ", Es invalido", EventLogEntryType.Error);
                        return false;
                    }
                }
            }

            //Validando el correo para
            if (string.IsNullOrEmpty(this.CorreoPara))
            {
                EventLog.WriteEntry(this.Seccion, "Debe indicar el(los) Correo(s) del Para", EventLogEntryType.Error);
                return false;
            }
            else
            {
                NavigatorLibreria Libreria = new NavigatorLibreria();
                if (this.CorreoPara.Contains(";"))
                {
                    string[] Correos = this.CorreoPara.Split(';');
                    foreach (string Correo in Correos)
                    {
                        if (Libreria.ValidarEmail(Correo) == false)
                        {
                            EventLog.WriteEntry(this.Seccion, "El Correo para " + Correo + ", Es invalido", EventLogEntryType.Error);
                            return false;
                        }
                    }
                }
                else
                {
                    if (Libreria.ValidarEmail(this.CorreoPara) == false)
                    {
                        EventLog.WriteEntry(this.Seccion, "El Correo para " + this.CorreoPara + ", Es invalido", EventLogEntryType.Error);
                        return false;
                    }
                }
            }

            if (string.IsNullOrEmpty(this.CorreoAsunto))
            {
                EventLog.WriteEntry(this.Seccion, "Debe indicar el asunto.", EventLogEntryType.Error);
                return false;
            }

            if (string.IsNullOrEmpty(this.CorreoMensaje))
            {
                EventLog.WriteEntry(this.Seccion, "Debe indicar el cuerpo del mensaje.", EventLogEntryType.Error);
                return false;
            }

            return true;

        }

        /// <summary>
        /// Imprime los valores de configuracion obtenidos del registro
        /// </summary>
        public void ImprimeConfiguracion()
        {
            EventLog.WriteEntry(this.Seccion, "Conexión UDL a Oracle de Schema LOGIN_CENTRALIZADO en Servicio SERVICIO_ATT_VALIDACION_CUENTAS : " + this.Udl, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Package de Schema LOGIN_CENTRALIZADO en SERVICIO_ATT_VALIDACION_CUENTAS : " + this.Pkg, EventLogEntryType.Information);

            EventLog.WriteEntry(this.Seccion, "Ruta de log : " + this.LogServicio, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Intervalo de tiempo de espera del servicio : " + this.IntervaloEspera, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Intervalo de tiempo de ejecucion del servicio : " + this.Intervalo, EventLogEntryType.Information);

            EventLog.WriteEntry(this.Seccion, "Servicio Correo : " + this.CorreoServidor, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Correo de       : " + this.UsuarioSMTP, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Correo para     : " + this.CorreoPara, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Correo asunto   : " + this.CorreoAsunto, EventLogEntryType.Information);
            EventLog.WriteEntry(this.Seccion, "Correo cuerpo   : " + this.CorreoMensaje, EventLogEntryType.Information);
        }

        /// <summary>
        /// Obtiene la informacion de la conexion desde archivo UDL
        /// </summary>
        /// <param name="archivo">El nombre del archivo UDL</param>
        /// <returns>La informacion de la conexion obtenida del archivo UDL</returns>
        private string ObtieneConexionOracle(string archivo)
        {

            string ret = String.Empty;

            if (String.IsNullOrEmpty(archivo)) return ret;
            if (!File.Exists(archivo)) return ret;

            try
            {
                StreamReader swr = new StreamReader(archivo);

                while (true)
                {
                    string linea = swr.ReadLine();

                    if (linea == null) break;

                    if (linea.StartsWith("Provider="))
                    {
                        string buffer = linea.Substring(9);
                        string[] data = buffer.Split(';');
                        ret = data[3] + ";" + data[2] + ";" + data[1];
                        break;
                    }
                }

                swr.Close();
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(this.Seccion, "Error al abrir archivo de conexión a base : " + archivo + " error : " + ex.Message, EventLogEntryType.Error);
            }

            return ret;
        }

        private string ObtieneConexionSQLServer(string archivo)
        {

            string ret = String.Empty;

            if (String.IsNullOrEmpty(archivo)) return ret;
            if (!File.Exists(archivo)) return ret;

            try
            {
                StreamReader swr = new StreamReader(archivo);

                while (true)
                {
                    string linea = swr.ReadLine();

                    if (linea == null) break;

                    if (linea.StartsWith("Provider="))
                    {
                        string[] data = linea.Split(';');
                        ret = data[1] + ";" + data[2] + ";" + data[3] + ";" + data[4] + ";" + data[5];
                        break;
                    }
                }

                swr.Close();
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(this.Seccion, "Error al abrir archivo de conexión a base : " + archivo + " error : " + ex.Message, EventLogEntryType.Error);
            }

            return ret;
        }

    }
}
