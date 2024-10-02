using Limilabs.Mail;
using SERVICIO_ATT_VALIDACION_CUENTAS;
using SERVICIO_ATT_VALIDACION_CUENTAS.App_Code;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ServiceMail
{
    public partial class SERVICIO_ATT_VALIDACION_CUENTAS : ServiceBase
    {

        // CONFIGURACION INDIVIDUAL
        private string seccion_evt = "SERVICIO_ATT_VALIDACION_CUENTAS";
        private string Sistema = "Servicio Validacion Cuentas";
        private Thread workerThread = null;
        private bool Ejecutar = false;
        //private string timer = "60000";
        //private string timerEspera = "120000"; // dos minutos
        private string modoActual = "EJECUTAR"; // Indica si el servicio se ejecuta o no
        private string IdAplicacion = "17"; // Este es el ID de la aplicacion en la Tabla TB_APPSERVICES
        private string DataSource = "13";   // 
        //private Configuracion configuracionMail;
        private List<ConfiguracionAplicacion> Configuracion;        // TIENE LAS CONFIGURACIONES DE MI APLICACION
        //private List<ConfiguracionAplicacion> ConfiguracionGeneral; // TIENE LAS CONFIGURACIONES DE GENERALES DE TODAS LAS APLICACIONES
        ConfiguracionAplicacion ConfiguracionEjecucion;             // TIENE EL PRIMER ELEMENTO DE LA LITA Configuracion con la finalidad de evaluar registros

        // CONFIGURACION GENERAL
        private NavigatorLibreria Libreria = new NavigatorLibreria();
        private Log log;
        private Config config;
        private Config config2;
        private BD bd;
        private BD bdLogs;
        private Tb_Monitor Monitor;


        public SERVICIO_ATT_VALIDACION_CUENTAS()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() + " - OnStart");

            if (!this.Start())
            {
                EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                " - No se pudo iniciar el servicio. Revisar configuracion de parametros.", EventLogEntryType.Error);

                this.Ejecutar = false;
            }
            else
            {
                EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                " - Iniciando el metodo de trabajo del servicio.");

                this.Ejecutar = true;

            }

            if ((workerThread == null) || ((workerThread.ThreadState & (System.Threading.ThreadState.Unstarted | System.Threading.ThreadState.Stopped)) != 0))
            {
                workerThread = new Thread(new ThreadStart(ServiceWorkerMethod));
                workerThread.Start();
            }

            if (workerThread != null)
            {
                EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                    " - Estado del metodo de trabajo = " + workerThread.ThreadState.ToString());
            }
        }

        protected override void OnStop()
        {
            EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() + " - OnStop");

            EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                    " - Enviando señal de detencion de servicio.");

            this.RequestAdditionalTime(4000);

            // enviar señal de detencion del thread
            if ((workerThread != null) && (workerThread.IsAlive))
            {
                Thread.Sleep(5000);
                workerThread.Abort();
            }

            if (workerThread != null)
            {
                EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                    " - estado de detencion del metodo de trabajo = " + workerThread.ThreadState.ToString());

                // Indicate a successful exit.
                this.ExitCode = 0;
            }
        }

        /// <summary>
        /// Metodo de monitoreo de trabajo en segundo plano
        /// </summary>
        public void ServiceWorkerMethod()
        {

            if (!this.Ejecutar)
            {
                EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                    " - Deteniendo (Thread.CurrentThread.Abort) (ServiceWorkerMethod)");
                this.OnStop();
                return;
            }

            try
            {
                do
                {
                    Iniciar();

                    string espera = "0";
                    try
                    {
                        Config config = new Config();

                        espera = !string.IsNullOrEmpty(config.Intervalo.ToString()) ? config.Intervalo.ToString() : "86400000";
                    }
                    catch (Exception)
                    {
                        espera = "86400000";
                    }
                    Thread.Sleep(int.Parse(espera));

                    Application.DoEvents();
                }
                while (true);
            }
            catch (ThreadAbortException)
            {
                EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                    " - Señal de detencion de hilo en (ServiceWorkerMethod)");
            }

            EventLog.WriteEntry(this.seccion_evt, DateTime.Now.ToLongTimeString() +
                                " - Saliendo del metodo de trabajo (ServiceWorkerMethod)");
        }

        public bool Start()
        {
            return true;
        }

        /// <summary>
        /// Metodo de que inicia el servicio
        /// </summary>
        public bool Iniciar()
        {
            config = new Config();

            bool bret = true;
            string errores = string.Empty;
            string mensajeFinal = string.Empty;

            if (config.ValidaInfo())
            {
                try
                {
                    Monitor = new Tb_Monitor();
                    Monitor.IDAPPSERVICES = this.IdAplicacion;
                    log = new Log(this.config.LogServicio);
                    bd = new BD(this.config.ConUdl, this.config.Pkg);
                    bdLogs = new BD(this.config.ConUdl, this.config.PkgLogs);
                    log.Write("Inicio Servicio");

                    //Inicializo el correo
                    MailImap imap = new MailImap(this.config.CorreoServidor, this.config.UsuarioSMTP, this.config.PasswordSMTP, "993", this.config.PuertoSMTP.ToString());

                    try
                    {
                        Retorno ret = bdLogs.GetConfiguracionAplicacion(this.IdAplicacion);
                        if (ret.ret.Equals("OK"))
                        {
                            this.Configuracion = (List<ConfiguracionAplicacion>)ret.values[0];

                            this.ConfiguracionEjecucion = this.Configuracion.FirstOrDefault();

                            EjecutarServicio(this.ConfiguracionEjecucion);

                            if (this.modoActual.Equals("EJECUTAR"))
                            {
                                // this.LogActivo = ConfiguracionEjecucion.Log_Activo.Equals("0") || String.IsNullOrEmpty(ConfiguracionEjecucion.Log_Activo) ? false : true;

                                ret = bdLogs.GrabarMonitoreo(Monitor);
                                if (ret.ret.Equals("OK"))
                                {

                                    Monitor.ID = (string)ret.values[0];
                                    List<LOGS> LogsBd = new List<LOGS>(); // Es la lista de errores para el Login centralizado
                                    bret = ProcesarDatos(ref Monitor, ref LogsBd);

                                    mensajeFinal = this.config.CorreoMensaje + "\n" + (string)ret.values[0] ?? "";

                                    Retorno retLogs = GrabarLogs(LogsBd);
                                    if (retLogs.ret.Equals("OK"))
                                    {
                                        log.Write("Registro de Logs grabados en la Base de datos correctamente");
                                    }
                                    else
                                    {
                                        config2 = new Config();
                                        mensajeFinal = mensajeFinal + ", Grabando Logs: " + retLogs.msg;
                                        log.Write(retLogs.msg + ", error:" + retLogs.debug);
                                        string DestinarioPara = this.config2.CorreoPara.Split(';')[0];
                                        List<string> OtrosDestinatarios = this.config2.CorreoPara.Split(';').ToList();
                                        Retorno retMensaje = imap.EnviarMensajeConPlantillaLogCServicio(DestinarioPara, OtrosDestinatarios, new List<string>(), this.config2.CorreoAsunto, this.Sistema, mensajeFinal, Monitor.ID.ToString(), ref errores);
                                        if (!retMensaje.ret.Equals("OK"))
                                        {
                                            log.Write(retMensaje.msg + ", error:" + retMensaje.debug);
                                        }
                                    }

                                    ret = bdLogs.CerrarMonitoreo(Monitor);
                                    if (!ret.ret.Equals("OK"))
                                    {
                                        bret = false;
                                        log.Write(ret.msg + ", error:" + ret.debug);
                                    }
                                }
                                else
                                {
                                    bret = false;
                                    log.Write(ret.msg + ", error:" + ret.debug);
                                }
                            }
                        }
                        else
                        {
                            bret = false;
                            log.Write(ret.msg + ", error:" + ret.debug);
                        }
                    }
                    catch (Exception ex)
                    {
                        bret = false;
                        try
                        {
                            log.Write("Ocurrio una excepcion al ejecutar el servicio, error:" + ex.Message + " - Detalles del error: " + ex.StackTrace);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                EventLog.WriteEntry("SERVICIO_ATT_VALIDACION_CUENTAS", DateTime.Now.ToLongTimeString() + " Ocurrio un error al cerrar el trafico del servicio : " + ex.Message + " - Detalles del error: " + ex.StackTrace);
                            }
                            catch (Exception) { }
                        }
                    }
                    log.Write("Fin Servicio");
                }
                catch (Exception ex)
                {
                    bret = false;
                    try
                    {
                        log.Write("Ocurrio una excepción general en el servicio : " + ex.Message);
                        log.Write("Fin Servicio");
                    }
                    catch (Exception)
                    {
                        try
                        {
                            EventLog.WriteEntry("SERVICIO_ATT_VALIDACION_CUENTAS", DateTime.Now.ToLongTimeString() + " Ocurrio una excepción general en el servicio : " + ex.Message);
                        }
                        catch (Exception) { }
                    }
                }
            }
            else
            {
                bret = false;
                try
                {
                    log.Write("No se inicio el servicio por que la información del registro es invalida, se debe ejecutar el configurador del servicio");
                    log.Write("Fin Servicio");
                }
                catch (Exception)
                {
                    try
                    {
                        EventLog.WriteEntry("SERVICIO_ATT_VALIDACION_CUENTAS", DateTime.Now.ToLongTimeString() + " - No se inicio el servicio por que la información del registro es invalida, se debe ejecutar el configurador del servicio");
                    }
                    catch (Exception) { }
                }
            }

            return bret;
        }

        private void EjecutarServicio(ConfiguracionAplicacion configuracion)
        {
            try
            {

                // PASO 1 VERIFICO SI EL MICROSERVICIO ESTA CONFIGURADOR PARA EJECUTARSE A UNA HORA
                if (!String.IsNullOrWhiteSpace(configuracion.Hora_Agendada) && !configuracion.Hora_Agendada.Equals("-1"))
                {
                    var horaActual = DateTime.Now.ToString("HH");
                    var horas = configuracion.Hora_Agendada.Split(';');
                    DateTime FechaUltimaEjecucion = this.Libreria.GetFechaDataTime(configuracion.Fecha_Ultima_Ejecucion) ?? new DateTime(2000, 1, 1);

                    var horaUltimaEjecucion = FechaUltimaEjecucion.ToString("HH");

                    var HoraActualTicks = TimeSpan.FromTicks(DateTime.ParseExact(horaActual, "HH", CultureInfo.InvariantCulture).Ticks);
                    //var HoraUltimaEjecucion = TimeSpan.FromTicks(FechaUltimaEjecucion.TimeOfDay.Ticks);

                    // Debido a que la aplicacion se ejecuta en configuraciones horarias debo validar que la ultima hora en la que se ejecuto el dia actual
                    // no sea la mimsa que la hora actual, para evitar que la app se ejecute dos veces en la misma hora
                    bool EjecutadoHoyHora = false;
                    if (FechaUltimaEjecucion.Date == DateTime.Today && horaActual == horaUltimaEjecucion)
                    {
                        EjecutadoHoyHora = true;
                    }

                    // Si estoy configurado para ejecutarme a una hora y ya me ejecute durante esta hora entonces debo esperar y pasar la bandera a otro microservicio
                    if (EjecutadoHoyHora == true)
                    {
                        {
                            this.modoActual = "ESPERAR";
                        }
                    }
                    else
                    {
                        // Si no verifico si la hora actual esta dentro de las horas configuradas para ejecucion
                        foreach (var hora in horas)
                        {
                            var ParteHora = hora.Split(':');

                            string HoraEvaluar = ParteHora[0];

                            var HoraConfigurada = TimeSpan.FromTicks(DateTime.ParseExact(HoraEvaluar, "HH", CultureInfo.InvariantCulture).Ticks);

                            if (HoraActualTicks == HoraConfigurada)
                            {
                                this.modoActual = "EJECUTAR";
                                break;
                            }
                            else
                            {
                                this.modoActual = "ESPERAR";
                            }
                        }
                    }
                }
                else
                {
                    // PASO 2 VERIFICO SI EL MICROSERVICIO ESTA CONFIGURADO PARA EJECUTARSE CADA CIERTO TIEMPO
                    int tiempoEjecucion = Int32.Parse(configuracion.Segundos_Ejecucion);
                    int ultimaEjecucion = Int32.Parse(configuracion.Segundos_Ultima_Ejecucion);


                    if (ultimaEjecucion > tiempoEjecucion)
                    {
                        this.modoActual = "EJECUTAR";
                    }
                    else
                    {
                        this.modoActual = "ESPERAR";
                    }
                }
            }
            catch (Exception ex)
            {
                log.Write("Ocurrio una excepcion al ejecutar metodo EjecutarServicio MicroServicio: " + configuracion.Id_Aplicacion + " - " + configuracion.Aplicacion + " : error" + ex.Message + " - Detalles del error" + ex.StackTrace);
                this.modoActual = "CERRAR_TRAFICO";
            }
        }

        private Boolean ProcesarDatos(ref Tb_Monitor Lectura, ref List<LOGS> LogsBd)

        {
            Retorno ret = new Retorno();
            try
            {
                Lectura.IDSTATE = "4";
                Log log = new Log(this.config.LogServicio);
                MailImap imap;
                List<ValidacionCuentasTempServicio> ListadoValidar = new List<ValidacionCuentasTempServicio>();
                List<ListadoEmailMensajes> ListadoEmails = new List<ListadoEmailMensajes>();
                List<ValidacionCuentasTempServicio> Resultados = new List<ValidacionCuentasTempServicio>();
                string Errores = string.Empty;
                ret = bd.GetListadoCuentas();
                if (ret.ret.Equals("OK"))
                {
                    ListadoValidar = (List<ValidacionCuentasTempServicio>)ret.values[0];
                    if (ListadoValidar.Count == 0)
                    {
                        log.Write("No hay cuentas de correo que validar");
                        //Lectura.IDSTATE = "10"; // SIN DATOS QUE PROCESAR
                        //return true;
                    }
                }
                else
                {
                    log.Write(ret.msg + ", error:" + ret.debug);
                    Lectura.IDSTATE = "3"; // ERROR EN LA INSERCCION DE REGISTROS EN EL MODELO LOGIN CENTRALIZADO
                    return false;
                }

                //string path = Limilabs.Mail.Licensing.LicenseHelper.GetLicensePath();
                //log.Write("license path: " + path);

                // Obtengo la configuracion del SMTP e IMAP para este servicio
                ConfiguracionAplicacion configuracionMail = this.Configuracion.Where(x => x.Id_Aplicacion.Equals(this.IdAplicacion) && x.Id_Configuracion.Equals("5")).FirstOrDefault();

                //Leo los Correos electronicos de la bandeja de entrada
                imap = new MailImap(configuracionMail.Parametro1, configuracionMail.Parametro2, configuracionMail.Parametro3, configuracionMail.Parametro4, configuracionMail.Parametro5);
                ret = imap.LeerCasillaCorreos();
                if (ret.ret.Equals("OK"))
                {
                    ListadoEmails = (List<ListadoEmailMensajes>)ret.values[0];
                    if (ListadoValidar.Count == 0)
                    {
                        log.Write("No hay mensajes de correo nuevos por leer");
                    }
                }
                else
                {
                    log.Write(ret.msg + ", error:" + ret.debug);
                    Lectura.IDSTATE = "2"; // ERROR EN LA EJECUCION DE LA API O FTP, RESULTADO DE RETORNO INVALIDO
                    return false;
                }


                foreach (ValidacionCuentasTempServicio Cuenta in ListadoValidar)
                {
                    // Si es la primera vez que enviare el mensaje
                    if (Cuenta.INTERACCION == 0)
                    {
                        string Asunto = "Sistema de validación de correos personales";
                        string Mensaje = "Le hemos enviado un correo para confirmar que esta dirección de correo electrónico es la suya. Por favor, le solicitamos que responda a este mensaje para confirmar la validez de su dirección. Si ha recibido este correo y no reconoce su origen, le pedimos que lo omita y lo elimine.";
                        ret = imap.EnviarMensajeConPlantillaLogCID(Cuenta.EMAIL_PERSONAL, new List<string>(), new List<string>(), Asunto, Mensaje, Cuenta.NOMBRES_TRABAJADOR + " " + Cuenta.APELLIDOS_TRABAJADOR, ref Errores);
                        if (ret.ret.Equals("OK"))
                        {
                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            Cuenta.INTERACCION++;
                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            Cuenta.ID_ESTADO = 1;
                            Cuenta.ESTADO = "EN PROCESO";
                            Cuenta.OBSERVACIONES = "Correo enviado con exito";
                        }
                        else if (ret.ret.Equals("ERROR_EMAIL"))
                        {
                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            Cuenta.INTERACCION++;
                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            Cuenta.ID_ESTADO = 5;
                            Cuenta.ESTADO = "NO VALIDOS";
                            Cuenta.OBSERVACIONES = "El Correo electronico no es valido";
                            log.Write(ret.msg + " - " + ret.debug);
                        }
                        else
                        {
                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            Cuenta.INTERACCION++;
                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            Cuenta.ID_ESTADO = 5;
                            Cuenta.ESTADO = "NO VALIDOS";
                            Cuenta.OBSERVACIONES = "Ocurrio un error al enviar el mensaje a esta cuenta";
                            log.Write(ret.msg + " - " + ret.debug);
                        }
                    }
                    else
                    {
                        bool encontrado = false; // Si no encuentro una respuesta del trabajador y han pasado dos dias hago un segundo intento de envio
                        foreach (ListadoEmailMensajes mensaje in ListadoEmails)
                        {
                            // Si recibi un mensaje por parte del trabajador
                            if (mensaje.From.ToUpper().Contains(Cuenta.EMAIL_PERSONAL))
                            {
                                encontrado = true;
                                Cuenta.INTERACCION++;
                                Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                Cuenta.ID_ESTADO = 2;
                                Cuenta.ESTADO = "VALIDADOS";
                                Cuenta.FECHA_RESPUESTA = mensaje.FechaRecepcion;
                                string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "Se recibio la respuesta del trabajador y valido correctamente";
                                Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                break;
                            }
                            else // Si el que envia el mensaje no es el trabajador, verifico si el mensaje trae el correo dentro, una señal de que no se pudo enviar el correo
                            if (mensaje.Text.ToUpper().Contains(Cuenta.EMAIL_PERSONAL))
                            {
                                encontrado = true;
                                Cuenta.INTERACCION++;
                                Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                Cuenta.ID_ESTADO = 5;
                                Cuenta.ESTADO = "NO VALIDOS";
                                Cuenta.FECHA_RESPUESTA = mensaje.FechaRecepcion;
                                string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "No se pudo validar el correo";
                                Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                //string CorreoJefe = !string.IsNullOrWhiteSpace(Cuenta.MANAGER_EMAIL) ? Cuenta.MANAGER_EMAIL : Cuenta.EMAIL_JEFATURA;
                                string CorreoJefe = Cuenta.EMAIL_JEFATURA;
                                string Asunto = "Sistema de validación de correos personales";
                                string Mensaje = "No se pudo Validar la cuenta del trabajador" + "\n" + "\n" + "Identificacion: " + Cuenta.DNI_RUT + "\n" + "Nombres: " + Cuenta.NOMBRES_TRABAJADOR + " " + Cuenta.APELLIDOS_TRABAJADOR + "\n" + "Correo Personal:" + Cuenta.EMAIL_PERSONAL;

                                ret = imap.EnviarMensajeConPlantillaLogCID(CorreoJefe, new List<string>(), new List<string>(), Asunto, Mensaje, Cuenta.MANAGER_NAME, ref Errores);
                                if (!ret.ret.Equals("OK"))
                                {
                                    Cuenta.OBSERVACIONES += "\n" + "No se pudo enviar el mensaje al jefe";
                                }
                                break;
                            }
                        }

                        if (encontrado == false)
                        {
                            try
                            {
                                if (Cuenta.INTERACCION == 1)
                                {
                                    if (Cuenta.HAN_PASADO_DOS_DIAS.Equals("SI"))
                                    {
                                        ret = imap.EnviarMensajeConPlantillaLogCID(Cuenta.EMAIL_PERSONAL, new List<string>(), new List<string>(), "Sistema de validación de correos personales", "Sistema de validación de correos personales, se hace un segundo intento por validar su cuenta personal", Cuenta.NOMBRES_TRABAJADOR + " " + Cuenta.APELLIDOS_TRABAJADOR, ref Errores);
                                        if (ret.ret.Equals("OK"))
                                        {
                                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.INTERACCION++;
                                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.ID_ESTADO = 3;
                                            Cuenta.ESTADO = "REINTENTO";
                                            string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "Se envia el correo de validacion por segunda ocasion para validar la cuenta";
                                            Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                        }
                                        else if (ret.ret.Equals("ERROR_EMAIL"))
                                        {
                                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.INTERACCION++;
                                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-dd-MM HH:mm:ss");
                                            Cuenta.ID_ESTADO = 5;
                                            Cuenta.ESTADO = "NO VALIDOS";
                                            string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "El Correo electronico no es valido";
                                            Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                            log.Write(ret.msg + " - " + ret.debug);
                                        }
                                        else
                                        {
                                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.INTERACCION++;
                                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-dd-MM HH:mm:ss");
                                            Cuenta.ID_ESTADO = 5;
                                            Cuenta.ESTADO = "NO VALIDOS";
                                            string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "Ocurrio un error al enviar el mensaje a esta cuenta";
                                            Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                            log.Write(ret.msg + " - " + ret.debug);
                                        }
                                    }
                                }
                                else if (Cuenta.INTERACCION == 2)
                                {
                                    if (Cuenta.ENVIAR_MENSAJE_JEFE.Equals("ENVIAR_MENSAJE"))
                                    {
                                        ret = imap.EnviarMensajeConPlantillaLogCID(Cuenta.EMAIL_JEFATURA, new List<string>(), new List<string>(), "Sistema de validación de correos personales", "Sistema de validación de correos personales, no se recibe respuesta de validacion del correo: " + Cuenta.EMAIL_PERSONAL + ", correspondiente a:  " + Cuenta.NOMBRES_TRABAJADOR + " " + Cuenta.APELLIDOS_TRABAJADOR + ".", " ", ref Errores);
                                        if (ret.ret.Equals("OK"))
                                        {
                                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.INTERACCION++;
                                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.ID_ESTADO = 4;
                                            Cuenta.ESTADO = "ACUSE";
                                            string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "Se envia el correo de validacion por segunda ocasion para validar la cuenta";
                                            Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                        }
                                        else if (ret.ret.Equals("ERROR_EMAIL"))
                                        {
                                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.INTERACCION++;
                                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-dd-MM HH:mm:ss");
                                            Cuenta.ID_ESTADO = 5;
                                            Cuenta.ESTADO = "NO VALIDOS";
                                            string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "El Correo electronico no es valido";
                                            Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                            log.Write(ret.msg + " - " + ret.debug);
                                        }
                                        else
                                        {
                                            Cuenta.FECHA_ENVIO = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            Cuenta.INTERACCION++;
                                            Cuenta.FECHA_ULTIMA_INTERACCION = DateTime.Now.ToString("yyyy-dd-MM HH:mm:ss");
                                            Cuenta.ID_ESTADO = 5;
                                            Cuenta.ESTADO = "NO VALIDOS";
                                            string mensajeTexto = Cuenta.OBSERVACIONES + "\n" + "Ocurrio un error al enviar el mensaje a esta cuenta";
                                            Cuenta.OBSERVACIONES = (mensajeTexto).Length <= 3999 ? mensajeTexto : mensajeTexto.Substring(0, 3999);
                                            log.Write(ret.msg + " - " + ret.debug);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                log.Write("Ocurrio una excepcion en la rutina de verificación de envio de mensaje Jefe, error: " + ex.Message + " - Detalles del error: " + ex.StackTrace);
                            }
                        }
                    }

                    Cuenta.OBSERVACIONES = Cuenta.OBSERVACIONES.ToUpper();
                    Resultados.Add(Cuenta);
                }

                if (Resultados.Count > 0)
                {
                    ret = bd.CrearRespaldo();
                    if (ret.ret.Equals("OK"))
                    {
                        log.Write("Se creo el respaldo correctamente");

                        ret = GrabarDatos(Resultados);
                        if (ret.ret.Equals("OK"))
                        {
                            log.Write("Se grabaron los resultados de las cuentas correctamente");

                            ret = bd.ProcesarDatos();
                            if (ret.ret.Equals("OK"))
                            {
                                log.Write("Se procesaron los resultados de las cuentas correctamente");
                            }
                            else
                            {
                                Retorno retBack = bd.RollBack();
                                if (retBack.ret.Equals("OK"))
                                {
                                    log.Write(ret.msg + ", error:" + ret.debug);
                                    Lectura.IDSTATE = "3"; // ERROR EN LA INSERCCION DE REGISTROS EN EL MODELO LOGIN CENTRALIZADO
                                    log.Write("Se realizo el rollback correctamente");
                                    return false;
                                }
                                {
                                    log.Write(retBack.msg + ", error:" + retBack.debug);
                                    Lectura.IDSTATE = "3"; // ERROR EN LA INSERCCION DE REGISTROS EN EL MODELO LOGIN CENTRALIZADO
                                    return false;
                                }
                            }
                        }
                        else
                        {
                            log.Write(ret.msg + ", error:" + ret.debug);
                            Lectura.IDSTATE = "3"; // ERROR EN LA INSERCCION DE REGISTROS EN EL MODELO LOGIN CENTRALIZADO
                            return false;
                        }
                    }
                    else
                    {
                        log.Write(ret.msg + ", error:" + ret.debug);
                        Lectura.IDSTATE = "3"; // ERROR EN LA INSERCCION DE REGISTROS EN EL MODELO LOGIN CENTRALIZADO
                        return false;
                    }
                }
                else
                {
                    Lectura.IDSTATE = "10"; // SIN DATOS QUE PROCESAR
                    return false;
                }
                //Aqui debo grabar y procesar los resultados

            }
            catch (Exception ex)
            {
                log.Write("Ocurrio una excepcion en la rutina procesa, error: " + ex.Message + " - Detalles del error: " + ex.StackTrace);
            }


            return true;
        }

        #region BASE DE DATOS


        /// <summary>
        /// Graba los datos de la tabla TB_ADACTIVITY_TEMP en Base de Datos
        /// </summary>
        /// <returns> retorna un objeto Retorno</returns>
        public Retorno GrabarDatos(List<ValidacionCuentasTempServicio> Listado)
        {
            Retorno ret = new Retorno();
            //string tabla = "TB_SERVACCOUNTVALIDATION";
            //DataTable dt = Libreria.ToDataTable(Listado);
            //ret = CargaMasiva(tabla, dt);
            //return ret;

            string columnas = "ID,DNI_RUT,NOMBRES_TRABAJADOR,APELLIDOS_TRABAJADOR,EMAIL_PERSONAL,EMAIL_JEFATURA,MANAGER_RUT,MANAGER_NAME,MANAGER_EMAIL,INTERACCION,FECHA_ULTIMA_INTERACCION,ID_ESTADO,ESTADO,FECHA_RESPUESTA,FECHA_ENVIO,OBSERVACIONES";
            string tabla = "TB_ATT_SERV_ACCOUNT_VALIDATION";
            DataTable dt = Libreria.ToDataTable(Listado);
            ret = CargaMasivaViejo(tabla, columnas, dt);
            return ret;
        }

        public Retorno CargaMasiva(string tabla, DataTable dt)
        {
            Retorno ret = new Retorno();
            try
            {
                StringBuilder parametros = new StringBuilder();
                string error = String.Empty;

                //DataSet set = new DataSet();
                //set.Tables.Add(dt);

                bool result = bd.MultiInsertDataNuevo(dt, tabla, ref error);
                if (result)
                {
                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty;
                }
                else
                {
                    ret.ret = "ERROR";
                    ret.msg = "Fallo al realizar carga: " + error;
                    ret.debug = error;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepcion al realizar carga: " + ex.Message;
                ret.debug = ex.Message;
            }
            return ret;
        }

        ///// <summary>
        ///// Consume el Metodo MultiInsertData para añadir valores en lote a la BD a una tabla en especifico
        ///// </summary>
        ///// <returns> retorna un objeto DataTable</returns>
        //public Retorno CargaMasiva(string tabla, DataTable dt)
        //{
        //    Retorno ret = new Retorno();
        //    try
        //    {
        //        StringBuilder parametros = new StringBuilder();
        //        string error = String.Empty;

        //        bool result = bd.MultiInsertData(dt, tabla, ref error);
        //        if (result)
        //        {
        //            ret.ret = "OK";
        //            ret.msg = String.Empty;
        //            ret.debug = String.Empty;
        //        }
        //        else
        //        {
        //            ret.ret = "ERROR";
        //            ret.msg = "Fallo al realizar carga";
        //            ret.debug = error;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ret.ret = "ERROR";
        //        ret.msg = "Excepcion al realizar carga: " + ex.Message + " - Detalles del error: " + ex.StackTrace;
        //        ret.debug = ex.Message;
        //    }
        //    return ret;
        //}
        #endregion BASE DE DATOS

        #region LOG BD

        /// <summary>
        /// Graba los datos de la tabla TB_LOGS en Base de Datos
        /// </summary>
        /// <returns> retorna un objeto Retorno</returns>
        public Retorno GrabarLogs(List<LOGS> Listado)
        {
            Retorno ret = new Retorno();
            if (this.ConfiguracionEjecucion.Log_Activo.Equals("1") && Listado.Count > 0)
            {
                string columnas = "IDMONITOR,IDDATASOURCE,IDLOGPROCESS,ERROR,MEASURE";
                string tabla = "TB_LOGS";
                DataTable dt = Libreria.ToDataTable(Listado);
                ret = CargaMasivaViejo(tabla, columnas, dt);
            }
            else
            {
                ret.ret = "OK";
                ret.msg = String.Empty;
                ret.debug = String.Empty;
            }
            return ret;
        }

        /// <summary>
        /// Añade un registro a a lista de Logs del proceso
        /// </summary>
        /// <param name="LogsBd">List<LOGS> este parametro Es una listado de registro del tipo LOGS, que seran guardado en BD</param>
        /// <param name="Mensaje">string Contiene el mensaje de error</param>
        /// <param name="Solucion">string Contiene un mensaje con la posible solucion del error</param>
        /// <returns> retorna true si es correcto el proceso, false en caso de error o en caso que el servicio este apagado para grabar registros logs en bd</returns>
        public Boolean AñadeLog(ref List<LOGS> LogsBd, string Mensaje, string Solucion)
        {
            if (this.ConfiguracionEjecucion.Log_Activo.Equals("1"))
            {
                try
                {
                    LOGS Error = new LOGS
                    {
                        IDMONITOR = Monitor.ID,
                        IDDATASOURCE = this.DataSource,
                        IDLOGPROCESS = "1",
                        MEASURE = Solucion,
                        ERROR = Mensaje
                    };
                    LogsBd.Add(Error);
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Consume el Metodo MultiInsertData para añadir valores en lote a la BD a una tabla en especifico
        /// </summary>
        /// <returns> retorna un objeto DataTable</returns>
        public Retorno CargaMasivaViejo(string tabla, string columnas, DataTable dt)
        {
            Retorno ret = new Retorno();
            try
            {
                StringBuilder parametros = new StringBuilder();
                string error = String.Empty;

                DataSet set = new DataSet();
                set.Tables.Add(dt);

                bool result = bdLogs.MultiInsertData(set, columnas, tabla, ref error);
                if (result)
                {
                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty;
                }
                else
                {
                    ret.ret = "ERROR";
                    ret.msg = "Fallo al realizar carga: " + error;
                    ret.debug = error;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepcion al realizar carga: " + ex.Message;
                ret.debug = ex.Message;
            }
            return ret;
        }
        #endregion LOG BD
    }
}
