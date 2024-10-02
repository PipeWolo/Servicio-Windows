using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ServiceMail
{
    #region CLASES GENERICAS
    public class Retorno
    {
        public string ret { get; set; } = "OK";
        public string msg { get; set; } = String.Empty;
        public string debug { get; set; } = String.Empty;
        public List<object> values { get; set; } = new List<object>();
    }

    class KVP
    {
        public string KeyName { get; set; }
        public string KeyValue { get; set; }
        public string KeyValue2 { get; set; }

        public KVP()
        {

        }

        public KVP(string keyname, string keyvalue)
        {
            this.KeyName = keyname;
            this.KeyValue = keyvalue;
        }
    }

    public class Configuracion
    {
        public string Id { get; set; }
        public string IdEmpresa { get; set; }
        public string Empresa { get; set; }
        public string IdServicio { get; set; }
        public string Servicio { get; set; }
        public string Dominio { get; set; }
        public string Puerto { get; set; }
        public string IdTipo { get; set; }
        public string Tipo { get; set; }
        public string Usuario { get; set; }
        public string Password { get; set; }
        public string Ssl { get; set; }
        public string Carga { get; set; }
        public string PuertoStmp { get; set; }
        public string AsuntoStmp { get; set; }
        public string AsuntoStmpNuevosCorreos { get; set; }
        public string NoReply { get; set; }
        public string DominioSalida { get; set; }
        public string UltimoId { get; set; }
        public string UIDValidity { get; set; }
    }

    public class ApiResponse
    {
        public ApiResponseValue response { get; set; }
    }
    public class ApiResponseValue
    {
        public string status { get; set; }
        public string msg { get; set; }
        public string values { get; set; }
    }
    public class ApiFilter
    {
        public string fcm { get; set; }
        public string title { get; set; }
        public string message { get; set; }
    }

    public class Notificacion
    {
        public string fcm { get; set; }
        public string title { get; set; }
        public string message { get; set; }
    }

    #endregion GENERICAS

    #region CLASES DEL SERVICIO

    // Con esta Clase leeo los correos por validar
    public class ValidacionCuentasTempServicio
    {
        public int ID { get; set; }
        public string DNI_RUT { get; set; }
        public string NOMBRES_TRABAJADOR { get; set; }
        public string APELLIDOS_TRABAJADOR { get; set; }
        public string EMAIL_PERSONAL { get; set; }
        public string EMAIL_JEFATURA { get; set; }
        public string MANAGER_RUT { get; set; }
        public string MANAGER_NAME { get; set; }
        public string MANAGER_EMAIL { get; set; }
        public int INTERACCION { get; set; }
        public string FECHA_ULTIMA_INTERACCION { get; set; }
        public string ENVIAR_MENSAJE_JEFE { get; set; }
        public string HAN_PASADO_DOS_DIAS { get; set; }
        public int ID_ESTADO { get; set; }
        public string ESTADO { get; set; }
        public string FECHA_RESPUESTA { get; set; }
        public string FECHA_ENVIO { get; set; }
        public string OBSERVACIONES { get; set; }

    }

    public class ListadoEmailMensajes
    {
        public long Uid { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Subject { get; set; }
        public string Text { get; set; }
        public string FechaRecepcion { get; set; }
    }

    #endregion CLASES DEL SERVICIO

    #region CLASE DE MONITOREO
    public class Tb_Monitor
    {
        public string ID { get; set; }
        public string IDSITE { get; set; } = "1";
        public string IDSTATE { get; set; } = "1";// campo que representa el id del registro vinculado en la tabla TB_STATE
        public string IDAPPSERVICES { get; set; } // este es el ID que se asigno a esta app en la tabla TB_APPSERVICES
    }

    public class ConfiguracionServicio
    {
        public string LogActivo { get; set; }
    }

    public class ConfiguracionDataSource
    {
        public string Resultado { get; set; }
        public string UltimaEjecucion { get; set; } // En Segundos
    }

    public class LOGS
    {
        public string IDMONITOR { get; set; } // campo que representa el id del registro vinculado en la tabla TB_MONITOR
        public string IDDATASOURCE { get; set; } // campo que representa el id del registro vinculado en la tabla TB_DATASOURCE
        public string IDLOGPROCESS { get; set; } // campo que representa el id del registro vinculado en la tabla TB_LOGPROCESS
        public string ERROR { get; set; } // Campo donde se guarda el error mismo
        public string MEASURE { get; set; } // Medida correctiva propuesta para la correccion
    }

    public class ConfiguracionAplicacion
    {
        public string Id { get; set; }
        public string Descripcion { get; set; }
        public string Id_Aplicacion { get; set; }
        public string Aplicacion { get; set; }
        public string Id_Configuracion { get; set; }
        public string Desc_Configuracion { get; set; }
        public string Parametro1 { get; set; }
        public string Parametro2 { get; set; }
        public string Parametro3 { get; set; }
        public string Parametro4 { get; set; }
        public string Parametro5 { get; set; }
        public string Parametro6 { get; set; }
        public string Parametro7 { get; set; }
        public string Parametro8 { get; set; }
        public string Parametro9 { get; set; }
        public string Parametro10 { get; set; }
        public string Parametro11 { get; set; }
        public string Parametro12 { get; set; }
        public string Parametro13 { get; set; }
        public string Parametro14 { get; set; }
        public string Parametro15 { get; set; }
        public string Parametro16 { get; set; }
        public string Parametro17 { get; set; }
        public string Parametro18 { get; set; }
        public string Parametro19 { get; set; }
        public string Parametro20 { get; set; }
        public string Parametro21 { get; set; }
        public string Parametro22 { get; set; }
        public string Parametro23 { get; set; }
        public string Parametro24 { get; set; }
        public string Parametro25 { get; set; }
        public string Fecha_Ultima_Ejecucion { get; set; }
        public string Log_Activo { get; set; }
        public string Segundos_Ultima_Ejecucion { get; set; }
        public string Hora_Agendada { get; set; }
        public string Segundos_Ejecucion { get; set; }
        public string Mover_Log { get; set; }
    }
    #endregion MONITOREO
}


