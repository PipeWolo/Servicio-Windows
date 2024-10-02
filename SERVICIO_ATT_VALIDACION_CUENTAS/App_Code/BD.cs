using System;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Linq;
using ServiceMail;
using SERVICIO_ATT_VALIDACION_CUENTAS.App_Code;
using System.Data.OracleClient;
using System.Globalization;
using Limilabs.Mail.Headers;
using Limilabs.Mail;

namespace SERVICIO_ATT_VALIDACION_CUENTAS
{
    class BD
    {
        public string Conexion { get; set; }
        public string Package { get; set; }

        public BD(string conexion, string package)
        {
            this.Conexion = conexion;
            this.Package = package;
        }


        #region MONITOREO

        public Retorno GetConfiguracionAplicacion(string IdAplicacion)
        {

            Retorno ret = new Retorno();

            try
            {
                NavigatorLibreria Libreria = new NavigatorLibreria();
                Oracle.ManagedDataAccess.Client.OracleConnection con = new Oracle.ManagedDataAccess.Client.OracleConnection();
                con.ConnectionString = this.Conexion;
                con.Open();
                Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand();

                cmd.Connection = con;
                cmd.CommandText = this.Package + ".GET_CONFIGURACION_APP2";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 100;
                cmd.Parameters.Add("PID_APLICACION", OracleDbType.Int32).Value = IdAplicacion;
                cmd.Parameters.Add("IO_CURSOR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter(cmd);

                DataSet ds = new DataSet();
                da.Fill(ds);

                string error = cmd.Parameters["PERROR"].Value != null ? cmd.Parameters["PERROR"].Value.ToString() : "";

                if (error.Equals("OK"))
                {
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        List<ConfiguracionAplicacion> Configuracion = ds.Tables[0].AsEnumerable().Select(item => new ConfiguracionAplicacion
                        {
                            Id = Convert.ToString(item.Field<decimal?>("ID")),
                            Descripcion = item.Field<string>("DESCRIPCION"),
                            Id_Aplicacion = Convert.ToString(item.Field<decimal?>("ID_APLICACION")),
                            Aplicacion = item.Field<string>("APLICACION"),
                            Id_Configuracion = Convert.ToString(item.Field<decimal?>("ID_CONFIGURACION")),
                            Desc_Configuracion = item.Field<string>("DESC_CONFIGURACION"),
                            Parametro1 = item.Field<string>("PARAMETRO1"),
                            Parametro2 = item.Field<string>("PARAMETRO2"),
                            Parametro3 = item.Field<string>("PARAMETRO3"),
                            Parametro4 = item.Field<string>("PARAMETRO4"),
                            Parametro5 = item.Field<string>("PARAMETRO5"),
                            Parametro6 = item.Field<string>("PARAMETRO6"),
                            Parametro7 = item.Field<string>("PARAMETRO7"),
                            Parametro8 = item.Field<string>("PARAMETRO8"),
                            Parametro9 = item.Field<string>("PARAMETRO9"),
                            Parametro10 = item.Field<string>("PARAMETRO10"),
                            Parametro11 = item.Field<string>("PARAMETRO11"),
                            Parametro12 = item.Field<string>("PARAMETRO12"),
                            Parametro13 = item.Field<string>("PARAMETRO13"),
                            Parametro14 = item.Field<string>("PARAMETRO14"),
                            Parametro15 = item.Field<string>("PARAMETRO15"),
                            Fecha_Ultima_Ejecucion = item.Field<string>("FECHA_ULTIMA_EJECUCION"),
                            Log_Activo = Convert.ToString(item.Field<decimal?>("LOG_ACTIVO")),
                            Segundos_Ultima_Ejecucion = Convert.ToString(item.Field<decimal?>("SEGUNDOS_ULTIMA_EJECUCION")),
                            Hora_Agendada = item.Field<string>("HORA_AGENDADA"),
                            Segundos_Ejecucion = Convert.ToString(item.Field<decimal?>("SEGUNDOS_EJECUCION")),
                            Mover_Log = Convert.ToString(item.Field<decimal?>("MOVER_LOG"))
                        }).ToList();

                        ret.values.Add(Configuracion);
                    }
                    else
                    {
                        ret.ret = "ERROR";
                        ret.msg = "La Aplicacion no tiene configuracion en la base de datos";
                    }
                }
                else
                {
                    ret.ret = "ERROR";
                    ret.msg = "Ocurrio un error al consultar la configuracion de la aplicacion";
                    ret.debug = error;
                }

                if (con.State != ConnectionState.Closed)
                {
                    con.Close();
                }
                con.Dispose();
                cmd.Dispose();
                con = null;
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al obtener la configuracion del servicio";
                ret.debug = "Excepción al obtener la configuracion del servicio, Error: " + ex.Message;
            }


            return ret;
        }
        public Retorno GetConfiguracionAplicacionOLD(string IdAplicacion)
        {

            Retorno ret = new Retorno();

            try
            {
                using (Oracle.ManagedDataAccess.Client.OracleConnection con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    con.Open();
                    using (Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".GET_CONFIGURACION_APP", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 300; // 5 minutes (300 seconds)
                        cmd.Parameters.Add("PID_APLICACION", OracleDbType.Int32).Value = IdAplicacion;
                        cmd.Parameters.Add("IO_CURSOR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 1024).Direction = ParameterDirection.Output;

                        using (Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter(cmd))
                        {
                            DataSet ds = new DataSet();
                            da.Fill(ds);

                            string error = cmd.Parameters["PERROR"].Value != null ? cmd.Parameters["PERROR"].Value.ToString() : "";

                            if (error.Equals("OK"))
                            {
                                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    List<ConfiguracionAplicacion> Configuracion = ds.Tables[0].AsEnumerable().Select(item => new ConfiguracionAplicacion
                                    {
                                        Id = Convert.ToString(item.Field<decimal?>("ID")),
                                        Descripcion = item.Field<string>("DESCRIPCION"),
                                        Id_Aplicacion = Convert.ToString(item.Field<decimal?>("ID_APLICACION")),
                                        Aplicacion = item.Field<string>("APLICACION"),
                                        Id_Configuracion = Convert.ToString(item.Field<decimal?>("ID_CONFIGURACION")),
                                        Desc_Configuracion = item.Field<string>("DESC_CONFIGURACION"),
                                        Parametro1 = item.Field<string>("PARAMETRO1"),
                                        Parametro2 = item.Field<string>("PARAMETRO2"),
                                        Parametro3 = item.Field<string>("PARAMETRO3"),
                                        Parametro4 = item.Field<string>("PARAMETRO4"),
                                        Parametro5 = item.Field<string>("PARAMETRO5"),
                                        Parametro6 = item.Field<string>("PARAMETRO6"),
                                        Parametro7 = item.Field<string>("PARAMETRO7"),
                                        Parametro8 = item.Field<string>("PARAMETRO8"),
                                        Parametro9 = item.Field<string>("PARAMETRO9"),
                                        Parametro10 = item.Field<string>("PARAMETRO10"),
                                        Parametro11 = item.Field<string>("PARAMETRO11"),
                                        Parametro12 = item.Field<string>("PARAMETRO12"),
                                        Parametro13 = item.Field<string>("PARAMETRO13"),
                                        Parametro14 = item.Field<string>("PARAMETRO14"),
                                        Parametro15 = item.Field<string>("PARAMETRO15"),
                                        Fecha_Ultima_Ejecucion = item.Field<string>("FECHA_ULTIMA_EJECUCION"),
                                        Log_Activo = Convert.ToString(item.Field<decimal?>("LOG_ACTIVO")),
                                        Segundos_Ultima_Ejecucion = Convert.ToString(item.Field<decimal?>("SEGUNDOS_ULTIMA_EJECUCION")),
                                        Hora_Agendada = item.Field<string>("HORA_AGENDADA"),
                                        Segundos_Ejecucion = Convert.ToString(item.Field<decimal?>("SEGUNDOS_EJECUCION")),
                                        Mover_Log = Convert.ToString(item.Field<decimal?>("MOVER_LOG"))
                                    }).ToList();

                                    ret.values.Add(Configuracion);
                                }
                                else
                                {
                                    ret.ret = "ERROR";
                                    ret.msg = "La Aplicacion no tiene configuracion en la base de datos";
                                }
                            }
                            else
                            {
                                ret.ret = "ERROR";
                                ret.msg = "Ocurrio un error al consultar la configuracion de la aplicacion";
                                ret.debug = error;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Ocurrio una excepcion al consultar la configuracion de la aplicacion";
                ret.debug = ex.Message + " - Detalles del error: " + ex.StackTrace;
            }

            return ret;
        }

        public Retorno GrabarMonitoreo(Tb_Monitor Monitor)
        {
            try
            {

                using (Oracle.ManagedDataAccess.Client.OracleConnection con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    using (Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".GRABAR_MONITOREO", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 300; // 5 minutes (300 seconds)

                        cmd.Parameters.Add("PIDSITE", OracleDbType.Int32).Value = Monitor.IDSTATE;
                        cmd.Parameters.Add("PIDAPPSERVICES", OracleDbType.Int32).Value = Monitor.IDAPPSERVICES;
                        cmd.Parameters.Add("IO_CURSOR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 1024).Direction = ParameterDirection.Output;

                        con.Open();

                        using (Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter(cmd))
                        {
                            DataSet ds = new DataSet();
                            da.Fill(ds);

                            string error = cmd.Parameters["PERROR"].Value != null ? cmd.Parameters["PERROR"].Value.ToString() : "";

                            if (error.Equals("OK"))
                            {
                                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    Retorno ret = new Retorno();
                                    string IdMonitoreo = ds.Tables[0].Rows[0]["ID_MONITOREO"].ToString();
                                    ret.values.Add(IdMonitoreo);
                                    return ret;
                                }
                                else
                                {
                                    return new Retorno
                                    {
                                        ret = "ERROR",
                                        msg = "La consulta no retornó ningún registro de monitoreo",
                                        debug = String.Empty
                                    };
                                }
                            }
                            else
                            {
                                return new Retorno
                                {
                                    ret = "ERROR",
                                    msg = "Ocurrió un error al obtener el registro de Monitoreo",
                                    debug = error
                                };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrió una excepcion al obtener el registro de Monitoreo",
                    debug = ex.Message + " - Detales del error: " + ex.StackTrace
                };
            }
        }

        public Retorno CerrarMonitoreo(Tb_Monitor Monitor)
        {
            try
            {
                using (var con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    using (var cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".CERRAR_MONITOREO", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("PID_MONITOREO", OracleDbType.Int32).Value = Monitor.ID;
                        cmd.Parameters.Add("PIDSTATE", OracleDbType.Int32).Value = Monitor.IDSTATE;
                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 1024).Direction = ParameterDirection.Output;

                        con.Open();
                        cmd.ExecuteNonQuery();

                        string error = cmd.Parameters["PERROR"].Value.ToString();
                        if (error == "OK")
                        {
                            return new Retorno();
                        }
                        else
                        {
                            return new Retorno
                            {
                                ret = "ERROR",
                                msg = "Ocurrió un error al cerrar el Monitoreo",
                                debug = error
                            };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrió una excepcion al cerrar el Monitoreo",
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
        }

        public Retorno CrearRespaldo()
        {
            try
            {
                using (var con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    using (var cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".CREAR_RESPALDO", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 4000).Value = Oracle.ManagedDataAccess.Types.OracleString.Null;
                        cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                        con.Open();
                        cmd.ExecuteNonQuery();

                        string error = cmd.Parameters["PERROR"].Value.ToString();
                        if (error == "OK")
                        {
                            return new Retorno();
                        }
                        else
                        {
                            return new Retorno
                            {
                                ret = "ERROR",
                                msg = "Ocurrio un error al crear el respaldo",
                                debug = error
                            };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrio una excepcion al crear el respaldo",
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
        }

        public Retorno RollBack()
        {
            try
            {
                using (var con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    using (var cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".ROLL_BACK", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 4000).Value = Oracle.ManagedDataAccess.Types.OracleString.Null;
                        cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                        con.Open();
                        cmd.ExecuteNonQuery();

                        string error = cmd.Parameters["PERROR"].Value.ToString();
                        if (error == "OK")
                        {
                            return new Retorno();
                        }
                        else
                        {
                            return new Retorno
                            {
                                ret = "ERROR",
                                msg = "Ocurrio un error al ejecutar el rollback",
                                debug = error
                            };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrio una excepcion al ejecutar el rollback",
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
        }

        #endregion MONITOREO

        #region METODO SERVICIO

        public Retorno GetListadoCuentas()
        {

            Retorno ret = new Retorno();

            try
            {
                NavigatorLibreria Libreria = new NavigatorLibreria();
                Oracle.ManagedDataAccess.Client.OracleConnection con = new Oracle.ManagedDataAccess.Client.OracleConnection();
                con.ConnectionString = this.Conexion;
                con.Open();
                Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand();

                cmd.Connection = con;
                cmd.CommandText = this.Package + ".CARGAR_VALIDACIONES_CUENTAS";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 100;
                cmd.Parameters.Add("IO_CURSOR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter(cmd);

                DataSet ds = new DataSet();
                da.Fill(ds);

                string error = cmd.Parameters["PERROR"].Value != null ? cmd.Parameters["PERROR"].Value.ToString() : "";


                CultureInfo culturaActual = CultureInfo.CurrentCulture;
                string nombreCultura = culturaActual.Name;

                if (error.Equals("OK"))
                {
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        List<ValidacionCuentasTempServicio> Listado = ds.Tables[0].AsEnumerable().Select(item => new ValidacionCuentasTempServicio
                        {
                            ID = (int)item.Field<decimal?>("ID"),
                            DNI_RUT = item.Field<string>("DNI_RUT"),
                            NOMBRES_TRABAJADOR = item.Field<string>("NOMBRES_TRABAJADOR"),
                            APELLIDOS_TRABAJADOR = item.Field<string>("APELLIDOS_TRABAJADOR"),
                            EMAIL_PERSONAL = item.Field<string>("EMAIL_PERSONAL"),
                            EMAIL_JEFATURA = item.Field<string>("EMAIL_JEFATURA"),
                            MANAGER_RUT = item.Field<string>("MANAGER_RUT"),
                            MANAGER_NAME = item.Field<string>("MANAGER_NAME"),
                            MANAGER_EMAIL = item.Field<string>("MANAGER_EMAIL"),
                            INTERACCION = (int)item.Field<decimal?>("INTERACCION"),
                            FECHA_ULTIMA_INTERACCION = nombreCultura == "en-US" ? item.Field<string>("FECHA_ULTIMA_INTERACCION2") : item.Field<string>("FECHA_ULTIMA_INTERACCION"),
                            ENVIAR_MENSAJE_JEFE = item.Field<string>("ENVIAR_MENSAJE_JEFE"),
                            HAN_PASADO_DOS_DIAS = item.Field<string>("HAN_PASADO_DOS_DIAS"),
                            ID_ESTADO = (int)item.Field<decimal?>("ID_ESTADO"),
                            ESTADO = item.Field<string>("ESTADO"),
                            FECHA_RESPUESTA = item.Field<string>("FECHA_RESPUESTA"),
                            FECHA_ENVIO = item.Field<string>("FECHA_ENVIO"),
                            OBSERVACIONES = item.Field<string>("OBSERVACIONES")
                        }).ToList();

                        ret.values.Add(Listado);
                    }
                    else
                    {
                        ret.ret = "OK";
                        ret.msg = "No hay cuentas que validar";

                        List<ValidacionCuentasTempServicio> Listado = new List<ValidacionCuentasTempServicio>();
                        ret.values.Add(Listado);
                    }
                }
                else
                {
                    ret.ret = "ERROR";
                    ret.msg = "Ocurrio un error al consultar la configuracion de la aplicacion";
                    ret.debug = error;
                }

                if (con.State != ConnectionState.Closed)
                {
                    con.Close();
                }
                con.Dispose();
                cmd.Dispose();
                con = null;
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al obtener la configuracion del servicio";
                ret.debug = "Excepción al obtener la configuracion del servicio, Error: " + ex.Message;
            }


            return ret;
        }
        public Retorno GetListadoCuentasOld()
        {

            Retorno ret = new Retorno();

            try
            {
                using (Oracle.ManagedDataAccess.Client.OracleConnection con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    con.Open();
                    using (Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".CARGAR_VALIDACIONES_CUENTAS", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 300; // 5 minutes (300 seconds)
                        cmd.Parameters.Add("IO_CURSOR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 1024).Direction = ParameterDirection.Output;

                        using (Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter(cmd))
                        {
                            DataSet ds = new DataSet();
                            da.Fill(ds);

                            string error = cmd.Parameters["PERROR"].Value != null ? cmd.Parameters["PERROR"].Value.ToString() : "";

                            if (error.Equals("OK"))
                            {
                                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    List<ValidacionCuentasTempServicio> Listado = ds.Tables[0].AsEnumerable().Select(item => new ValidacionCuentasTempServicio
                                    {
                                        ID = (int)item.Field<decimal?>("ID"),
                                        DNI_RUT = item.Field<string>("DNI_RUT"),
                                        NOMBRES_TRABAJADOR = item.Field<string>("NOMBRES_TRABAJADOR"),
                                        APELLIDOS_TRABAJADOR = item.Field<string>("APELLIDOS_TRABAJADOR"),
                                        EMAIL_PERSONAL = item.Field<string>("EMAIL_PERSONAL"),
                                        EMAIL_JEFATURA = item.Field<string>("EMAIL_JEFATURA"),
                                        MANAGER_RUT = item.Field<string>("MANAGER_RUT"),
                                        MANAGER_NAME = item.Field<string>("MANAGER_NAME"),
                                        MANAGER_EMAIL = item.Field<string>("MANAGER_EMAIL"),
                                        INTERACCION = (int)item.Field<decimal?>("INTERACCION"),
                                        FECHA_ULTIMA_INTERACCION = item.Field<string>("FECHA_ULTIMA_INTERACCION"),
                                        ID_ESTADO = (int)item.Field<decimal?>("ID_ESTADO"),
                                        ESTADO = item.Field<string>("ESTADO"),
                                        FECHA_RESPUESTA = item.Field<string>("FECHA_RESPUESTA"),
                                        FECHA_ENVIO = item.Field<string>("FECHA_ENVIO"),
                                        OBSERVACIONES = item.Field<string>("OBSERVACIONES")
                                    }).ToList();

                                    ret.values.Add(Listado);
                                }
                                else
                                {
                                    ret.ret = "ERROR";
                                    ret.msg = "La Aplicacion no tiene configuracion en la base de datos";
                                }
                            }
                            else
                            {
                                ret.ret = "ERROR";
                                ret.msg = "Ocurrio un error al consultar la configuracion de la aplicacion";
                                ret.debug = error;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Ocurrio una excepcion al consultar la configuracion de la aplicacion";
                ret.debug = ex.Message + " - Detalles del error: " + ex.StackTrace;
            }

            return ret;
        }

        public Retorno ProcesarDatos()
        {
            try
            {
                using (var con = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
                {
                    using (var cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(this.Package + ".PROCESAR_DATOS", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("PERROR", OracleDbType.Varchar2, 4000).Value = Oracle.ManagedDataAccess.Types.OracleString.Null;
                        cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                        con.Open();
                        cmd.ExecuteNonQuery();

                        string error = cmd.Parameters["PERROR"].Value.ToString();
                        if (error == "OK")
                        {
                            return new Retorno();
                        }
                        else
                        {
                            return new Retorno
                            {
                                ret = "ERROR",
                                msg = "Ocurrio un error al crear el respaldo",
                                debug = error
                            };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrio una excepcion al crear el respaldo",
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
        }

        #endregion METODOS SERVICIOS

        #region GENERICOS

        public bool MultiInsertData(DataTable dt, string Tabla, ref string error)
        {
            error = "OK";
            using (var connection = new Oracle.ManagedDataAccess.Client.OracleConnection(this.Conexion))
            {
                try
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        using (var bulkCopy = new OracleBulkCopy(connection))
                        {
                            bulkCopy.BatchSize = 10000;
                            bulkCopy.DestinationTableName = Tabla;
                            bulkCopy.WriteToServer(dt);
                        }
                        transaction.Commit();
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }
        }

        public bool MultiInsertData(DataSet ds, string Columns, string tableName, ref string error)
        {
            bool ret = false;
            string connectionString = this.Conexion;
            Oracle.ManagedDataAccess.Client.OracleConnection connection = new Oracle.ManagedDataAccess.Client.OracleConnection();
            Oracle.ManagedDataAccess.Client.OracleDataAdapter myDataAdapter = new Oracle.ManagedDataAccess.Client.OracleDataAdapter();
            try
            {
                using (connection = new Oracle.ManagedDataAccess.Client.OracleConnection(connectionString))
                {
                    string SQLString = string.Format("select {0} from {1} where rownum=0", Columns, tableName);
                    using (Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(SQLString, connection))
                    {
                        try
                        {
                            connection.Open();
                            myDataAdapter = new Oracle.ManagedDataAccess.Client.OracleDataAdapter();
                            myDataAdapter.SelectCommand = new Oracle.ManagedDataAccess.Client.OracleCommand(SQLString, connection);
                            myDataAdapter.UpdateBatchSize = 0;
                            Oracle.ManagedDataAccess.Client.OracleCommandBuilder custCB = new Oracle.ManagedDataAccess.Client.OracleCommandBuilder(myDataAdapter);
                            DataTable dt = ds.Tables[0].Copy();
                            DataTable dtTemp = dt.Clone();

                            int times = 0;
                            for (int count = 0; count < dt.Rows.Count; times++)
                            {
                                for (int i = 0; i < 400 && 400 * times + i < dt.Rows.Count; i++, count++)
                                {
                                    dtTemp.Rows.Add(dt.Rows[count].ItemArray);
                                }
                                myDataAdapter.Update(dtTemp);
                                dtTemp.Rows.Clear();
                            }

                            dt.Dispose();
                            dtTemp.Dispose();
                            ret = true;
                        }
                        catch (Oracle.ManagedDataAccess.Client.OracleException ex)
                        {
                            error = ex.Message;
                            ret = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                ret = false;
            }
            finally
            {
                myDataAdapter.Dispose();
                connection.Close();
                connection.Dispose();
            }

            return ret;
        }

        /// <summary>
        /// Inserta Datos en Lote a una tabla especifica en la base de datos
        /// </summary>
        /// <param name="dt">Es un datatable  que contiene los datos a insertar</param>
        /// <param name="Tabla">Es un string que contiene el nombre de la tabla donde se insertan los datos</param>
        /// <param name="error">Es un string que contiene el mensaje de error en caso de haberlo en el proceso</param>
        /// <returns> true en caso exito, false en caso de error</returns> 
        //[Obsolete]
        public bool MultiInsertDataNuevo(DataTable dt, string Tabla, ref string error)
        {

            string connectionString = this.Conexion;
            error = "OK";
            using (Oracle.ManagedDataAccess.Client.OracleConnection connection = new Oracle.ManagedDataAccess.Client.OracleConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OracleBulkCopy bulkCopy = new OracleBulkCopy(connection))
                    {
                        bulkCopy.DestinationTableName = Tabla;
                        bulkCopy.WriteToServer(dt);
                    }

                    dt.Dispose();

                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message + " - Detalles del error: " + ex.StackTrace;
                    connection.Close();
                    connection.Dispose();
                    return false;
                }
                finally
                {
                    connection.Close();
                    connection.Dispose();
                }
            }
        }
        #endregion GENERICOS
    }
}
