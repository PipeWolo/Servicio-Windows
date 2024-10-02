﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Web;
using System.Data.OracleClient;
using System.Web.Script.Serialization;
using Servicio.Clases;
using OpenPop.Common.Logging;
using OpenPop.Mime;
using OpenPop.Mime.Decode;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using System.Globalization;
using OpenPop.Mime.Traverse;
using Message = OpenPop.Mime.Message;
using Servicio.Notificaciones;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;

namespace Servicio.ReadMail
{
    public class ReaderPOP3
    {

        private string hostname { get; set; }
        private string port { get; set; }
        private bool useSsl { get; set; }
        private string username { get; set; }
        private string password { get; set; }
        private Log log { get; set; }
        private string conexion { get; set; }
        private string conexionCordillera { get; set; }
        private Configuracion config { get; set; }

        public Retorno Leer(string idservicio)
        {

            Retorno ret = new Retorno();

            log.Write("Inicio proceso lectura de casillas de correo pop3 configuradas");

            Retorno ret2 = Pop3Ids(idservicio);

            if (ret2.ret.Equals("OK"))
            {
                List<string> ids = (List<string>)ret2.values[0];

                ret2 = FetchAllMessages(ids, log);

                if (ret2.ret.Equals("OK"))
                {
                    Dictionary<Message, string> messages = (Dictionary<Message, string>)ret2.values[0];

                    foreach (var kvp in messages)
                    {
                        try
                        {
                            Message msg = (Message)kvp.Key;

                            string id = (string)kvp.Value;

                            OpenPop.Mime.MessagePart plainText = msg.FindFirstPlainTextVersion();

                            StringBuilder builder = new StringBuilder();

                            if (plainText != null)
                            {
                                builder.Append(plainText.GetBodyAsText());
                            }
                            else
                            {
                                OpenPop.Mime.MessagePart html = msg.FindFirstHtmlVersion();
                                if (html != null)
                                {
                                    builder.Append(html.GetBodyAsText());
                                }
                            }

                            ret2 = Grabar(id, msg, builder);

                            if (ret2.ret.Equals("OK"))
                            {
                                string idemail = (string)ret2.values[0];
                                string clasificacion = (string)ret2.values[1];
                                string ticket = (string)ret2.values[2];

                                if (clasificacion.Equals("OK"))
                                {
                                    if (config.IdEmpresa == "2")
                                    {
                                        ret = GrabarCordillera(ticket, msg);
                                        if (!ret.ret.Equals("OK"))
                                        {
                                            log.Write("Excepcion al grabar correo en Cordillera: " + ret.debug);
                                        }
                                    }
                                    if (msg.FindAllAttachments().Count() != 0)
                                    {
                                        ret2 = Adjuntos(msg, idemail, log);

                                        if (ret2.ret.Equals("OK"))
                                        {
                                            log.Write("Archivos adjuntos grabados con éxito del email : " + msg.Headers.From.Address);
                                        }
                                    }

                                    Notificador notificador = new Notificador();

                                    log.Write("Enviando notificación ticket : " + ticket);

                                    notificador.Notificar(ticket);

                                    log.Write("Fin envio notificación ticket : " + ticket);
                                }

                                log.Write("Eliminado mensaje : " + msg.Headers.From.Address + " id : " + msg.Headers.MessageId);

                                ret2 = Eliminar(msg.Headers.MessageId);

                                if (ret2.ret.Equals("OK"))
                                {
                                    log.Write("Mensaje : " + msg.Headers.From.Address + " id : " + msg.Headers.MessageId + " eliminado con éxito");
                                }
                                else
                                {
                                    log.Write("Falló al eliminar mensaje : " + msg.Headers.From.Address + " id : " + msg.Headers.MessageId + " error : " + ret2.msg + " excepcion : " + ret2.debug);
                                }
                            }
                            else
                            {
                                log.Write("Excepción al grabar email id : " + id + " from : " + msg.Headers.From.Address + " error : " + ret.msg + " debug : " + ret.debug);
                            }
                        }
                        catch (Exception ex)
                        {
                            log.Write("Excepción al procesar mensaje. Error : " + ex.Message + " debug : " + ex.StackTrace);
                        }
                    }
                }
                else
                {
                    log.Write("Excepción al procesar mensaje. Error : " + ret.msg + " debug : " + ret.debug);
                }
            }
            else
            {
                log.Write("Excepción al procesar mensaje. Error : " + ret.msg + " debug : " + ret.debug);
            }

            return ret;
        }

        private Retorno Eliminar(string messageId)
        {

            Retorno ret = new Retorno();

            var client = new Pop3Client();

            client.Connect(this.hostname, Convert.ToInt32(this.port), this.useSsl);

            client.Authenticate(username, password);

            try
            {

                bool found = false;

                int messageCount = client.GetMessageCount();

                for (int messageItem = messageCount; messageItem > 0; messageItem--)
                {
                    if (client.GetMessageHeaders(messageItem).MessageId == messageId)
                    {
                        client.DeleteMessage(messageItem);
                        client.Disconnect();
                        found = true;
                        break;
                    }
                }

                if (found)
                {
                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty;
                }
                else
                {
                    ret.ret = "ERROR";
                    ret.msg = "No se encontro el id de mensaje a eliminar : " + messageId;
                    ret.debug = String.Empty;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Error al eliminar id de mensaje : " + messageId + ". Exception : " + ex.Message;
                ret.debug = String.Empty;
            }

            return ret;

        }

        private Retorno FetchAllMessages(List<string> dbids, Log log)
        {

            Retorno ret = new Retorno();

            Dictionary<Message, string> allMessages = new Dictionary<Message, string>();

            try
            {
                var client = new Pop3Client();

                client.Connect(this.hostname, Convert.ToInt32(this.port), this.useSsl);

                client.Authenticate(this.username, this.password);

                List<string> msgids = client.GetMessageUids();

                for (int i = msgids.Count; i > 0; i--)
                {
                    try
                    {

                        string id = client.GetMessageUid(i);

                        int count = (from num in dbids where num.Equals(id) select num).Count();

                        if (count == 0)
                        {
                            Message msg = client.GetMessage(i);

                            allMessages.Add(msg, id);

                        }
                    }
                    catch (Exception ex)
                    {
                        log.Write("Error al procesar lectura de correos para cuenta : host " + this.hostname + " puerto : " + this.port + " usuario : " + this.username + " password : " + this.password + " error :" + ex.Message);
                    }
                }

                client.Disconnect();
                client.Dispose();

                ret.ret = "OK";
                ret.msg = String.Empty;
                ret.debug = String.Empty;

            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Error al procesar lectura de correos para cuenta : ";
                ret.debug = ex.Message;
            }

            ret.values = new List<object>();
            ret.values.Add(allMessages);

            return ret;

        }

        private Retorno Pop3Ids(string id)
        {

            Retorno ret = new Retorno();

            List<string> ids = new List<string>();

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();

            try
            {
                con.ConnectionString = this.conexion;  //"user id=ECCMAIL;data source=ENTELCC_ISADESA;password=DES#ECCMAIL123";
                con.Open();

                cmd.Connection = con;
                cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.CARGA_POP3_IDS";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("PID_SERVICIO", OracleType.Number, 10).Value = id;
                cmd.Parameters.Add("IO_CURSOR", OracleType.Cursor).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                OracleDataAdapter da = new OracleDataAdapter(cmd);

                DataSet ds = new DataSet();
                da.Fill(ds);

                if (ds != null)
                {

                    ids = ds.Tables[0].AsEnumerable().Select(r => Convert.ToString(r.Field<decimal>("pop3_id"))).ToList();

                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty; ;

                }
                else
                {
                    ret.ret = "ERROR";
                    ret.msg = "Falló al obtener último id de correo";
                    ret.debug = cmd.Parameters["PERROR"].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al obtener último id de correo";
                ret.debug = ex.Message + " " + ex.StackTrace;
            }
            finally
            {
                con.Close();
                cmd.Dispose();
            }

            ret.values = new List<object>();
            ret.values.Add(ids);

            return ret;
        }

        private Retorno Grabar(string id, Message msg, StringBuilder body)
        {

            Retorno ret = new Retorno();

            string ticket = "0";
            string idemail = String.Empty;
            string clasificacion = String.Empty;

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();

            try
            {

                con.ConnectionString = this.conexion;
                con.Open();

                cmd.Connection = con;
                cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.GRABA_EMAIL";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("PID_EMPRESA", OracleType.Number, 10).Value = config.IdEmpresa;
                cmd.Parameters.Add("PEMPRESA", OracleType.VarChar, 100).Value = config.Empresa;
                cmd.Parameters.Add("PID_SERVICIO", OracleType.Number, 10).Value = config.IdServicio;
                cmd.Parameters.Add("PSERVICIO", OracleType.VarChar, 100).Value = config.Servicio;

                cmd.Parameters.Add("PPOP3_ID", OracleType.Number, 10).Value = id;
                cmd.Parameters.Add("PMESSAGE_ID", OracleType.VarChar, 255).Value = msg.Headers.MessageId;
                cmd.Parameters.Add("PFROM_ADDRESS", OracleType.VarChar, 255).Value = msg.Headers.From.Address;
                cmd.Parameters.Add("PDISPLAY_NAME", OracleType.VarChar, 255).Value = msg.Headers.From.DisplayName;
                cmd.Parameters.Add("PHOST", OracleType.VarChar, 255).Value = msg.Headers.From.MailAddress.Host;
                cmd.Parameters.Add("PUSERNAME", OracleType.VarChar, 255).Value = msg.Headers.From.MailAddress.User;

                int indexOfAt = msg.Headers.From.Address.IndexOf('@');
                string dominio = msg.Headers.From.Address.Substring(indexOfAt + 1);
                dominio = dominio.Substring(dominio.IndexOf('.')).ToUpper();

                string datesend = String.Format("{0:dd/MM/yyyy HH:mm:ss}", Convert.ToDateTime(msg.Headers.DateSent));
                string date = String.Format("{0:dd/MM/yyyy HH:mm:ss}", Convert.ToDateTime(msg.Headers.Date));

                cmd.Parameters.Add("PEMAIL_DATE_SEND", OracleType.VarChar, 20).Value = datesend;
                cmd.Parameters.Add("PEMAIL_DATE", OracleType.VarChar, 20).Value = date;

                cmd.Parameters.Add("PIMPORTANCE", OracleType.VarChar, 255).Value = msg.Headers.Importance;

                if (msg.Headers.From.HasValidMailAddress)
                {
                    cmd.Parameters.Add("PISVALID_MESSAGE", OracleType.Number, 10).Value = 1;
                }
                else
                {
                    cmd.Parameters.Add("PISVALID_MESSAGE", OracleType.Number, 10).Value = 0;
                }

                if (msg.MessagePart.IsMultiPart)
                {
                    cmd.Parameters.Add("PMULTIPART", OracleType.Number, 10).Value = 1;
                }
                else
                {
                    cmd.Parameters.Add("PMULTIPART", OracleType.Number, 10).Value = 0;
                }

                if (!String.IsNullOrEmpty(msg.Headers.Subject))
                {
                    cmd.Parameters.Add("PSUBJECT", OracleType.VarChar, 255).Value = msg.Headers.Subject;
                }
                else
                {
                    cmd.Parameters.Add("PSUBJECT", OracleType.VarChar, 255).Value = String.Empty;
                }

                cmd.Parameters.Add("PBODY", OracleType.Clob).Value = body.ToString();

                int k = msg.FindAllAttachments().Count();

                cmd.Parameters.Add("PHAS_ATTACHMENT", OracleType.Number, 10).Value = k.ToString();

                cmd.Parameters.Add("PDOMINIO", OracleType.VarChar, 255).Value = dominio;

                if (!String.IsNullOrEmpty(msg.Headers.Subject))
                {
                    if (msg.Headers.Subject.IndexOf("[Ticket#") != -1)
                    {
                        string header = msg.Headers.Subject;

                        string asunto = header.Substring(header.IndexOf("[Ticket#"));

                        string[] data = asunto.Split('#');

                        ticket = data[1].Replace("]", "");

                        //string numero = asunto.Substring(1, asunto.IndexOf("]") - 1);
                        //string[] data = numero.Split('#');

                        //ticket = data[1];
                    }
                    else
                    {
                        ticket = "0";
                    }
                }

                if (!String.IsNullOrEmpty(ticket))
                {
                    cmd.Parameters.Add("PTICKET", OracleType.Number).Value = ticket;
                }
                else
                {
                    cmd.Parameters.Add("PTICKET", OracleType.Number).Value = "0";
                }
                cmd.Parameters["PTICKET"].Direction = ParameterDirection.InputOutput;

                cmd.Parameters.Add("PID_EMAIL", OracleType.Number, 1024).Value = 0;
                cmd.Parameters["PID_EMAIL"].Direction = ParameterDirection.Output;

                cmd.Parameters.Add("PCLASIFICACION", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters["PCLASIFICACION"].Direction = ParameterDirection.Output;

                cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                try
                {
                    cmd.ExecuteNonQuery();

                    idemail = cmd.Parameters["PID_EMAIL"].Value.ToString();
                    clasificacion = cmd.Parameters["PCLASIFICACION"].Value.ToString();
                    ticket = cmd.Parameters["PTICKET"].Value.ToString();

                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty;
                }
                catch (Exception ex)
                {
                    ret.ret = "ERROR";
                    ret.msg = "Excepción al grabar email id : " + id + " from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message;
                    ret.debug = ex.Message;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al grabar email id : " + id + " from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message;
                ret.debug = ex.Message;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

            ret.values = new List<object>();
            ret.values.Add(idemail);
            ret.values.Add(clasificacion);
            ret.values.Add(ticket);

            return ret;

        }

        private Retorno GrabarCordillera(string ticket, Message msg)
        {

            Retorno ret = new Retorno();

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();

            try
            {

                con.ConnectionString = this.conexionCordillera;
                con.Open();

                cmd.Connection = con;
                cmd.CommandText = "PKG_CORDILLERA_CRMMAIL.CREAR_ATENCION_CORDILLERA";
                cmd.CommandType = CommandType.StoredProcedure;


                cmd.Parameters.Add("PID_TICKET", OracleType.Number).Value = ticket;

                string date = String.Format("{0:dd/MM/yyyy HH:mm:ss}", Convert.ToDateTime(msg.Headers.Date));

                cmd.Parameters.Add("PFECHA_INGRESO", OracleType.VarChar, 20).Value = date;
                cmd.Parameters.Add("PPATENTE", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PMARCA", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PMODELO", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PANO", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PCOLOR", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PEMPRESA", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PSKILL", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PRUT_AGENTE", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PAGENTE", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PRUT", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PDV", OracleType.VarChar, 255).Value = "";
                cmd.Parameters["PID_TICKET"].Direction = ParameterDirection.InputOutput;
                cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters.Add("IO_CURSOR", OracleType.Cursor).Direction = ParameterDirection.Output;
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                try
                {
                    OracleDataAdapter da = new OracleDataAdapter(cmd);

                    DataSet ds = new DataSet();

                    da.Fill(ds);

                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty;
                }
                catch (Exception ex)
                {
                    ret.ret = "ERROR";
                    ret.msg = "Excepción al grabar ticket cordillera id : " + ticket + " from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message;
                    ret.debug = ex.Message;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al grabar ticket cordillera id : " + ticket + " from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message;
                ret.debug = ex.Message;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

            return ret;

        }
        private Retorno Adjuntos(Message msg, string idemail, Log log)
        {

            Retorno ret = new Retorno();

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();

            try
            {
                foreach (var attachment in msg.FindAllAttachments())
                {

                    con.ConnectionString = this.conexion;
                    con.Open();

                    OracleTransaction tx = con.BeginTransaction();

                    cmd = con.CreateCommand();
                    cmd.Transaction = tx;
                    cmd.CommandText = "declare xx blob; begin dbms_lob.createtemporary(xx, false, 0); :tempblob := xx; end;";
                    cmd.Parameters.Add(new OracleParameter("tempblob", OracleType.Blob)).Direction = ParameterDirection.Output;
                    cmd.ExecuteNonQuery();

                    OracleLob tempLob = (OracleLob)cmd.Parameters[0].Value;
                    tempLob.BeginBatch(OracleLobOpenMode.ReadWrite);
                    tempLob.Write(attachment.Body, 0, attachment.Body.Length);
                    tempLob.EndBatch();

                    cmd.Parameters.Clear();
                    cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.GRABA_ARCHIVO";
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("PID_EMPRESA", OracleType.Number, 10).Value = config.IdEmpresa;
                    cmd.Parameters.Add("PEMPRESA", OracleType.VarChar, 100).Value = config.Empresa;
                    cmd.Parameters.Add("PID_SERVICIO", OracleType.Number, 10).Value = config.IdServicio;
                    cmd.Parameters.Add("PSERVICIO", OracleType.VarChar, 100).Value = config.Servicio;

                    cmd.Parameters.Add("PID_EMAIL", OracleType.Number, 10).Value = idemail;
                    cmd.Parameters.Add("PNOMBRE", OracleType.VarChar, 255).Value = attachment.FileName;
                    cmd.Parameters.Add(new OracleParameter("PARCHIVO", OracleType.Blob)).Value = tempLob;
                    cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                    cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                    try
                    {
                        cmd.ExecuteNonQuery();

                        ret.ret = "OK";
                        ret.msg = String.Empty;
                        ret.debug = String.Empty;
                    }
                    catch (Exception ex)
                    {
                        log.Write("Excepcion al grabar archivo adjunto : " + attachment.FileName + " from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message);
                    }
                    finally
                    {
                        tx.Commit();
                        con.Close();
                        con.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al grabar adjunto from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message;
                ret.debug = ex.Message;

                log.Write("Excepción al grabar adjunto from : " + msg.Headers.From.Address + " error : " + ex.StackTrace + " " + ex.Message);
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                    con.Dispose();
                }
            }

            return ret;

        }

        public ReaderPOP3()
        {

        }

        public ReaderPOP3(Configuracion config, Log log, string conexion, string conexionCordillera)
        {
            this.config = config;
            this.hostname = config.Dominio;
            this.port = config.Puerto;
            this.useSsl = (config.Ssl.Equals("1")) ? true : false;
            this.username = config.Usuario;
            this.password = config.Password;
            this.log = log;
            this.conexion = conexion;
            this.conexionCordillera = conexionCordillera;
        }
    }
}