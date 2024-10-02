using System;
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
using System.Globalization;
using Servicio.Notificaciones;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using MailKit.Security;

namespace Servicio.ReadMail
{
    public class ReaderImap
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

            Retorno ret2 = new Retorno();
            log.Write("Inicio proceso lectura de casillas de correo imap configuradas");

            //string baseDirectory = "d:\\apis";
            ////string baseDirectory = "e:\\apis";

            using (var client = new ImapClient())
            {
                try
                {
                    if (this.port.ToString() == "")
                    {
                        this.port = "993";
                    }
                    client.Connect(this.hostname, int.Parse(this.port), SecureSocketOptions.Auto);
                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                    client.AuthenticationMechanisms.Remove("NTLM");
                    client.Authenticate(this.username, this.password);
                    log.Write("Se autentica con las credenciales del correo");

                    client.Inbox.Open(FolderAccess.ReadWrite);

                    var query = SearchQuery.NotSeen;
                    var uids = client.Inbox.Search(query);
                    var items = client.Inbox.Fetch(uids, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                
                    foreach (var item in items)
                    {
                        var message = client.Inbox.GetMessage(item.UniqueId);

                        StringBuilder builder = new StringBuilder();

                        if (message.TextBody != null)
                        {
                            builder.Append(message.TextBody);
                        }
                        else
                        {
                            if (message.HtmlBody != null)
                            {
                                builder.Append(message.HtmlBody);
                            }
                        }
                        int attachments = item.Attachments.Count();

                        ret2 = Grabar(item.UniqueId.Id.ToString(), message, builder, attachments);
                        if (ret2.ret.Equals("OK"))
                        {
                            string idemail = (string)ret2.values[0];
                            string clasificacion = (string)ret2.values[1];
                            string ticket = (string)ret2.values[2];

                            if (clasificacion.Equals("OK"))
                            {
                                if (config.IdEmpresa == "2")
                                {
                                    ret = GrabarCordillera(ticket, message);
                                    if (!ret.ret.Equals("OK"))
                                    {
                                        log.Write("Excepcion al grabar correo en Cordillera: " + ret.debug);
                                        client.Inbox.AddFlags(item.UniqueId, MessageFlags.Seen, true);
                                    }
                                }
                                //var directory = Path.Combine(baseDirectory, item.UniqueId.ToString());

                                //if (item.Attachments.Count() != 0)
                                //    Directory.CreateDirectory(directory);

                                foreach (var attachment in item.Attachments)
                                {
                                    var entity = client.Inbox.GetBodyPart(item.UniqueId, attachment);

                                    var fileName = "";
                                    byte[] bytes;

                                    using (var memory = new MemoryStream())
                                    {
                                        if (entity is MessagePart)
                                        {
                                            var rfc822 = (MessagePart)entity;

                                            fileName = attachment.PartSpecifier + ".eml";
                                            rfc822.Message.WriteTo(memory);
                                        }
                                        else
                                        {
                                            var part = (MimePart)entity;

                                            fileName = part.FileName;
                                            part.Content.DecodeTo(memory);

                                        }
                                        bytes = memory.ToArray();
                                    }
                                    ret2 = Adjuntos(bytes, idemail, log, fileName, message.From.ToString());

                                    if (ret2.ret.Equals("OK"))
                                    {
                                        log.Write("Archivo adjunto grabado con éxito del email : " + message.From);
                                    }
                                }

                                Notificador notificador = new Notificador();

                                log.Write("Enviando notificación ticket : " + ticket);

                                notificador.Notificar(ticket);

                                log.Write("Fin envio notificación ticket : " + ticket);
                            }
                        }
                        else
                        {

                            log.Write("Excepcion al crear ticket CRM: " + ret2.debug);
                        }
                        client.Inbox.AddFlags(item.UniqueId, MessageFlags.Seen, true);
                    }
                    client.Disconnect(true);
                }
                catch (Exception ex)
                {
                    log.Write("Excepcion al leer correos imap: " + ex.Message);
                    client.Disconnect(true);
                }
            }

            return ret;
        }

        private Retorno Grabar(string id, MimeMessage msg, StringBuilder body, int attachments)
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
                cmd.Parameters.Add("PMESSAGE_ID", OracleType.VarChar, 255).Value = msg.MessageId.ToString();
                cmd.Parameters.Add("PFROM_ADDRESS", OracleType.VarChar, 255).Value = msg.From.ToString();
                cmd.Parameters.Add("PDISPLAY_NAME", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PHOST", OracleType.VarChar, 255).Value = "";
                cmd.Parameters.Add("PUSERNAME", OracleType.VarChar, 255).Value = msg.From.ToString();

                int indexOfAt = msg.From.ToString().IndexOf('@');
                string dominio = msg.From.ToString().Substring(indexOfAt + 1);
                dominio = dominio.Substring(dominio.IndexOf('.')).ToUpper();

                string datesend = String.Format("{0:dd/MM/yyyy HH:mm:ss}", Convert.ToDateTime(msg.Date.ToString()));
                string date = String.Format("{0:dd/MM/yyyy HH:mm:ss}", Convert.ToDateTime(msg.Date.ToString()));

                cmd.Parameters.Add("PEMAIL_DATE_SEND", OracleType.VarChar, 20).Value = datesend;
                cmd.Parameters.Add("PEMAIL_DATE", OracleType.VarChar, 20).Value = date;

                cmd.Parameters.Add("PIMPORTANCE", OracleType.VarChar, 255).Value = msg.Importance.ToString();

                cmd.Parameters.Add("PISVALID_MESSAGE", OracleType.Number, 10).Value = 0;

                cmd.Parameters.Add("PMULTIPART", OracleType.Number, 10).Value = 0;

                if (!String.IsNullOrEmpty(msg.Subject))
                {
                    cmd.Parameters.Add("PSUBJECT", OracleType.VarChar, 255).Value = msg.Subject.ToString();
                }
                else
                {
                    cmd.Parameters.Add("PSUBJECT", OracleType.VarChar, 255).Value = String.Empty;
                }

                cmd.Parameters.Add("PBODY", OracleType.Clob).Value = body.ToString();

                int k = attachments;

                cmd.Parameters.Add("PHAS_ATTACHMENT", OracleType.Number, 10).Value = k.ToString();

                cmd.Parameters.Add("PDOMINIO", OracleType.VarChar, 255).Value = dominio;

                if (!String.IsNullOrEmpty(msg.Subject.ToString()))
                {
                    if (msg.Subject.ToString().IndexOf("[Ticket#") != -1)
                    {
                        string header = msg.Subject.ToString();

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
                    ret.msg = "Excepción al grabar email id : " + id + " from : " + msg.From.ToString() + " error : " + ex.StackTrace + " " + ex.Message;
                    ret.debug = ex.Message;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al grabar email id : " + id + " from : " + msg.From.ToString() + " error : " + ex.StackTrace + " " + ex.Message;
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


        private Retorno GrabarCordillera(string ticket, MimeMessage msg)
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

                string date = String.Format("{0:dd/MM/yyyy HH:mm:ss}", Convert.ToDateTime(msg.Date.ToString()));

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
                    ret.msg = "Excepción al grabar ticket cordillera id : " + ticket + " from : " + msg.From.ToString() + " error : " + ex.StackTrace + " " + ex.Message;
                    ret.debug = ex.Message;
                }
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al grabar ticket cordillera id : " + ticket + " from : " + msg.From.ToString() + " error : " + ex.StackTrace + " " + ex.Message;
                ret.debug = ex.Message;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

            return ret;

        }

        /*
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
              else {
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

        */


        private Retorno Adjuntos(byte[] attachment, string idemail, Log log, string fileName, string from)
        {

            Retorno ret = new Retorno();

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();

            try
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
                tempLob.Write(attachment, 0, attachment.Length);
                tempLob.EndBatch();

                cmd.Parameters.Clear();
                cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.GRABA_ARCHIVO";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("PID_EMPRESA", OracleType.Number, 10).Value = config.IdEmpresa;
                cmd.Parameters.Add("PEMPRESA", OracleType.VarChar, 100).Value = config.Empresa;
                cmd.Parameters.Add("PID_SERVICIO", OracleType.Number, 10).Value = config.IdServicio;
                cmd.Parameters.Add("PSERVICIO", OracleType.VarChar, 100).Value = config.Servicio;

                cmd.Parameters.Add("PID_EMAIL", OracleType.Number, 10).Value = idemail;
                cmd.Parameters.Add("PNOMBRE", OracleType.VarChar, 255).Value = fileName;
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
                    log.Write("Excepcion al grabar archivo adjunto : " + fileName + " from : " + from + " error : " + ex.StackTrace + " " + ex.Message);
                }
                finally
                {
                    tx.Commit();
                    con.Close();
                    con.Dispose();
                }
                
            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al grabar adjunto from : " + from + " error : " + ex.StackTrace + " " + ex.Message;
                ret.debug = ex.Message;

                log.Write("Excepción al grabar adjunto from : " + from + " error : " + ex.StackTrace + " " + ex.Message);
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

        public byte[] FileToByteArray(string fileName)
        {
            return File.ReadAllBytes(fileName);
        }

        public ReaderImap()
        {

        }

        public ReaderImap(Configuracion config, Log log, string conexion, string conexionCordillera)
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