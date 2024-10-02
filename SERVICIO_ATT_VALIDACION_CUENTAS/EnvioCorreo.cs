using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Web;
using System.Data.OracleClient;
using System.Web.Script.Serialization;
using Servicio.Clases;
using Servicio.ReadMail;

namespace Servicio.SendMail
{
    public class Sender
    {

        private Config config_serv = new Config();

        private Retorno Cargar(Configuracion config)
        {

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();
            Retorno ret = new Retorno();

            List<Correo> correos = new List<Correo>();

            try
            {

                con.ConnectionString = config_serv.ConMail;
                con.Open();

                cmd.Connection = con;
                cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.CARGA_CORREOS_ENVIAR";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("PID_EMPRESA", OracleType.Number, 10).Value = config.IdEmpresa;
                cmd.Parameters.Add("PID_SERVICIO", OracleType.Number, 10).Value = config.IdServicio;
                cmd.Parameters.Add("IO_CURSOR", OracleType.Cursor).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                OracleDataAdapter da = new OracleDataAdapter(cmd);

                DataSet ds = new DataSet();

                da.Fill(ds);

                if (ds != null)
                {

                    var cor = ds.Tables[0].AsEnumerable();

                    correos = (from item in cor
                               select new Correo
                               {
                                   Id = Convert.ToString(item.Field<decimal>("id")),
                                   IdEmpresa = Convert.ToString(item.Field<decimal>("id_empresa")),
                                   IdServicio = Convert.ToString(item.Field<decimal>("id_servicio")),
                                   FechaRegistro = Convert.ToString(item.Field<DateTime>("fecha_registro")),
                                   Estado = Convert.ToString(item.Field<decimal>("estado")),
                                   FechaEnvio = Convert.ToString(item.Field<DateTime?>("fecha_envio")),
                                   correo = item.Field<string>("correo"),
                                   IdTicket = Convert.ToString(item.Field<decimal?>("id_ticket")),
                                   Para = item.Field<string>("para"),
                                   CC = item.Field<string>("cc"),
                                   CCO = item.Field<string>("cco"),
                               }).ToList();

                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty; ;

                }
                else
                {
                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty; ;
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
            ret.values.Add(correos);

            return ret;
        }

        private Retorno Adjuntos(string id)
        {

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();
            Retorno ret = new Retorno();

            List<Adjunto> adjuntos = new List<Adjunto>();
            List<Cid> cids = new List<Cid>();

            try
            {

                con.ConnectionString = config_serv.ConMail;  //"user id=ECCMAIL;data source=ENTELCC_ISADESA;password=DES#ECCMAIL123";
                con.Open();

                cmd.Connection = con;
                cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.CARGA_ADJUNTOS_CORREO";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("PID", OracleType.Number, 10).Value = id;
                cmd.Parameters.Add("IO_CURSOR", OracleType.Cursor).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("IO_CURSOR2", OracleType.Cursor).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                OracleDataAdapter da = new OracleDataAdapter(cmd);

                DataSet ds = new DataSet();

                da.Fill(ds);

                if (ds != null)
                {

                    var adj = ds.Tables[0].AsEnumerable();
                    var cid = ds.Tables[1].AsEnumerable();

                    adjuntos = (from item in adj
                                select new Adjunto
                                {
                                    Id = Convert.ToString(item.Field<decimal>("id")),
                                    IdCorreo = Convert.ToString(item.Field<decimal>("id_correo")),
                                    NombreAdjunto = item.Field<string>("nombre_adjunto"),
                                    Extension = item.Field<string>("extension"),
                                    adjunto = item.Field<byte[]>("adjunto")
                                }).ToList();

                    cids = (from item in cid
                            select new Cid
                            {
                                Id = Convert.ToString(item.Field<decimal>("id")),
                                IdCorreo = Convert.ToString(item.Field<decimal>("id_correo")),
                                LLave = item.Field<string>("llave"),
                                NombreCid = item.Field<string>("nombre_cid"),
                                Extension = item.Field<string>("extension"),
                                cid = item.Field<byte[]>("cid")
                            }).ToList();

                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty; ;

                }
                else
                {
                    ret.ret = "OK";
                    ret.msg = String.Empty;
                    ret.debug = String.Empty; ;
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
            ret.values.Add(adjuntos);
            ret.values.Add(cids);

            return ret;
        }

        private Retorno Estado(string id, Retorno estado)
        {

            OracleConnection con = new OracleConnection();
            OracleCommand cmd = new OracleCommand();
            Retorno ret = new Retorno();

            try
            {

                con.ConnectionString = config_serv.ConMail; //"user id=ECCMAIL;data source=ENTELCC_ISADESA;password=DES#ECCMAIL123";
                con.Open();

                cmd.Connection = con;
                cmd.CommandText = "PKG_SERVICIO_SERVICE_MAIL.ACTUALIZA_ESTADO_ENVIO";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("PID", OracleType.Number, 10).Value = id;

                if (estado.ret.Equals("OK"))
                {
                    cmd.Parameters.Add("PESTADO", OracleType.Number).Value = 1;
                    cmd.Parameters.Add("PMENSAJE", OracleType.VarChar, 1024).Value = "OK";
                }
                else if (estado.ret.Equals("ERRORENVIO"))
                {
                    cmd.Parameters.Add("PESTADO", OracleType.Number).Value = 0;
                    cmd.Parameters.Add("PMENSAJE", OracleType.VarChar, 1024).Value = "Problema al conectarse con servidor de envio de correo.";
                }
                else
                {
                    cmd.Parameters.Add("PESTADO", OracleType.Number).Value = 2;
                    cmd.Parameters.Add("PMENSAJE", OracleType.VarChar, 1024).Value = estado.debug;
                }

                cmd.Parameters.Add("PERROR", OracleType.VarChar, 1024).Value = "";
                cmd.Parameters["PERROR"].Direction = ParameterDirection.Output;

                cmd.ExecuteNonQuery();

                ret.ret = "OK";
                ret.msg = String.Empty;
                ret.debug = String.Empty; ;

            }
            catch (Exception ex)
            {
                ret.ret = "ERROR";
                ret.msg = "Excepción al actualizar estado envio correo";
                ret.debug = ex.Message + " " + ex.StackTrace;
            }
            finally
            {
                con.Close();
                cmd.Dispose();
            }

            ret.values = new List<object>();

            return ret;
        }

        private List<string> Destinarios(string email)
        {

            List<string> ret = new List<string>();

            if (!String.IsNullOrEmpty(email))
            {
                if (email.IndexOf(';') != -1)
                {
                    string[] data = email.Split(';');

                    foreach (string em in data)
                    {
                        if (!String.IsNullOrEmpty(em))
                        {
                            ret.Add(em);
                        }
                    }
                }
                else
                {
                    ret.Add(email);
                }
            }
            else
            {
                ret.Add(String.Empty);
            }

            return ret;

        }

        private Retorno EnviarCCO(Log log, List<Configuracion> configuraciones)
        {

            Retorno ret = new Retorno();

            string path = config_serv.LogServicio;

            foreach (Configuracion config in configuraciones)
            {

                Retorno ret2 = Cargar(config);

                if (ret2.ret.Equals("OK"))
                {
                    List<Correo> correos = (List<Correo>)ret2.values[0];

                    foreach (Correo correo in correos)
                    {
                        if (!String.IsNullOrEmpty(correo.CCO))
                        {
                            Retorno ret3 = Adjuntos(correo.Id);

                            if (ret3.ret.Equals("OK"))
                            {

                                List<Adjunto> adjuntos = (List<Adjunto>)ret3.values[0];
                                List<Cid> cids = (List<Cid>)ret3.values[1];

                                try
                                {

                                    MailMessage mail = new MailMessage();
                                    SmtpClient SmtpServer = new SmtpClient(config.Dominio);

                                    SmtpServer.Port = Convert.ToInt32(config.PuertoStmp);
                                    SmtpServer.Credentials = new NetworkCredential(config.Usuario, config.Password);
                                    SmtpServer.EnableSsl = false;

                                    mail.From = new MailAddress(config.NoReply, config.AsuntoStmp);

                                    List<string> cco = this.Destinarios(correo.CCO);

                                    foreach (string em in cco)
                                    {
                                        if (!String.IsNullOrEmpty(em))
                                            mail.To.Add(em);
                                    }

                                    mail.Subject = "[Ticket#" + correo.IdTicket + "] " + config.AsuntoStmp;

                                    foreach (Adjunto adjunto in adjuntos)
                                    {
                                        string archivo = path + adjunto.NombreAdjunto;

                                        if (File.Exists(archivo))
                                            File.Delete(archivo);

                                        using (FileStream fstream = new FileStream(archivo, FileMode.Create, FileAccess.Write, FileShare.Read))
                                        {
                                            fstream.Write(adjunto.adjunto, 0, adjunto.adjunto.Length);
                                        }

                                        using (FileStream fs = new FileStream(archivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                        {
                                            MemoryStream memStream = new MemoryStream();
                                            memStream.SetLength(fs.Length);
                                            fs.Read(memStream.GetBuffer(), 0, (int)fs.Length);
                                            fs.Close();
                                            mail.Attachments.Add(new Attachment(memStream, new FileInfo(archivo).Name));
                                        }
                                    }

                                    if (cids.Count > 0)
                                    {

                                        AlternateView av1 = AlternateView.CreateAlternateViewFromString(correo.correo, Encoding.GetEncoding("iso8859-1"), MediaTypeNames.Text.Html);

                                        foreach (Cid cid in cids)
                                        {
                                            string archivo = path + cid.NombreCid;

                                            if (File.Exists(archivo))
                                                File.Delete(archivo);

                                            using (FileStream fstream = new FileStream(archivo, FileMode.Create, FileAccess.Write, FileShare.Read))
                                            {
                                                fstream.Write(cid.cid, 0, cid.cid.Length);
                                                fstream.Close();
                                                fstream.Dispose();
                                            }


                                            LinkedResource cidimg = new LinkedResource(archivo, MediaTypeNames.Image.Jpeg);

                                            string[] data = cid.LLave.Split(':');

                                            cidimg.ContentId = data[1];

                                            av1.LinkedResources.Add(cidimg);

                                        }

                                        mail.AlternateViews.Add(av1);

                                    }
                                    else
                                    {
                                        mail.Body = correo.correo;
                                    }

                                    mail.IsBodyHtml = true;

                                    try
                                    {
                                        SmtpServer.Send(mail);
                                        mail.Dispose();

                                        ret.ret = "OK";
                                        ret.msg = String.Empty;
                                        ret.debug = String.Empty;
                                    }
                                    catch (Exception ex)
                                    {
                                        ret.ret = "ERROR";
                                        ret.msg = "Falló al enviar email de notificación";
                                        ret.debug = ex.Message;
                                    }

                                    Retorno ret4 = Estado(correo.Id, ret);

                                    if (!ret4.ret.Equals("OK"))
                                    {
                                        log.Write("Excepción al actualizar estado envia email : " + config.Empresa + " servicio : " + config.Servicio + " correo : " + correo.Para + " error : " + ret4.msg);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Write("Excepción al enviar email : " + config.Empresa + " servicio : " + config.Servicio + " correo : " + correo.Para + " error : " + ex.Message);
                                }
                            }
                            else
                            {
                                log.Write("Excepción al cargar adjuntos email : " + config.Empresa + " servicio : " + config.Servicio + " correo : " + correo.Para + " error : " + ret.msg);
                            }
                        }
                    }
                }
                else
                {
                    log.Write("Excepción al procesar cliente : " + config.Empresa + " servicio : " + config.Servicio + " error : " + ret.msg);
                }
            }

            return ret;

        }

        public void Enviar()
        {

            Reader reader = new Reader();

            Retorno ret = reader.Servicios(config_serv.Pais);

            string path = config_serv.LogServicio;

            Log log = new Log(path);

            log.Write("Inicio proceso envio correos");

            if (ret.ret.Equals("OK"))
            {

                List<Configuracion> configuraciones = (List<Configuracion>)ret.values[0];

                foreach (Configuracion config in configuraciones)
                {

                    Retorno ret2 = Cargar(config);

                    if (ret2.ret.Equals("OK"))
                    {
                        List<Correo> correos = (List<Correo>)ret2.values[0];

                        foreach (Correo correo in correos)
                        {

                            Retorno ret3 = Adjuntos(correo.Id);

                            if (ret3.ret.Equals("OK"))
                            {

                                List<Adjunto> adjuntos = (List<Adjunto>)ret3.values[0];
                                List<Cid> cids = (List<Cid>)ret3.values[1];

                                try
                                {

                                    MailMessage mail = new MailMessage();
                                    SmtpClient SmtpServer = new SmtpClient(config.DominioSalida);

                                    SmtpServer.Port = Convert.ToInt32(config.PuertoStmp);
                                    if(config.Tipo == "IMAP")
                                    {
                                        SmtpServer.EnableSsl = true;
                                        SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                                        SmtpServer.UseDefaultCredentials = false;
                                    }
                                    SmtpServer.Credentials = new NetworkCredential(config.Usuario, config.Password);
                                    if(config.Tipo != "IMAP")
                                    {
                                        SmtpServer.EnableSsl = false;
                                    }
                                    if (correo.IdTicket != "0")
                                    {
                                        mail.From = new MailAddress(config.NoReply, config.AsuntoStmp);
                                    }
                                    else
                                    {
                                        mail.From = new MailAddress(config.NoReply, config.AsuntoStmpNuevosCorreos);
                                    }

                                    List<string> para = this.Destinarios(correo.Para);
                                    List<string> cc = this.Destinarios(correo.CC);
                                    List<string> cco = this.Destinarios(correo.CCO);

                                    foreach (string em in para)
                                    {
                                        if (!String.IsNullOrEmpty(em))
                                            mail.To.Add(em);
                                    }

                                    foreach (string em in cc)
                                    {
                                        if (!String.IsNullOrEmpty(em))
                                            mail.CC.Add(em);
                                    }

                                    bool bcc = false;

                                    foreach (string em in cco)
                                    {
                                        if (!String.IsNullOrEmpty(em))
                                            bcc = true;
                                    }

                                    if (correo.IdTicket != "0")
                                    {
                                        mail.Subject = config.AsuntoStmp + " [Ticket#" + correo.IdTicket + "]";
                                    }
                                    else
                                    {
                                        mail.Subject = config.AsuntoStmpNuevosCorreos;
                                    }
                                    
                                    foreach (Adjunto adjunto in adjuntos)
                                    {
                                        string archivo = path + adjunto.NombreAdjunto;

                                        if (File.Exists(archivo))
                                            File.Delete(archivo);

                                        using (FileStream fstream = new FileStream(archivo, FileMode.Create, FileAccess.Write, FileShare.Read))
                                        {
                                            fstream.Write(adjunto.adjunto, 0, adjunto.adjunto.Length);
                                        }

                                        using (FileStream fs = new FileStream(archivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                        {
                                            MemoryStream memStream = new MemoryStream();
                                            memStream.SetLength(fs.Length);
                                            fs.Read(memStream.GetBuffer(), 0, (int)fs.Length);
                                            fs.Close();
                                            mail.Attachments.Add(new Attachment(memStream, new FileInfo(archivo).Name));
                                        }
                                    }

                                    if (cids.Count > 0)
                                    {

                                        AlternateView av1 = AlternateView.CreateAlternateViewFromString(correo.correo, Encoding.GetEncoding("iso8859-1"), MediaTypeNames.Text.Html);

                                        foreach (Cid cid in cids)
                                        {
                                            string archivo = path + cid.NombreCid;

                                            if (File.Exists(archivo))
                                                File.Delete(archivo);

                                            using (FileStream fstream = new FileStream(archivo, FileMode.Create, FileAccess.Write, FileShare.Read))
                                            {
                                                fstream.Write(cid.cid, 0, cid.cid.Length);
                                                fstream.Close();
                                                fstream.Dispose();
                                            }


                                            LinkedResource cidimg = new LinkedResource(archivo, MediaTypeNames.Image.Jpeg);

                                            string[] data = cid.LLave.Split(':');

                                            cidimg.ContentId = data[1];

                                            av1.LinkedResources.Add(cidimg);

                                        }

                                        mail.AlternateViews.Add(av1);

                                    }
                                    else
                                    {
                                        mail.Body = correo.correo;
                                    }

                                    mail.IsBodyHtml = true;

                                    try
                                    {
                                        SmtpServer.Send(mail);
                                        mail.Dispose();

                                        ret.ret = "OK";
                                        ret.msg = String.Empty;
                                        ret.debug = String.Empty;
                                    }
                                    catch(SmtpException ex)
                                    {
                                        log.Write("Status Code: " + ex.StatusCode.ToString());
                                        ret.ret = "ERRORENVIO";
                                        ret.msg = "Falló al enviar email de notificación " + ex.StatusCode;
                                        ret.debug = ex.Message;
                                    }
                                    catch (Exception ex)
                                    {
                                        ret.ret = "ERROR";
                                        ret.msg = "Falló al enviar email de notificación";
                                        ret.debug = ex.Message;
                                    }

                                    //correos copia oculta
                                    if (bcc)
                                    {
                                        ret = this.EnviarCCO(log, configuraciones);
                                    }
                                    else
                                    {
                                        Retorno ret4 = Estado(correo.Id, ret);

                                        if (!ret4.ret.Equals("OK"))
                                        {
                                            log.Write("Excepción al actualizar estado envia email : " + config.Empresa + " servicio : " + config.Servicio + " correo : " + correo.Para + " error : " + ret4.msg);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Write("Excepción al enviar email : " + config.Empresa + " servicio : " + config.Servicio + " correo : " + correo.Para + " error : " + ex.Message);
                                }
                            }
                            else
                            {
                                log.Write("Excepción al cargar adjuntos email : " + config.Empresa + " servicio : " + config.Servicio + " correo : " + correo.Para + " error : " + ret.msg);
                            }
                        }
                    }
                    else
                    {
                        log.Write("Excepción al procesar cliente : " + config.Empresa + " servicio : " + config.Servicio + " error : " + ret.msg);
                    }
                }
            }
            else
            {
                log.Write("Error al cargar servicios a procesar : error : " + ret.msg);
                log.Write("Error debug : error : " + ret.debug);
            }

            log.Write("Fin proceso envio correos");
        }
    }
}