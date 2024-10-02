using System;
using System.Collections.Generic;
using Limilabs.Mail;
using Limilabs.Client.IMAP;
using Limilabs.Client.SMTP;
using System.Net.Mail;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Diagnostics;
using System.Security.Authentication;

namespace ServiceMail
{
    public class MailImap
    {
        private string Servidor { get; set; }
        private string Usuario { get; set; }
        private string Password { get; set; }
        private int PuertoIMAP { get; set; }
        private int PuertoSMTP { get; set; }
        public MailImap(string Servidor, string Usuario, string Password, string PuertoIMAP, string puertoSMTP)
        {
            this.Servidor = Servidor;
            this.Usuario = Usuario;
            this.Password = Password;
            this.PuertoIMAP = int.Parse(PuertoIMAP);
            this.PuertoSMTP = int.Parse(puertoSMTP);
        }

        public Retorno LeerCasillaCorreos()
        {
            try
            {
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                

                using (Imap imap = new Imap())
                {
                    // Conectar al servidor IMAP
                    imap.SSLConfiguration.EnabledSslProtocols = SslProtocols.Tls12;
                    imap.ConnectSSL(this.Servidor, this.PuertoIMAP);
                    imap.UseBestLogin(this.Usuario, this.Password);

                    // Seleccionar la carpeta Inbox o carpeta de entrada
                    imap.SelectInbox();

                    // Recupero los correos electronicos que no han sido leidos
                    var uids = imap.Search(Flag.Unseen);

                    // Creo una lista de los correos no leeidos, solo los que estan en la bandeja de entrada
                    List<ListadoEmailMensajes> emails = new List<ListadoEmailMensajes>();
                    foreach (long uid in uids)
                    {
                        IMail email = new MailBuilder()
                            .CreateFromEml(imap.GetMessageByUID(uid));

                        ListadoEmailMensajes message = new ListadoEmailMensajes
                        {
                            Uid = uid,
                            From = email.From.ToString(),
                            To = email.To.ToString(),
                            Cc = email.Cc.ToString(),
                            Subject = email.Subject,
                            Text = email.Text,
                            FechaRecepcion = email.Date?.ToString("yyyy-MM-dd HH:mm:ss")
                        };

                        emails.Add(message);
                    }

                    // Cierro la conexión al servidor IMAP
                    imap.Close();

                    //Retorno la lista de correos
                    Retorno ret = new Retorno();
                    ret.values.Add(emails);
                    return ret;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    using (Imap imap = new Imap())
                    {
                        // Conectar al servidor IMAP
                        imap.Connect(this.Servidor, this.PuertoIMAP);
                        imap.UseBestLogin(this.Usuario, this.Password);

                        // Seleccionar la carpeta Inbox o carpeta de entrada
                        imap.SelectInbox();

                        // Recupero los correos electronicos que no han sido leidos
                        var uids = imap.Search(Flag.Unseen);

                        // Creo una lista de los correos no leeidos, solo los que estan en la bandeja de entrada
                        List<ListadoEmailMensajes> emails = new List<ListadoEmailMensajes>();
                        foreach (long uid in uids)
                        {
                            IMail email = new MailBuilder()
                                .CreateFromEml(imap.GetMessageByUID(uid));

                            ListadoEmailMensajes message = new ListadoEmailMensajes
                            {
                                Uid = uid,
                                From = email.From.ToString(),
                                To = email.To.ToString(),
                                Cc = email.Cc.ToString(),
                                Subject = email.Subject,
                                Text = email.Text,
                                FechaRecepcion = email.Date?.ToString("yyyy-MM-dd HH:mm:ss")
                            };

                            emails.Add(message);
                        }

                        // Cierro la conexión al servidor IMAP
                        imap.Close();

                        //Retorno la lista de correos
                        Retorno ret = new Retorno();
                        ret.values.Add(emails);
                        return ret;
                    }
                }
                catch (Exception)
                {

                    return new Retorno
                    {
                        ret = "ERROR",
                        msg = "Ocurrio una excepcion al crear el respaldo",
                        debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                    };
                }
            }
        }

        public Retorno EnviarEmail(string Destinatario, string Asunto, string Mensaje)
        {
            try
            {
                using (SmtpClient smtp = new SmtpClient(this.Servidor, this.PuertoSMTP))
                {
                    using (MailMessage message = new MailMessage())
                    {
                        message.From = new MailAddress(this.Usuario);
                        message.To.Add(new MailAddress(Destinatario));
                        message.Subject = Asunto;
                        message.Body = Mensaje;

                        smtp.EnableSsl = true;
                        smtp.Credentials = new NetworkCredential(this.Usuario, this.Password);

                        // Envía el mensaje
                        smtp.Send(message);

                        return new Retorno();
                    }
                }
            }

            catch (FormatException ex)
            {
                return new Retorno
                {
                    ret = "ERROR_EMAIL",
                    msg = "El destinatario " + Destinatario + ", no es un correo electrónico válido",
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
            catch (SmtpException ex)
            {
                return new Retorno
                {
                    ret = "ERROR_ENVIO",
                    msg = "Error al enviar correo electrónico a " + Destinatario,
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrio una excepcion al enviar el mensaje a " + Destinatario,
                    debug = ex.Message + " - Detalles del error: " + ex.StackTrace
                };
            }
        }

        public Retorno EnviarMensajeConPlantillaLogCID(string destinatario, List<string> otrosDestinatarios, List<string> otrosDestinatariosOcultos, string Asunto, string Mensaje, string Monitoreo, ref string error)
        {
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

                SmtpClient smtpClient = new SmtpClient(this.Servidor, this.PuertoSMTP);
                smtpClient.Credentials = new NetworkCredential(this.Usuario, this.Password);
                smtpClient.EnableSsl = true;

                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(this.Usuario);
                mailMessage.To.Add(destinatario);

                // Agregar copias, es decir los otros destinatarios
                if (otrosDestinatarios != null)
                {
                    foreach (var otroDestinatario in otrosDestinatarios)
                    {
                        mailMessage.CC.Add(otroDestinatario.Trim());
                    }
                }

                // Agregar copias ocultas, es decir los destinatarios ocultos
                if (otrosDestinatariosOcultos != null)
                {
                    foreach (var destinatarioOculto in otrosDestinatariosOcultos)
                    {
                        mailMessage.Bcc.Add(destinatarioOculto.Trim());
                    }
                }

                mailMessage.Subject = Asunto;

                // Leer la plantilla HTML desde la ruta especificada
                string rutaPlantillaHtml = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "plantilla.html");
                string plantillaHtml = File.ReadAllText(rutaPlantillaHtml);

                plantillaHtml = plantillaHtml.Replace("###SISTEMA###", "Sistema de validación de cuentas personales");
                plantillaHtml = plantillaHtml.Replace("###FECHA###", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                plantillaHtml = plantillaHtml.Replace("###MONITOREO###", string.IsNullOrWhiteSpace(Monitoreo) ? "" : Monitoreo);
                plantillaHtml = plantillaHtml.Replace("###MENSAJE###", Mensaje);

                // Crear un LinkedResource para la primera imagen y asignar un CID
                string rutaImagen1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "Logo_Entel_connect_center.png");
                LinkedResource imagen1 = new LinkedResource(rutaImagen1);
                imagen1.ContentId = Guid.NewGuid().ToString(); // Un CID único

                // Sustituir la URL de la primera imagen en la plantilla HTML con el CID
                plantillaHtml = plantillaHtml.Replace("Logo_Entel_connect_center.png", "cid:" + imagen1.ContentId);

                // Crear un LinkedResource para la segunda imagen y asignar un CID
                string rutaImagen2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "r_Header.png");
                LinkedResource imagen2 = new LinkedResource(rutaImagen2);
                imagen2.ContentId = Guid.NewGuid().ToString(); // Otro CID único

                // Sustituir la URL de la segunda imagen en la plantilla HTML con el CID
                plantillaHtml = plantillaHtml.Replace("r_Header.png", "cid:" + imagen2.ContentId);

                // Crear una vista de cuerpo HTML y agregar los LinkedResources (imágenes) a ella
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(plantillaHtml, null, "text/html");
                htmlView.LinkedResources.Add(imagen1);
                htmlView.LinkedResources.Add(imagen2);

                // Agregar la vista de cuerpo HTML al mensaje
                mailMessage.AlternateViews.Add(htmlView);

                smtpClient.Send(mailMessage);


                //// Crear un LinkedResource para la imagen y asignar un CID
                //string rutaImagen = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "Logo_Entel_connect_center.png");
                //LinkedResource imagen = new LinkedResource(rutaImagen);
                //imagen.ContentId = Guid.NewGuid().ToString(); // Un CID único

                //// Sustituir la URL de la imagen en la plantilla HTML con el CID
                //plantillaHtml = plantillaHtml.Replace("Logo_Entel_connect_center.png", "cid:" + imagen.ContentId);

                //// Crear una vista de cuerpo HTML y agregar el LinkedResource (imagen) a ella
                //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(plantillaHtml, null, "text/html");
                //htmlView.LinkedResources.Add(imagen);

                //// Agregar la vista de cuerpo HTML al mensaje
                //mailMessage.AlternateViews.Add(htmlView);

                //smtpClient.Send(mailMessage);

                return new Retorno();
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrio una excepcion al enviar el mensaje a " + destinatario ?? "",
                    debug = ex.Message + ", Linea del error: " + new StackTrace(ex, true).GetFrame(0).GetFileLineNumber()
                };
            }
        }


        public Retorno EnviarMensajeConPlantillaLogCServicio(string destinatario, List<string> otrosDestinatarios, List<string> otrosDestinatariosOcultos, string Asunto, string Sistema, string Mensaje, string Monitoreo, ref string error)
        {
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

                SmtpClient smtpClient = new SmtpClient(this.Servidor, this.PuertoSMTP);
                smtpClient.Credentials = new NetworkCredential(this.Usuario, this.Password);
                smtpClient.EnableSsl = true;

                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(this.Usuario);
                mailMessage.To.Add(destinatario);

                // Agregar copias, es decir los otros destinatarios
                if (otrosDestinatarios != null)
                {
                    foreach (var otroDestinatario in otrosDestinatarios)
                    {
                        mailMessage.CC.Add(otroDestinatario.Trim());
                    }
                }

                // Agregar copias ocultas, es decir los destinatarios ocultos
                if (otrosDestinatariosOcultos != null)
                {
                    foreach (var destinatarioOculto in otrosDestinatariosOcultos)
                    {
                        mailMessage.Bcc.Add(destinatarioOculto.Trim());
                    }
                }

                mailMessage.Subject = Asunto;

                // Leer la plantilla HTML desde la ruta especificada
                string rutaPlantillaHtml = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "plantilla.html");
                string plantillaHtml = File.ReadAllText(rutaPlantillaHtml);


                plantillaHtml = plantillaHtml.Replace("###SISTEMA###", Sistema);
                plantillaHtml = plantillaHtml.Replace("###FECHA###", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                plantillaHtml = plantillaHtml.Replace("###MONITOREO###", string.IsNullOrWhiteSpace(Monitoreo) ? "" : Monitoreo);
                plantillaHtml = plantillaHtml.Replace("###MENSAJE###", Mensaje);

                // Crear un LinkedResource para la primera imagen y asignar un CID
                string rutaImagen1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "Logo_Entel_connect_center.png");
                LinkedResource imagen1 = new LinkedResource(rutaImagen1);
                imagen1.ContentId = Guid.NewGuid().ToString(); // Un CID único

                // Sustituir la URL de la primera imagen en la plantilla HTML con el CID
                plantillaHtml = plantillaHtml.Replace("Logo_Entel_connect_center.png", "cid:" + imagen1.ContentId);

                // Crear un LinkedResource para la segunda imagen y asignar un CID
                string rutaImagen2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "r_Header.png");
                LinkedResource imagen2 = new LinkedResource(rutaImagen2);
                imagen2.ContentId = Guid.NewGuid().ToString(); // Otro CID único

                // Sustituir la URL de la segunda imagen en la plantilla HTML con el CID
                plantillaHtml = plantillaHtml.Replace("r_Header.png", "cid:" + imagen2.ContentId);

                // Crear una vista de cuerpo HTML y agregar los LinkedResources (imágenes) a ella
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(plantillaHtml, null, "text/html");
                htmlView.LinkedResources.Add(imagen1);
                htmlView.LinkedResources.Add(imagen2);

                // Agregar la vista de cuerpo HTML al mensaje
                mailMessage.AlternateViews.Add(htmlView);

                smtpClient.Send(mailMessage);


                //// Crear un LinkedResource para la imagen y asignar un CID
                //string rutaImagen = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "Logo_Entel_connect_center.png");
                //LinkedResource imagen = new LinkedResource(rutaImagen);
                //imagen.ContentId = Guid.NewGuid().ToString(); // Un CID único

                //// Sustituir la URL de la imagen en la plantilla HTML con el CID
                //plantillaHtml = plantillaHtml.Replace("Logo_Entel_connect_center.png", "cid:" + imagen.ContentId);

                //// Crear una vista de cuerpo HTML y agregar el LinkedResource (imagen) a ella
                //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(plantillaHtml, null, "text/html");
                //htmlView.LinkedResources.Add(imagen);

                //// Agregar la vista de cuerpo HTML al mensaje
                //mailMessage.AlternateViews.Add(htmlView);

                //smtpClient.Send(mailMessage);

                return new Retorno();
            }
            catch (Exception ex)
            {
                return new Retorno
                {
                    ret = "ERROR",
                    msg = "Ocurrio una excepcion al enviar el mensaje a " + destinatario ?? "",
                    debug = ex.Message + ", Linea del error: " + new StackTrace(ex, true).GetFrame(0).GetFileLineNumber()
                };
            }
        }
    }
}