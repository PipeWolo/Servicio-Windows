using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace SERVICIO_ATT_VALIDACION_CUENTAS.App_Code
{
    public class NavigatorSMTP
    {
        string smtpServer { get; set; } // Servidor SMTP
        int smtpPort { get; set; } // Puerto SMTP
        string smtpUsername { get; set; } // Usuario SMTP
        string smtpPassword { get; set; } // Contraseña SMTP
        string CorreoDesde { get; set; } // Mi correo SMTO

        public NavigatorSMTP(string smtpServer, int smtpPort, string smtpUsername, string smtpPassword, string correoDesde)
        {
            this.smtpServer = smtpServer;
            this.smtpPort = smtpPort;
            this.smtpUsername = smtpUsername;
            this.smtpPassword = smtpPassword;
            CorreoDesde = correoDesde;
        }

        public bool EnviarMensaje(string destinatario, List<string> otrosDestinatarios, List<string> otrosDestinatariosOcultos, string Asunto, string Mensaje, ref string error)
        {
            try
            {
                SmtpClient smtpClient = new SmtpClient(this.smtpServer, this.smtpPort);
                smtpClient.Credentials = new NetworkCredential(this.smtpUsername, this.smtpPassword);
                smtpClient.EnableSsl = true;

                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(this.CorreoDesde);
                mailMessage.To.Add(destinatario);

                // Agregar copias, es decir los otros detinatarios
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
                mailMessage.Body = Mensaje;

                smtpClient.Send(mailMessage);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public bool EnviarMensajeConPlantilla(string destinatario, List<string> otrosDestinatarios, List<string> otrosDestinatariosOcultos, string Asunto, string NombrePersona, ref string error)
        {
            try
            {
                SmtpClient smtpClient = new SmtpClient(this.smtpServer, this.smtpPort);
                smtpClient.Credentials = new NetworkCredential(this.smtpUsername, this.smtpPassword);
                smtpClient.EnableSsl = true;

                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(this.CorreoDesde);
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

                // Reemplazar el marcador ###PERSONA### con el nombre de la persona
                plantillaHtml = plantillaHtml.Replace("###PERSONA###", NombrePersona);

                // Establecer el cuerpo del mensaje con la plantilla actualizada
                mailMessage.Body = plantillaHtml;
                mailMessage.IsBodyHtml = true;

                smtpClient.Send(mailMessage);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public bool EnviarMensajeConPlantillaLogCID(string destinatario, List<string> otrosDestinatarios, List<string> otrosDestinatariosOcultos, string Asunto, string NombrePersona, ref string error)
        {
            try
            {
                SmtpClient smtpClient = new SmtpClient(this.smtpServer, this.smtpPort);
                smtpClient.Credentials = new NetworkCredential(this.smtpUsername, this.smtpPassword);
                smtpClient.EnableSsl = true;

                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(this.CorreoDesde);
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

                // Reemplazar el marcador ###PERSONA### con el nombre de la persona
                plantillaHtml = plantillaHtml.Replace("###PERSONA###", NombrePersona);

                // Crear un LinkedResource para la imagen y asignar un CID
                string rutaImagen = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Recursos", "logo-entel-new.png");
                LinkedResource imagen = new LinkedResource(rutaImagen);
                imagen.ContentId = Guid.NewGuid().ToString(); // Un CID único

                // Sustituir la URL de la imagen en la plantilla HTML con el CID
                plantillaHtml = plantillaHtml.Replace("https://menu.en.tel/images/logo-entel-new.png", "cid:" + imagen.ContentId);

                // Crear una vista de cuerpo HTML y agregar el LinkedResource (imagen) a ella
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(plantillaHtml, null, "text/html");
                htmlView.LinkedResources.Add(imagen);

                // Agregar la vista de cuerpo HTML al mensaje
                mailMessage.AlternateViews.Add(htmlView);

                smtpClient.Send(mailMessage);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }
}
