using System;
using System.Configuration;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;

namespace Navigator.Librerias
{
    /// <summary>
    /// Clase que encapsula funcionalidad de manejo de archivos usando FTP
    /// </summary>
    public class FTP
    {
        private string m_ftpserver = String.Empty;
        private string m_usuario = String.Empty;
        private string m_password = String.Empty;
        private String m_error = String.Empty;
        
        /// <summary>
        /// La direccion IP del servidor FTP
        /// </summary>
        public string FtpServer {
            get { return this.m_ftpserver; }
            set { this.m_ftpserver = value; }
        }

        /// <summary>
        /// El usuario FTP de la conexion
        /// </summary>
        public string Usuario {
            get { return this.m_usuario; }
            set { this.m_usuario = value; }
        }

        /// <summary>
        /// La password de la conexion al servidor FTP
        /// </summary>
        public string Password {
            get { return this.m_password; }
            set { this.m_password = value; }
        }

        /// <summary>
        /// El ultimo mensaje de error generado por la aplicacion
        /// </summary>
        public string Error {
            get { return this.m_error; }
            set { this.m_error = value; } 
        }

        /// <summary>
        /// Conectar al servidor FTP
        /// </summary>
        /// <returns>Verdadero si la conexion al servidor FTP fue exitosa</returns>
        public bool Conectar() 
        {
            return true;
        }

        /// <summary>
        /// Sube un archivo al servidor
        /// </summary>
        /// <param name="filename">El archivo a subir al servidor</param>
        /// <returns>Verdadero si el archivo fue subido al servidor de forma exitosa</returns>
        public bool Upload(string filename)
        {
            bool ret = false;

            FileInfo fileInf = new FileInfo(filename);
            string uri = this.m_ftpserver + "/" + fileInf.Name;
            FtpWebRequest reqFTP;

            reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(this.m_ftpserver + "/" + fileInf.Name));
            reqFTP.Credentials = new NetworkCredential(this.Usuario, this.m_password);
            reqFTP.KeepAlive = false;
            reqFTP.Method = WebRequestMethods.Ftp.UploadFile;
            reqFTP.UseBinary = true;
            reqFTP.ContentLength = fileInf.Length;

            int buffLength = 2048;
            byte[] buff = new byte[buffLength];
            int contentLen;

            FileStream fs = fileInf.OpenRead();

            this.m_error = "";

            try
            {
                Stream strm = reqFTP.GetRequestStream();

                contentLen = fs.Read(buff, 0, buffLength);

                while (contentLen != 0)
                {
                    strm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff, 0, buffLength);
                }

                strm.Close();
                fs.Close();

                ret = true;
            }
            catch (Exception ex)
            {
                this.m_error = ex.Message;
            }

            return ret;
        }

        /// <summary>
        /// Borrar un archivo desde el servidor FTP
        /// </summary>
        /// <param name="fileName">El archivo a eliminar</param>
        /// <returns>Verdadero si el archivo fue eliminado del servidor de forma exitosa</returns>
        public bool Delete(string fileName)
        {
            bool ret = false;

            try
            {
                string uri = this.m_ftpserver + "/" + fileName;
                FtpWebRequest reqFTP;
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(this.m_ftpserver + "/" + fileName));

                reqFTP.Credentials = new NetworkCredential(this.m_usuario, this.m_password);
                reqFTP.KeepAlive = false;
                reqFTP.Method = WebRequestMethods.Ftp.DeleteFile;

                string result = String.Empty;
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                long size = response.ContentLength;
                Stream datastream = response.GetResponseStream();
                StreamReader sr = new StreamReader(datastream);
                result = sr.ReadToEnd();
                sr.Close();
                datastream.Close();
                response.Close();
                ret = true;
            }
            catch (Exception ex)
            {
                this.m_error = ex.Message;
            }

            return ret;
        }

        private string[] GetFilesDetailList()
        {
            string[] downloadFiles;
            try
            {
                StringBuilder result = new StringBuilder();
                FtpWebRequest reqFTP;
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(this.m_ftpserver + "/"));
                reqFTP.Credentials = new NetworkCredential(this.m_usuario, this.m_password);
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                WebResponse response = reqFTP.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    result.Append(line);
                    result.Append("\n");
                    line = reader.ReadLine();
                }

                result.Remove(result.ToString().LastIndexOf("\n"), 1);
                reader.Close();
                response.Close();
                
                return result.ToString().Split('\n');
            }
            catch (Exception ex)
            {
                this.m_error = ex.Message;
                downloadFiles = null;
                return downloadFiles;
            }
        }

        /// <summary>
        /// Obtiene la lista de archivos desde el servidor FTP
        /// </summary>
        /// <returns>Arreglo conteniendo la lista de archivos desde el servidor FTP</returns>
        public string[] GetFileList()
        {
            string[] downloadFiles;
            StringBuilder result = new StringBuilder();
            FtpWebRequest reqFTP;
            try
            {
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(this.m_ftpserver + "/"));
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(this.m_usuario, this.m_password);
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectory;
                WebResponse response = reqFTP.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    result.Append(line);
                    result.Append("\n");
                    line = reader.ReadLine();
                }
                result.Remove(result.ToString().LastIndexOf('\n'), 1);
                reader.Close();
                response.Close();
                return result.ToString().Split('\n');
            }
            catch (Exception ex)
            {
                this.m_error = ex.Message;
                downloadFiles = null;
                return downloadFiles;
            }
        }

        /// <summary>
        /// Bajar un archivo desde el servidor FTP
        /// </summary>
        /// <param name="filePath">El path del archivo</param>
        /// <param name="fileName">El nombre del archivo</param>
        /// <returns>Verdadero si el archivo fue descargado desde el servidor FTP de forma exitosa</returns>
        public bool Download(string filePath, string fileName)
        {
            bool ret = false;
            FtpWebRequest reqFTP;
            try
            {
                FileStream outputStream = new FileStream(filePath + "\\" + fileName, FileMode.Create);
                
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(this.m_ftpserver + "/" + fileName));
                reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(this.m_usuario, this.m_password);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();
                long cl = response.ContentLength;
                int bufferSize = 2048;
                int readCount;
                byte[] buffer = new byte[bufferSize];

                readCount = ftpStream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    outputStream.Write(buffer, 0, readCount);
                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                }

                ftpStream.Close();
                outputStream.Close();
                response.Close();

                ret = true;
            }
            catch (Exception ex)
            {
                this.m_error = ex.Message;
            }

            return ret;
        }

        /// <summary>
        /// Constructor sobrecargado de la clase
        /// </summary>
        /// <param name="ftpserver">Direccion IP del servidor FTP</param>
        /// <param name="usuario">Usuario de la conexion FTP</param>
        /// <param name="password">Password de la conexion FTP</param>
        public FTP(string ftpserver, string usuario, string password) {
            this.m_ftpserver = ftpserver;
            this.m_usuario = usuario;
            this.m_password = password;
        }

        /// <summary>
        /// Constructor base de la clase
        /// </summary>
        public FTP(string flag) {

           if (flag.Equals("RECLAMOS")) {
              this.FtpServer = "ftp://164.77.160.150/DRs/PBS/";
              this.Usuario = "fonasa";
              this.Password = "fonasa01";
           }
           else if (flag.Equals("WEBSCRIPT")) {
              this.FtpServer = ConfigurationManager.AppSettings.Get("FTP");
              this.Usuario = ConfigurationManager.AppSettings.Get("FTP_USER");
              this.Password = ConfigurationManager.AppSettings.Get("FTP_PASSWORD"); ;
           }
        }
    }
}
