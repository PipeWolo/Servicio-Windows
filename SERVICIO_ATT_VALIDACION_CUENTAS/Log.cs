using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

public class Log
{
    public string File { get; set; }
    public string Path { get; set; }
    public string SeccionEvt { get; set; }

    public void Write(string log)
    {
         try
         {
            string fecha_archivo = String.Format("{0:00}", DateTime.Now.Year) +
                                   String.Format("{0:00}", DateTime.Now.Month) +
                                   String.Format("{0:00}", DateTime.Now.Day) + ".log";

            string fecha_log = String.Format("{0:00}", DateTime.Now.Day) + "/" +
                               String.Format("{0:00}", DateTime.Now.Month) + "/" +
                               String.Format("{0:00}", DateTime.Now.Year) + " " +
                               String.Format("{0:00}", DateTime.Now.Hour) + ":" +
                               String.Format("{0:00}", DateTime.Now.Minute) + ":" +
                               String.Format("{0:00}", DateTime.Now.Second) + ":" +
                               String.Format("{0:00}", DateTime.Now.Millisecond);

            using (var file = new StreamWriter(this.Path + fecha_archivo, true))
            {
               file.WriteLine("[" + fecha_log + "]" + log);
               file.Close();
            }

        }
        catch (Exception)
        {
            //EventLog.WriteEntry(this.SeccionEvt, "Write : " + ex.Message, EventLogEntryType.Error);
        }
    }

    public Log(string path) 
    {
       this.Path = path; 
    }
}