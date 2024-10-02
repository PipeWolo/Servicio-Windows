using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace SERVICIO_ATT_VALIDACION_CUENTAS.App_Code
{
    class NavigatorLibreria
    {
        public bool ValidarEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Convierte una lista en un Datatable
        /// </summary>
        /// <returns> retorna un objeto DataTable</returns>
        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties by using reflection   
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names 
                try
                {
                    dataTable.Columns.Add(prop.Name, prop.PropertyType);
                }
                catch (Exception ex)
                {
                    string exs = ex.Message;
                    dataTable.Columns.Add(prop.Name, System.Type.GetType("System.DateTime"));
                }

            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    try
                    {
                        values[i] = Props[i].GetValue(item, null);
                    }
                    catch (Exception)
                    {
                        values[i] = "";
                    }
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public DateTime? GetFechaDataTime(string dateString)
        {
            CultureInfo culturaChile = new CultureInfo("es-CL");
            CultureInfo culturaUSA = new CultureInfo("en-US");
            DateTime result = new DateTime();

            if (String.IsNullOrEmpty(dateString))
            {
                return null;
            }

            //  Espacio: " "
            //  Tabulación: "\t"
            //  Salto de línea: "\n"(para Windows: "\r\n")
            //  tabulación vertical "\v"
            //  Avance de página: "\f"
            //  Salto de línea de carro: "\r"
            //  secuencia de escape "\0" representa un carácter nulo.
            //  la T que se usa en sistemas como SQL SERVER

            dateString = dateString.Replace("T", " ").Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Replace("\v", " ").Replace("\f", " ");
            dateString = Regex.Replace(dateString, @"\s+", " ");

            try
            {
                string[] formats = {
                    "yyyy-MM-dd HH:mm:ss",
                    "MM/dd/yyyy HH:mm:ss",
                    "dd/MM/yyyy HH:mm:ss",
                    "d/MM/yyyy HH:mm:ss",
                    "dd/M/yyyy HH:mm:ss",
                    "dd/MM/yy HH:mm:ss",
                    "d/M/yyyy HH:mm:ss",
                    "d/M/yy HH:mm:ss",
                    "dd/MM/yyyy HH:mm",
                    "d/MM/yyyy HH:mm",
                    "d/M/yyyy HH:mm",
                    "d/M/yy HH:mm",
                    "d/MM/yy HH:mm",
                    "dd/M/yyyy HH:mm",
                    "dd/M/yy HH:mm",
                    "yyyy/MM/dd HH:mm:ss",
                    "dd-MM-yyyy HH:mm:ss",
                    "dd-MM-yyyy hh:mm:ss tt",
                    "MM/dd/yyyy hh:mm:ss tt",
                    "dd/MM/yyyy hh:mm:ss tt",
                    "yyyy/MM/dd hh:mm:ss tt",
                    "HH:mm:ss dd-MM-yyyy",
                    "hh:mm:ss tt MM/dd/yyyy",
                    "hh:mm:ss tt dd/MM/yyyy",
                    "hh:mm:ss tt yyyy/MM/dd",
                    "MM/dd/yyyy h:mm:ss tt",
                    "MM/d/yyyy h:mm:ss tt",
                    "MM/d/yy h:mm:ss tt",
                    "MM/dd/yy h:mm:ss tt",
                    "M/dd/yyyy h:mm:ss tt",
                    "M/d/yyyy h:mm:ss tt",
                    "M/d/yy h:mm:ss tt",
                    "dd-MM-yy H:mm",
                    "dd/MM/yyyy h:mm:ss tt",
                    "MM/dd/yyyy h:mm:ss tt",
                    "M-dd-yy H:mm",
                    "dd-M-yy H:mm",
                    "yyyy-MM-dd hh:mm:ss tt"
                };

                if (DateTime.TryParseExact(dateString, formats, culturaChile, DateTimeStyles.None, out result))
                {
                    return result;
                }
                else
                {
                    if (DateTime.TryParseExact(dateString, formats, culturaUSA, DateTimeStyles.None, out result))
                    {
                        return result;
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }
}
