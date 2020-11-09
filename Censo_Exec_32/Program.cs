using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Xml.Serialization;
using Newtonsoft.Json;

namespace Censo_Exec_32
{
    class Program
    {
        private const string URL = "http://intranethspm:5001/hspmsgh-api/censo/";
        public static DataTable CreateDataTable(List<Censo> arr)
        {
            XmlSerializer serializer = new XmlSerializer(arr.GetType());
            System.IO.StringWriter sw = new System.IO.StringWriter();
            serializer.Serialize(sw, arr);
            System.Data.DataSet ds = new System.Data.DataSet();
            System.Data.DataTable dt = new System.Data.DataTable();
            System.IO.StringReader reader = new System.IO.StringReader(sw.ToString());

            ds.ReadXml(reader);
            return ds.Tables[0];
        }


        public static string BlankFunction(string item)
        {
            return item ?? " "; 

        }


        public static string BlankFunctionTempo(string item)
        {
            return item.Replace("days", " ").Replace("day", " ").Replace("00:00:00", "0");

        }

        private static void Main(string[] args)
        {
            DateTime today = DateTime.Now;

            DataTable dataCenso = new DataTable();

            List<Censo> censos = new List<Censo>();


            WebRequest request = WebRequest.Create(URL);
            try
            {
                using (var twitpicResponse = (HttpWebResponse)request.GetResponse())
                {
                    using (var reader = new StreamReader(stream: twitpicResponse.GetResponseStream()))
                    {
                        JsonSerializer json = new JsonSerializer();
                        var objText = reader.ReadToEnd();
                        censos = JsonConvert.DeserializeObject<List<Censo>>(objText);
                        foreach (var obj in censos)
                        {
                            obj.tempo = BlankFunction(obj.tempo);
                            obj.tempo = BlankFunctionTempo(obj.tempo);

                        }

                        dataCenso = CreateDataTable(censos);

                    }
                }
            }

            catch (Exception ex)
            {
                String error = ex.Message;
                Console.ReadKey();

            }
            String excelFilePath = "\\\\hspmins2\\NIR_Nucleo_Interno_Regulacao\\2359\\Censo" + today.ToString().Replace('/', '_').Replace(' ', '_').Replace(':', '_');

            try
            {
                if (dataCenso == null || dataCenso.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)excelApp.ActiveSheet;

                // column headings
                for (var i = 0; i < dataCenso.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dataCenso.Columns[i].ColumnName;
                }

                // rows
                for (var i = 0; i < dataCenso.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < dataCenso.Columns.Count; j++)
                    {
                        /* if (j==9 || j == 10 || j == 13 || j == 24 || j == 25  )
                         {
                             var dt = dataCenso.Rows[i][j];
                             workSheet.Cells[i + 2, j + 1] = Convert.ToDateTime(dataCenso.Rows[i][j]);

                         }
                         else
                         {*/
                        workSheet.Cells[i + 2, j + 1] = dataCenso.Rows[i][j];


                        // }
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        //workSheet.Name = "Censo" + today.ToString().Replace('/', '_');
                        workSheet.Name = "Censo";
                        workSheet.SaveAs(excelFilePath);
                        excelApp.Quit();
                        Console.WriteLine("Excel file saved!");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                            + ex.Message);

                    }

                }
                else
                { // no file path is given
                    excelApp.Visible = true;

                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);

            }

        }


    }








    public class Censo
    {

        public string cd_prontuario { get; set; }

        public string nm_paciente { get; set; }

        public string nascimento { get; set; }
        public string nr_quarto { get; set; }

        public string dt_internacao_data { get; set; }

        public string dt_internacao_hora { get; set; }
        public string nm_especialidade { get; set; }

        public string nm_medico { get; set; }

        public string dt_ultimo_evento_data { get; set; }


        public string dt_ultimo_evento_hora { get; set; }

        public string nm_origem { get; set; }

        public string nr_convenio { get; set; }
        public string in_sexo { get; set; }

        public string nr_idade { get; set; }

        public string sg_cid { get; set; }
        public string descricao_cid { get; set; }

        public string nm_unidade_funcional { get; set; }

        public string tempo
        {
            get;set;
        }
        public string vinculo { get; set; }


        internal static IEnumerable<PropertyInfo> GetProperties()
        {
            throw new NotImplementedException();
        }
    }


}
