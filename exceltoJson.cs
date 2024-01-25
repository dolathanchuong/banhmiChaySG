using Nancy.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;

namespace exceltoJson
{
    class Program
    {
        static void Main(string[] args)
        {
            var pathToExcel = @"C:\Users\Acer\Desktop\2000ImportStudentResult.xlsx";
            //var Stream = ReadFileToMemoryStream(pathToExcel);
            //var chuoi = "Điện thoại PH";
            //UTF8Encoding utf8 = new UTF8Encoding();
            //var _end = utf8.GetString(utf8.GetBytes(chuoi));
            var _listJson = ReadExcelasJSON(pathToExcel);

            var jsonObject = JsonNode.Parse(_listJson);

            var _pp = jsonObject.AsArray().Count;


            var banhmi = _listJson.Replace("[", "").Replace("]", "");
            #region Parse String to JsonObject
            //var _ArrStudent = banhmi.Split(new char[] { '}' });
            //Split(_ArrStudent);
            //foreach (var r in _ArrStudent)
            //{
            //    dynamic data = JObject.Parse(r + "}");
            //    var _pp = data.sTT;
            //}
            #endregion
            #region Write to txt file
            //// Set a variable to the Documents path.
            //string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            //// Append text to an existing file named "WriteLines.txt".
            //using (StreamWriter outputFile = new StreamWriter(Path.Combine(docPath, "WriteLines.txt"), true))
            //{
            //    outputFile.WriteLine(banhmi);
            //}
            #endregion
            Console.WriteLine(_listJson);
        }
        static string ReadExcelasJSON(string name)
        {
            try
            {
                var pathToExcel = name;
                var sheetName = "Sheet1";

                //This connection string works if you have Office 2007+ installed and your 
                //data is saved in a .xlsx file|| CharacterSet=65001;IMEX=1;FMT=Delimited;
                var connectionString = String.Format(@"
            Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source={0};
            Extended Properties=""Excel 12.0 Xml;HDR=YES;CharacterSet=65001;IMEX=1;""
        ", pathToExcel);

                //Creating and opening a data connection to the Excel sheet 
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(
                        @"SELECT * FROM [{0}$]",
                        sheetName
                        );


                    using (var rdr = cmd.ExecuteReader())
                    {
                        //LINQ query - when executed will create anonymous objects for each row
                        var query =
                            (from DbDataRecord row in rdr
                             select row).Select(x =>
                             {
                                 //dynamic item = new ExpandoObject();
                                 Dictionary<string, object> item = new Dictionary<string, object>();
                                 item.Add("STT", x[0]);
                                 item.Add("HoTenPH", x[1]);
                                 item.Add("DienThoaiPH", x[2]);
                                 item.Add("EmailPH", x[3]);
                                 item.Add("TenHocSinh", x[4]);
                                 item.Add("NgayThangNamSinh", x[5]);
                                 item.Add("GioiTinh", x[6]);
                                 item.Add("Lop_i_course", x[7]);
                                 item.Add("LopTiengViet", x[8]);
                                 item.Add("Issue", x[9]);
                                 //item.Add(rdr.GetName(0), x[0]);
                                 //item.Add(rdr.GetName(1), x[1]);
                                 //item.Add(rdr.GetName(2), x[2]);
                                 //item.Add(rdr.GetName(3), x[3]);
                                 //item.Add(rdr.GetName(4), x[4]);
                                 //item.Add(rdr.GetName(5), x[5]);
                                 //item.Add(rdr.GetName(6), x[6]);
                                 //item.Add(rdr.GetName(7), x[7]);
                                 //item.Add(rdr.GetName(8), x[8]);
                                 //item.Add(rdr.GetName(9), x[9]);
                                 return item;
                             });

                        //Generates JSON from the LINQ query
                        var json = new JavaScriptSerializer().Serialize(query);
                        //var json = JsonConvert.SerializeObject(query);
                        return json;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        public static void Split<T>(T[] array)
        {
            var _lenght = array.Length;
            int dem = 0;
            var _listDetailStudent = new List<Array>();
            var _listGlobalStudent = new List<Object>();
            for (int i = 0; i < _lenght; i++)
            {
                if (dem == 11)
                {
                    _listGlobalStudent.Add(_listDetailStudent);
                    dem = 0;
                }
                else
                {
                    _listDetailStudent = new List<Array>();
                    _listDetailStudent.Add(array.Take(i).ToArray());
                    dem++;
                }
            }
        }
        public static object ReadFileToMemoryStream( string fileExcel)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(fileExcel);

            Stream stream = new MemoryStream(bytes);
            var _ms = new MemoryStream();
            CopyStream(stream, _ms);
            stream.Close();
            return _ms;
        }
        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[16 * 1024];
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, read);
            }
        }
    }
    public class Student
    {
        public string STT { get; set; }
        public string HoTenPH { get; set; }
        public string DienThoaiPH { get; set; }
        public string EmailPH { get; set; }
        public string TenHocSinh { get; set; }
        public string NgayThangNamSinh { get; set; }
        public string GioiTinh { get; set; }
        public string Lop_i_course { get; set; }
        public string LopTiengViet { get; set; }
        public string Issue { get; set; }
    }
}
