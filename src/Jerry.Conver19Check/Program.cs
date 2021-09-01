using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using ExcelDataReader;

namespace Jerry.Conver19Check
{
    class Program
    {
        public static readonly string _Url = "http://vaccinerecord.wsjkw.henan.gov.cn:38669/api/immurecord?name={0}&cardNo={1}";
        public static readonly string _Directory = Path.Combine(Environment.CurrentDirectory, "AccountInfo");
        private static readonly HttpClient _client = new HttpClient();
        private static Dictionary<string,string> _Dict=new Dictionary<string, string>();
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


            Console.WriteLine();
            Console.WriteLine("Read Excel from AccountInfo folder");
            ReadExcelToDict();

            Console.WriteLine();
            Console.WriteLine("Call API to get detail info");
            GetInfoByAPI();

            Console.WriteLine();
            Console.WriteLine("Save result to txt file");
            SaveToFile();
        }

        private static void SaveToFile()
        {

            // Set a variable to the Documents path.
            //string docPath = Environment.CurrentDirectory; //Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            

            // Write the string array to a new file named "WriteLines.txt".
            string filePath = Path.Combine(_Directory, $"result_{DateTime.Now.ToString("yyyy-MM-ddhhmmssfff")}.txt");
            using (StreamWriter outputFile = new StreamWriter(filePath))
            {
                foreach (var item in _Dict)
                {
                    outputFile.WriteLine($"{item.Key}|{item.Value}");
                }

                Console.WriteLine("File path: " + filePath);
            }
        }

        private static void GetInfoByAPI()
        {
            foreach (var item in _Dict)
            {
                var infos = item.Key.Split("|");
                var url = string.Format(_Url, infos[0], infos[1]);
                var response = _client.GetStringAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
                var result = JsonSerializer.Deserialize<ImmuRecord>(response);
                //{"msg":"成功","code":1,"data":"[{\"depaCode\":\"4104231001\",\"fullvaccination\":\"N\",\"inocBactCode\":\"5601\",\"inocBactName\":\"新冠疫苗（Vero细胞）\",\"inocBatchno\":\"202107116N\",\"inocCorpName\":\"北京科兴中维\",\"inocDate\":\"2021-08-30 09:51:38\",\"inocDepaName\":\"鲁山县熊背卫生院\",\"inocTime\":\"1\"}]","success":true}
                if (result.success)
                {
                    _Dict[item.Key] = result.data;
                    Console.WriteLine($"{item.Key}|{_Dict[item.Key]}");
                }
                else
                {
                    Console.WriteLine("call api error: account info:"+ item);
                }

                Thread.Sleep(200);
            }
        }

        public static bool ReadExcelToDict()
        {
            // path to your excel file
            //var filePath = Path.Combine(_Directory, "test.xlsx");
            //if (!File.Exists(filePath)) return false;
            
            var files = Directory.GetFiles(_Directory, "*.xls*");
            if (files == null || files.Length == 0)
            {
                Console.WriteLine("can not find the excel file in AccountInfo folder!");
                return false;
            }

            using (var stream = File.Open(files[0], FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var rows = reader.RowCount;
                    var cols = reader.FieldCount;

                    while (reader.Read())
                    {
                        if (reader.Depth < 2)
                            continue;

                        string key = $"{reader[3]}|{reader[4]}";
                        if (!_Dict.ContainsKey(key))
                        {
                            _Dict.Add(key, null);
                            //Console.WriteLine(key);
                        }

                    }


                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    //do
                    //{
                    //    while (reader.Read())
                    //    {
                    //        for (int j = 1; j < cols; j++)
                    //        {
                    //            Console.WriteLine(reader[j]);
                    //        }

                    //        // reader.GetDouble(0);
                    //    }
                    //} while (reader.NextResult());

                    //// 2. Use the AsDataSet extension method
                    ////var result = reader.AsDataSet();

                    //// The result of each spreadsheet is in result.Tables
                }
            }


            return true;
        }
    }

    class ImmuRecord
    {
        public string msg { get; set; }
        public int code { get; set; }
        public string data { get; set; }
        public bool success { get; set; }
    }
}
