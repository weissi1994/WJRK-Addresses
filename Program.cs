using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;
using WJRK;

namespace WJRK
{

    public class Address {
        public string Addr {get;set;}
        public int ZIP {get;set;}
        public string City {get;set;}
        public string Province {get;set;}
    }

    class Program
    {
        private static readonly string API_KEY = "<INSERT-YOUR-GOOGLE-MAPS-API-KEY-HERE>";

        private static readonly int FIRST_NAME_COLLUMN = 4;
        private static readonly int LAST_NAME_COLLUMN = 5;
        private static readonly int GENDER_COLLUMN = 6;
        private static readonly int BIRTHDAY_COLLUMN = 7;
        private static readonly int STREET_COLLUMN = 8;
        private static readonly int ZIP_COLLUMN = 9;
        private static readonly int CITY_COLLUMN = 10;
        private static readonly int PROVINCE_COLLUMN = 11;

        private static void ReadSheet(string FilePath) {
            Logger.Log(String.Format("Started on {0}", DateTime.Now.ToString()));
            FileInfo existingFile = new FileInfo(FilePath);
			using (ExcelPackage package = new ExcelPackage(existingFile))
			{
                int row = 2;
				ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
				while ( worksheet.Cells[row, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+1, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+2, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+3, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+4, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+5, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+6, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+7, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+8, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+9, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+10, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+11, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+12, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+13, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+14, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+15, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+16, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+17, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+18, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+19, FIRST_NAME_COLLUMN].Value != null || 
                        worksheet.Cells[row+20, FIRST_NAME_COLLUMN].Value != null) {
                    Logger.Info(String.Format("Working on {0} {1} {2}", worksheet.Cells[row, GENDER_COLLUMN].Value, worksheet.Cells[row, FIRST_NAME_COLLUMN].Value, worksheet.Cells[row, LAST_NAME_COLLUMN].Value));
                    if (worksheet.Cells[row, STREET_COLLUMN].Value == null) {Logger.Warn(String.Format("\tNo data for line '{0}'", row));}
                    else if (worksheet.Cells[row, ZIP_COLLUMN].Value == null && 
                             worksheet.Cells[row, CITY_COLLUMN].Value == null && 
                             worksheet.Cells[row, STREET_COLLUMN].Value != null) 
                    {
                        Address updated = UpdateAddress(worksheet.Cells[row, STREET_COLLUMN].Value.ToString() + ", Austria", API_KEY);
                        if (updated == null) {Logger.Error(String.Format("\tError resolving Address for '{0}' in Line: {1}", worksheet.Cells[row, STREET_COLLUMN].Value.ToString(), row)); row++; continue;}
                        worksheet.Cells[row, ZIP_COLLUMN].Value = updated.ZIP;
                        worksheet.Cells[row, CITY_COLLUMN].Value = updated.City;
                        worksheet.Cells[row, PROVINCE_COLLUMN].Value = updated.Province;
                        Logger.Success(String.Format("\tUpdated: {0}, {1} {2}, {3}",
                            worksheet.Cells[row, STREET_COLLUMN].Value,
                            worksheet.Cells[row, ZIP_COLLUMN].Value,
                            worksheet.Cells[row, CITY_COLLUMN].Value,
                            worksheet.Cells[row, PROVINCE_COLLUMN].Value));
                    }
                    else if (worksheet.Cells[row, CITY_COLLUMN].Value == null && 
                             worksheet.Cells[row, ZIP_COLLUMN].Value != null && 
                             worksheet.Cells[row, STREET_COLLUMN].Value != null) 
                    {
                        Address updated = UpdateAddress(worksheet.Cells[row, STREET_COLLUMN].Value.ToString() + ", "+ worksheet.Cells[row, ZIP_COLLUMN].Value.ToString(), API_KEY);
                        if (updated == null) {Logger.Error(String.Format("\tError resolving Address for '{0}' in Line: {1}", worksheet.Cells[row, STREET_COLLUMN].Value.ToString(), row)); row++; continue;}
                        worksheet.Cells[row, ZIP_COLLUMN].Value = updated.ZIP;
                        worksheet.Cells[row, CITY_COLLUMN].Value = updated.City;
                        worksheet.Cells[row, PROVINCE_COLLUMN].Value = updated.Province;
                        Logger.Success(String.Format("\tUpdated: {0}, {1} {2}, {3}",
                            worksheet.Cells[row, STREET_COLLUMN].Value,
                            worksheet.Cells[row, ZIP_COLLUMN].Value,
                            worksheet.Cells[row, CITY_COLLUMN].Value,
                            worksheet.Cells[row, PROVINCE_COLLUMN].Value));
                    }
                    else {
                        Logger.Warn(String.Format("\tNo need to Update: {0}, {1} {2}, {3}", 
                            worksheet.Cells[row, STREET_COLLUMN].Value,
                            worksheet.Cells[row, ZIP_COLLUMN].Value,
                            worksheet.Cells[row, CITY_COLLUMN].Value,
                            worksheet.Cells[row, PROVINCE_COLLUMN].Value));
                    }
                    row++;
                }
                package.Save();
                Logger.Log(String.Format("Stopped at Line: {0}", row));
            }
        }

        private static Address UpdateAddress(string query, string key) {
            // string url = String.Format(@"https://maps.googleapis.com/maps/api/place/textsearch/json?query={0}&location=48.2082,16.3738&radius=50000&language=de&key={1}", query.Replace(" ", "+"), key);
            string url = String.Format(@"https://maps.googleapis.com/maps/api/geocode/json?address={0}&key={1}", query.Replace(" ", "+"), key);
            WebRequest request = WebRequest.Create(url);
            WebResponse response = request.GetResponse();
            Stream data = response.GetResponseStream();
            StreamReader reader = new StreamReader(data);
            // json-formatted string from maps api
            string responseFromServer = reader.ReadToEnd();
            response.Close();
            dynamic stuff = JObject.Parse(responseFromServer);
            if (stuff.status == "ZERO_RESULTS") { return null; }
            dynamic tmp = stuff.results[0].address_components;
            Address ret = new Address{Addr = ""};
            foreach (var item in tmp)
            {
                if (item.types.ToObject<List<string>>().Contains("postal_code")){
                    ret.ZIP = item.long_name;
                } else if (item.types.ToObject<List<string>>().Contains("locality")){
                    ret.City = item.long_name;
                } else if (item.types.ToObject<List<string>>().Contains("sublocality")){
                    ret.Province = item.long_name;
                } else if (item.types.ToObject<List<string>>().Contains("route")){
                    ret.Addr = item.long_name + ret.Addr;
                } else if (item.types.ToObject<List<string>>().Contains("street_number")){
                    ret.Addr = ret.Addr + item.long_name;
                }
            }
            return ret;
        }

        private static bool CheckExists(string path){
            return File.Exists(path);
        }

        private static void help(){
            Console.WriteLine("usage: WJRK.exe <Path-to-Excel-List>");
        }

        static void Main(string[] args)
        {
            if (args.Any("-h".Contains) || args.Any("--help".Contains) || args.Any("?".Contains)){
                help();
            } else if (args.Count() == 1) {
                if (CheckExists(args[0])) {
                // @"C:\\Users\d.weissengruber\Kursteilnehmerliste 2017_Zivi.xlsx", 5867, 6091
                    ReadSheet(args[0]);
                    Console.ReadKey();
                } else {
                    Logger.Error("File not accessible.");
                    Console.ReadKey();
                }
            } else {help();}
        }
    }
}
