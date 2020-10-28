using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using GemBox.Spreadsheet;
using System.Windows.Forms;
namespace task1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> data = new List<string>();

            string text = RtfFileAsPlainText(@"../../../task1/data/testDoc.rtf");
            if(text is null)
            {
                text = "File doesn't exist";
            }
            else
            {
                string[] reg = { @"(?<=Регистрационный номер сделки\t)([0-9\/]+)", @"(?<=Номер договора\t)([0-9]+)" ,
                                @"(?<=Счет контрагента\t)([0-9A-Z]+)" , @"(?<=Адрес контрагента\t)([\p{IsCyrillic},.\s\d]+)(?=\n)",@"(?<=Наименование договора\t)([\p{IsCyrillic}\s]+)(?=\n)" };
                foreach (string i in reg)
                {
                    data.Add(Get_Data(text, i));
                }
                Make_Xls(data);
                text = "Mission complete";
            }
 
            Console.WriteLine(text);
        }
        public static string RtfFileAsPlainText(string rtfPathName)
        {
            using (var rtf = new RichTextBox())
            {
                try
                {
                    rtf.Rtf = File.ReadAllText(rtfPathName);
                    return rtf.Text;
                }
                catch (Exception)
                {
                    return null;
                }

            }
            
        }

        public static string Get_Data(string text, string reg)
        {
            try
            {
                Regex regex = new Regex(@reg);
                MatchCollection matches = regex.Matches(text);
                return matches[0].Value;
            }
            catch (Exception)
            {
                return "Null";
            }
        }
        public static void Make_Xls(List<string> data)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = new ExcelFile();
            // Add a new worksheet to the Excel file.
            var worksheet = workbook.Worksheets.Add("1");
            int i = 0;
            string[] ch = { "A", "B", "C", "D", "E" };
            string[] name = { "Регистрационный номер сделки", "Номер договора", "Счет контрагента", "Адрес контрагента", "Наименование договора" };
            while (i < 5)
            {

                worksheet.Cells[ch[i] + 1.ToString()].Value = name[i];
                worksheet.Cells[ch[i] + 2.ToString()].Value = data[i];
                i++;
            }
            workbook.Save(@"../../../task1/data/Create.xlsx");

        }
    }
}
