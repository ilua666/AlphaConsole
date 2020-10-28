using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Spire.Xls;

namespace AlphaConsole
{
    class Program
    {
        static Regex trRegex = new Regex("<tr[^>]*>(.*?)<\\/tr>");
        static Regex tdRegex = new Regex("<td[^>]*>(.*?)<\\/td>");

        const string PARSE_DEAL_REG_ID = "Регистрационный номер сделки";
        const string PARSE_CONTRACT_ID = "Номер договора";
        const string PARSE_CONTRACTOR_ACCOUNT = "Счет контрагента";
        const string PARSE_CONTRACTOR_ADDRESS = "Адрес контрагента";
        const string PARSE_CONTRACT_NAME = "Наименование договора";

        const string OUT_DEAL_REG_ID = "Регистрационный номер сделки";
        const string OUT_CONTRACT_ID = "Номер договора";
        const string OUT_CONTRACTOR_ACCOUNT = "Счет контрагента";
        const string OUT_CONTRACTOR_ADDRESS = "Адрес контрагента";
        const string OUT_CONTRACT_NAME = "Наименование договора";

        const string EXCEL_FILE_NAME = "testExcel.xlsx";

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Console.WriteLine("Enter the file name, or leave it empty for \"testDoc.rtf\"");
            string fileName = Console.ReadLine();
            if (fileName != null)
            {
                fileName = "../../../testDoc.rtf";
            }
            string rtfText;
            try
            {
                rtfText = File.ReadAllText(fileName);
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("FileError");
                return;
            }
            string htmlRtfText = RtfPipe.Rtf.ToHtml(rtfText);
            List<List<string>> tableData = new List<List<string>>();

            MatchCollection documentRows = trRegex.Matches(htmlRtfText);
            int i = 0;
            foreach(Match row in documentRows)
            {
                tableData.Add(new List<string>());
                MatchCollection rowValues = tdRegex.Matches(row.Value);
                foreach (Match rowValue in rowValues)
                {
                    tableData[i].Add(rowValue.Groups[1].Value);
                }
                i++;
            }

            string dealRegId = tableData.Find((v) => v[0].Equals(PARSE_DEAL_REG_ID))[1];
            string contractId = tableData.Find((v) => v[0].Equals(PARSE_CONTRACT_ID))[1];
            string contractorAccount = tableData.Find((v) => v[0].Equals(PARSE_CONTRACTOR_ACCOUNT))[1];
            string contractorAddress = tableData.Find((v) => v[0].Equals(PARSE_CONTRACTOR_ADDRESS))[1];
            string contractName = tableData.Find((v) => v[0].Equals(PARSE_CONTRACT_NAME))[1];

            Console.WriteLine(dealRegId);
            Console.WriteLine(contractId);
            Console.WriteLine(contractorAccount);
            Console.WriteLine(contractorAddress);
            Console.WriteLine(contractName);

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["A1"].Text =
            //Excel.Application excel = new Excel.Application();
            //Excel.Workbooks workBooks = excel.Workbooks;
            //Excel.Workbook workBook = workBooks.Add();
            //var sheet = (Excel.Worksheet)excel.ActiveSheet;

            sheet.Range["A1"].Text = OUT_DEAL_REG_ID;
            sheet.Range["B1"].Text = OUT_CONTRACT_ID;
            sheet.Range["C1"].Text = OUT_CONTRACTOR_ACCOUNT;
            sheet.Range["D1"].Text = OUT_CONTRACTOR_ADDRESS;
            sheet.Range["E1"].Text = OUT_CONTRACT_NAME;

            sheet.Range["A2"].Text = dealRegId;
            sheet.Range["B2"].Text = contractId;
            sheet.Range["C2"].Text = contractorAccount;
            sheet.Range["D2"].Text = contractorAddress;
            sheet.Range["E2"].Text = contractName;

            workbook.SaveToFile(EXCEL_FILE_NAME);

        }
    }
}
