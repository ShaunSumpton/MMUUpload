using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Outlook;

namespace MMUupload
{
    class Program
    {
        static void Main(string[] args)
        {

            string[] filePaths = Directory.GetFiles(@"\\6.1.1.37\SFTPRoot\Manchester Metropolitan University", "*.xlsx"); // find worksheet on SFTP

            Console.WriteLine(filePaths[0]);

            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application(); // Create Excel Instance

            Workbook exceldoc = application.Workbooks.Open(filePaths[0]); // create workbook
            Worksheet ws; // create worksheet

            ws = (Worksheet)exceldoc.Worksheets[1]; // worksheet assigned to 1st sheet in workbook


            int LastRow = ws.UsedRange.Rows.Count;    // find last row and last column of sheet
            int LastCol = ws.UsedRange.Columns.Count;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = ws.get_Range("A1", last);
            Range uknot = ws.Columns["Q"]; // column to count UK or NON UK sends
            int lastUsedRow = last.Row;

            var UK = application.WorksheetFunction.CountIf(uknot, "UK"); // count uk sends
            var NONUK = application.WorksheetFunction.CountIf(uknot, "Non-UK"); // count nonuk sends

            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application(); // create outlook instance
            MailItem mailItem = app.CreateItem(OlItemType.olMailItem); // create mail item




            mailItem.Subject = "MMU Data Notification ";                                                    // set up email with to,subject, body etc
            mailItem.To = "s.sumpton@agnortheast.com;"; //S.kent@agnortheast.com";

            // mailItem.Attachments.Add(UrisGroup.dir + "\\" + UrisGroup.JobNumber + " " + UrisGroup.tc + " Booklet.pgp");
            mailItem.Importance = OlImportance.olImportanceHigh;
            mailItem.Display(false); // dont display mail item before sending



            var signature = mailItem.HTMLBody;
            var body = "MMU Offer Guide Quantities <br /> <br />" + "Number of UK: " + UK + "<br />" + "Number of Non-UK: " + NONUK;
            mailItem.HTMLBody = body; //+ signature;
            mailItem.Send(); // send email confirming data count

            // create and set columns headers //

            Range aRng = ws.Range["A1"];
            aRng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                    XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            ws.Range["A1"].Value = "AG_SEQ";

            Range rRng = ws.Range["R1"];
            rRng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                    XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            ws.Range["R1"].Value = "PURL";

            Range r1Rng = ws.Range["R1"];
            r1Rng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                    XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            ws.Range["R1"].Value = "SUBMIT_DATE";

            Range r2Rng = ws.Range["R1"];
            r2Rng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                    XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            ws.Range["R1"].Value = "VISIT_DAY";


            ws.Range["V1"].Value = "SPARE1";
            ws.Range["W1"].Value = "SPARE2";
            ws.Range["X1"].Value = "SPARE3";
            ws.Range["Y1"].Value = "SPARE4";
            ws.Range["Z1"].Value = "SPARE5";
            ws.Range["AA1"].Value = "SPARE6";
            ws.Range["AB1"].Value = "SPARE7";
            ws.Range["AC1"].Value = "SPARE8";
            ws.Range["AD1"].Value = "SPARE9";
            ws.Range["AE1"].Value = "SPARE10";
            ws.Range["AF1"].Value = "BACKGROUND";

            // fill columns //

            ws.Range["R2:R" + LastRow].Value = "2";

            for (int i = 2; i < LastRow; i++)
            {

                string temp = ws.Range["D" + i].Value;
                int iTemp = temp.Length;

                if (iTemp > 7)
                {
                    iTemp = 8;

                }


                // build perl

                string Sname = temp.Substring(0, iTemp);

                Sname = Sname.Replace("@", "").Replace(" ", "").Replace("/", "").Replace(".", "").Replace(",", "").Replace("'", "")
                .Replace("&", "").Replace("(", "").Replace(")", "").Replace("\"", "").Replace("-", "").Replace(@"\", "").Replace("+", "");

                char l1 = RandomLetter.GetLetter();
                int n1 = RandomNumber.GetNumber();
                char l2 = RandomLetter.GetLetter();
                int n2 = RandomNumber.GetNumber();
                char l3 = RandomLetter.GetLetter();
                int n3 = RandomNumber.GetNumber();

                ws.Range["T" + i].Value =Sname +  l1 + n1 + l2 + n2 + l3 + n3;

            }


            // find what was the last live send and populate column with next increment of number //
            //create purls using surname and random 6 charecter code "Sumpton9Z9Z9Z" //



            exceldoc.SaveAs(@"C:\Users\Sumptons\Desktop\testMMU.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);


            exceldoc.Close();

        }

        static class RandomLetter
        {
            static Random _random = new Random();
            public static char GetLetter()
            {
                // This method returns a random lowercase letter.
                // ... Between 'a' and 'z' inclusize.
                int num = _random.Next(0, 26); // Zero to 25
                char let = (char)('A' + num);
                return let;
            }
        }

        static class RandomNumber
        {
            // ... Create new Random object.
            static Random random = new Random();

            public static int GetNumber()
            {
                // ... Get 3 random numbers.
                //     These are always 5, 6, 7, 8 or 9.
                int num = random.Next(1, 10);
                return num;
            }


        }

    }
}

//REPLICATE(FORENAME + ' ',15000/len(forename)) as BACKGROUND