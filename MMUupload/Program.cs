using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

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
            _ = ws.UsedRange.Columns.Count;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            _ = ws.get_Range("A1", last);
            Range uknot = ws.Columns["Q"]; // column to count UK or NON UK sends
            _ = last.Row;

            var UK = application.WorksheetFunction.CountIf(uknot, "UK"); // count uk sends
            var NONUK = application.WorksheetFunction.CountIf(uknot, "Non-UK"); // count nonuk sends



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
            //ws.Range["AF1"].Value = "BACKGROUND";
            ws.Range["U1"].Value = "UKORNONUK";

            // fill columns //

            ws.Range["R2:R" + LastRow].Value = "2";

            for (int i = 2; i < LastRow + 1; i++)
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
                .Replace("&", "").Replace("(", "").Replace(")", "").Replace("\"", "").Replace("-", "").Replace(@"\", "").Replace("+", ""); // Replace chars in purl

                char l1 = RandomLetter.GetLetter();
                int n1 = RandomNumber.GetNumber();
                char l2 = RandomLetter.GetLetter();
                int n2 = RandomNumber.GetNumber();
                char l3 = RandomLetter.GetLetter();
                int n3 = RandomNumber.GetNumber();

                ws.Range["T" + i].Value = Sname + l1 + n1 + l2 + n2 + l3 + n3;
            }

            //} // Generate Purls 


            // find what was the last live send and populate column with next increment of number. AGSEQ number //
            SqlDataReader dataReader;
            SqlCommand command;

            string sql = "SELECT TOP 52 * FROM mmu_offer_guide ORDER BY AG_SEQ DESC";
            SqlConnection conn = new SqlConnection(
                 new SqlConnectionStringBuilder()
                 {
                     DataSource = "AGSQL01",
                     InitialCatalog = "AG",
                     UserID = "AG_DB_autoapp",
                     Password = "AGuserRTP9845!"
                 }.ConnectionString
                );

            conn.Open();

            command = new SqlCommand(sql, conn);
            dataReader = command.ExecuteReader();
            //Create a new DataTable.
            var dt = new System.Data.DataTable();
            command.Dispose();

            dt.Load(dataReader);
            DataRow lr = dt.Rows[dt.Rows.Count - 1];
            long lr12 = Convert.ToInt64(lr["AG_SEQ"]);
            dataReader.Close();
            dt.Clear();

            for (int q = 2; q < LastRow + 1; q++)
            {
                ws.Range["A" + q].Value = lr12++;
            }

            SqlDataReader dataReader1;

            string sql1 = "select distinct [SPARE2] from mmu_offer_guide order by spare2 DESC";
            command = new SqlCommand(sql1, conn);
            dataReader1 = command.ExecuteReader();
            command.Dispose();
            int LiveSend = 0;


            var dt1 = new System.Data.DataTable();
            dt1.Load(dataReader1);
            DataRow row = dt1.Rows[0];
            LiveSend = Convert.ToInt32(row.ItemArray[0].ToString().Substring((row.ItemArray[0].ToString().Length - 1))) + 1;
            dataReader1.Close();
            ws.Range["W2:W" + LastRow].Value = "LIVESEND" + LiveSend;

            exceldoc.SaveAs(@"C:\Users\Sumptons\Desktop\testMMU.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

            exceldoc.Close(false);
            conn.Close();

            BackgroundWorker bw = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Sumptons\Desktop\testMMU.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""");
            OleDbConnection Econn = new OleDbConnection(constr);

            string Query = string.Format("Select * FROM [Sheet1$]");

            //[AG_SEQ],[MMU_ID],[FORNAME],[SURNAME] ,[AOS_CODE] ,[FULL_DESC] ,[FACULTY] ,[DEPARTMENT] ,[EXT_EMAIL] ,[ADD_1] ,[ADD_2] ,[ADD_3] ,[ADD_4],[POST_CODE],[OFFER],[STAGE_DATE],[Microsite],[VISIT_DAY],[SUBMIT_DATE]," +
            //"[PURL],[UKORNONUK],[SPARE1],[SPARE2],[SPARE3],[SPARE4],[SPARE5],[SPARE6],[SPARE7],[SPARE8],[SPARE9],[SPARE10]

            OleDbCommand Ecom = new OleDbCommand(Query, Econn);
            Econn.Open();

            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econn);
            oda.Fill(ds);
            System.Data.DataTable Exceldt = ds.Tables[0];


            SqlBulkCopy objbulk = new SqlBulkCopy(conn);

            objbulk.DestinationTableName = "mmu_offer_guide";
            //Mapping Table column    

            conn.Open(); //Open DataBase conection  

            objbulk.WriteToServer(Exceldt); //inserting Datatable Records to DataBase con.Close(); //Close DataBase conection  

           Console.WriteLine("Data has been Imported successfully.", "Imported");



            // Find 25 random rows in spreadsheet

            // Random rnd = new Random();
            //for (int i = 1; i < 24; i++)
            //{
            //   int num = rnd.Next(1, LastRow);
            //    ws.Range["AA" + num].Value = "RANDOM25";

            //}

            dataReader.Close(); // close data readers
            dataReader1.Close();
            command.Dispose(); // dispose of used command


            string sql2 = "SELECT distinct [MICROSITE] FROM[AG].[dbo].[mmu_offer_guide] where SPARE2 = 'LIVESEND" + LiveSend + "'";
            command = new SqlCommand(sql2, conn);
            dataReader1 = command.ExecuteReader(); //get distinct microsites for random25
            command.Dispose();

            List<string> MSlist = (from IDataRecord r in dataReader1
                                   select (string)r["MICROSITE"]).ToList(); // add to list
            dataReader1.Close(); // close data reader

            foreach (string i in MSlist) // for each microsite in list set the top 2 as random25
            {

                sql2 = "update TOP (2) mmu_offer_guide set[SPARE6] = 'RANDOM25' where [SPARE2] = 'LIVESEND" + LiveSend + "'" + " And [MICROSITE] =" + "'" + i + "'";
                command = new SqlCommand(sql2, conn);
                dataReader1 = command.ExecuteReader();
                dataReader1.Close();
            }


            // count records with random25 and if under 25 add more and non uk

            sql2 = "SELECT COUNT ([MICROSITE]) FROM[AG].[dbo].[mmu_offer_guide] where SPARE2 = 'LIVESEND" + LiveSend + "' AND SPARE6 = 'RANDOM25'";
            command = new SqlCommand(sql2, conn);
            Int32 count = (Int32)command.ExecuteScalar();
            command.Dispose();

            if ( count < 25) // if count is less than 25 add in extra records to "No Microsite"
            {
                int missing = 25 - count;
                    sql2 = "update TOP" +  "("+ (missing + 2) + ") mmu_offer_guide set[SPARE6] = 'RANDOM25' where [SPARE2] = 'LIVESEND" + LiveSend + "'" + " And [MICROSITE] = 'No Microsite'";
                    command = new SqlCommand(sql2, conn);
                    dataReader1 = command.ExecuteReader();
                    dataReader1.Close();
                command.Dispose();
            }

            sql2 = "SELECT COUNT ([MICROSITE]) FROM[AG].[dbo].[mmu_offer_guide] where SPARE2 = 'LIVESEND" + LiveSend + "' AND SPARE6 = 'RANDOM25'";
            command = new SqlCommand(sql2, conn);
            count = (Int32)command.ExecuteScalar();

            if (count >= 25)
            {
                int more = (count - 25) + 1;
                sql2 = "update TOP" + "(" + (more) + ") mmu_offer_guide set[SPARE6] = '' where [SPARE2] = 'LIVESEND" + LiveSend + "'" + " And [MICROSITE] = 'No Microsite'";
                command = new SqlCommand(sql2, conn);
                dataReader1 = command.ExecuteReader();
                dataReader1.Close();


              
                sql2 = "update TOP" + "(" + (1) + ") mmu_offer_guide set[SPARE6] = 'RANDOM25' where [SPARE2] = 'LIVESEND" + LiveSend + "'" + " And [UKORNONUK] = 'Non-UK'";
                command = new SqlCommand(sql2, conn);
                dataReader1 = command.ExecuteReader();
                dataReader1.Close();

            }
           

            // if over or exact delete as needed and include non uk
        }
     }

        class RandomLetter
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

        class RandomNumber
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


//REPLICATE(FORENAME + ' ',15000/len(forename)) as BACKGROUND