using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft;
using System.Data.SqlClient;
using System.Data;

namespace bridge
{
    class Program
    {
        static void Main(string[] args)
        {
            string door_number = ""; //params
            string quote_number = "";//"60870"; //same 
            door_number = args[0];
            quote_number = args[1];
            string rev_number = quote_number + "- Rev 1"; //+ quote_number.Substring(quote_number.Length - 1);

            string startFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\" + "DataOutput " + quote_number + "- Door Designer.DO"; ;//location
            string newFile = @"\\designsvr1\apps\Door Master\Orders\" + door_number + ".DO";
            string hardwareExcelFile = @"\\designsvr1\solidworks\DWDevelopment\Specifications\" + quote_number + @"\Documents\HWAllocation " + quote_number + "- Door Designer.xlsx";
            string checksheet = @"\\designsvr1\apps\all doors\CheckSheet.pdf";
            string packingFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\Packing List " + rev_number + ".xlsx"; //should be the default file path for the session for everyone
            string engineerFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\Engineers Notes word " + rev_number + ".docx";
            string newPackingLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\Packing List " + door_number + ".xlsx";
            string extraPackingLocation = @"\\DESIGNSVR1\terry\door_history 1\" + door_number + ".xlsx";
            string newEngineerLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\Engineer Notes " + door_number + ".docx";
            string newChecksheetLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\CheckSheet.pdf";
            string newHardwareExcelFile = @"\\designsvr1\subcontracts\Order_List_Data\" + door_number + ".csv";


            System.IO.Directory.CreateDirectory(@"\\designsvr1\apps\bridge_jobcard\" + door_number);
            //string fileName = "DataOutput " + quote_number + "- Door Designer.DO";

            //^^ we need to copy and move this file before editing 
            System.IO.File.Copy(startFile, newFile, true); //true = overwrite
            System.IO.File.Copy(checksheet, newChecksheetLocation, true); //true = overwrite

            string test = File.ReadAllText(newFile);
            //repplace the the - with ""
            quote_number = quote_number.Replace("-", "");
            test = test.Replace(quote_number, door_number);
            File.WriteAllText(newFile, test);
            int line_number = 224; //this is ALWAYS the beginning
            //vv will change
            for (int i = 0; i < door_number.Length; i++)
            {
                string singleDigit = door_number.Substring(i, 1);
                lineChanger(singleDigit, newFile, line_number);
                line_number = line_number + 1;
            }

            //here we change some stuff in the text file IF the door is a SR2
            string sql = "select dbo.door.id from dbo.door left join dbo.door_type on dbo.door.door_type_id = dbo.door_type.id where security_rating_level = 2 and door_id = " + door_number;
            int sr2 = 0;
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    string getData = Convert.ToString(cmd.ExecuteScalar());
                    if (getData != null)
                        sr2 = -1;
                    else
                        sr2 = 0;
                    conn.Close();
                }
            }

            if (sr2 == -1) //it is a SR2
            {
                //check what row 94 is
                int lineOfNumber = 94;
                string rowData = "";
                string row94 = "";
                string row96 = "";
                string sr2Check = File.ReadAllText(newFile);
                string[] arrLine = File.ReadAllLines(newFile);
                rowData = arrLine[lineOfNumber - 1].ToString();

                if (rowData != "112" || rowData != "111")
                {
                    row94 = arrLine[93].ToString(); //-1 because count starts at 0
                    arrLine[93] = "0";
                    arrLine[87] = row94.ToString();

                    row96 = arrLine[95].ToString(); // same 
                    arrLine[95] = "0";
                    arrLine[88] = row96.ToString();
                    File.WriteAllLines(newFile, arrLine);
                }
            }

            if (File.Exists(hardwareExcelFile) == true)
            {
                //edit the huw excel sheet thing
                Microsoft.Office.Interop.Excel.Application xlAppCSV = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbookCSV = xlAppCSV.Workbooks.Open(hardwareExcelFile);
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheetCSV = xlWorkbookCSV.Sheets[1]; // assume it is the first sheet
                Microsoft.Office.Interop.Excel.Range xlRangeCSV = xlWorksheetCSV.UsedRange; // get the entire used range

                //get the rows
                var hwFirstColumn = (string)(xlWorksheetCSV.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range).Value;
                var hwSecondtColumn = (string)(xlWorksheetCSV.Cells[1, 3] as Microsoft.Office.Interop.Excel.Range).Value;

                string[] splitFirstColumn = hwFirstColumn.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                string[] splitSecondColumn = hwSecondtColumn.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                //readd the new excel items
                for (int i = 0; i < splitFirstColumn.Count(); i++)
                {
                    xlWorksheetCSV.Cells[1][i + 1].Value2 = door_number.ToString();
                    xlWorksheetCSV.Cells[2][i + 1].Value2 = splitFirstColumn[i].ToString();
                }
                for (int i = 0; i < splitFirstColumn.Count(); i++)
                {
                    xlWorksheetCSV.Cells[1][i + 1].Value2 = door_number.ToString();
                    xlWorksheetCSV.Cells[3][i + 1].Value2 = splitSecondColumn[i].ToString();
                }

                Console.WriteLine("-------");
                for (int i = 0; i < 5; i++)
                {
                    Console.WriteLine(splitSecondColumn[i].ToString());
                }


                //temp = xlWorksheetCSV.Cells[2][0].Value2;

                xlWorksheetCSV.SaveAs(newHardwareExcelFile);
                xlWorkbookCSV.Close(true); //close the excel sheet
                xlAppCSV.Quit(); //close everything excel related so that theres no errors when the door program tries to connect 
            }


            //also edit the packing list 

            //at some point we are going to move this excel sheet to another directory too

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(packingFile);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange; // get the entire used range

            if (File.Exists(newPackingLocation))
                File.Delete(newPackingLocation);

            sql = "SELECT description FROM dbo.paint_to_door WHERE door_id = " + door_number;
            string touch_up = "";
            string pack_date = "";
            string stores_date = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (i == 0)
                                touch_up = dt.Rows[i][0].ToString();
                            else
                                touch_up = touch_up + " / " + dt.Rows[i][0].ToString();
                        }
                    }
                    conn.Close();
                }
            }
            catch { }


            try
            {
                using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT coalesce((CONVERT(VARCHAR(10), date_stores, 103)),'') ,coalesce(CONVERT(VARCHAR(10), date_pack, 103),'')  FROM dbo.door where id = " + door_number, conn))
                    {
                        DataTable dt = new DataTable();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                        stores_date = dt.Rows[0][0].ToString();
                        pack_date = dt.Rows[0][1].ToString();
                    }
                    conn.Close();
                }
            }
            catch
            { }
            xlWorksheet.Cells[5][7].Value2 = door_number.ToString();

            //stores date
            xlWorksheet.Cells[3][7].Value2 = stores_date.ToString();
            //packing date
            xlWorksheet.Cells[3][8].Value2 = pack_date.ToString();

            xlWorksheet.Cells[5][20].Value2 = "Touch up paint required: " + touch_up;
            xlWorksheet.SaveAs(newPackingLocation);
            xlWorkbook.Close(true); //close the excel sheet
            xlApp.Quit(); //close everything excel related so that theres no errors when the door program tries to connect 
            if (File.Exists(newEngineerLocation))
                File.Delete(newEngineerLocation);
            File.Copy(engineerFile, newEngineerLocation, true); // also move this one over with a newmono name
            File.Copy(newPackingLocation, extraPackingLocation, true); //copy to door history aswell

            //check if there is a entry in dbo.door_program
            sql = "select door_id FROM dbo.door_program WHERE door_id = " + door_number;
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    var getdata = cmd.ExecuteScalar();
                    if (getdata == null)
                    {
                        //This door does not exist in the program table so we need to insert it
                        sql = "INSERT INTO dbo.door_program (door_id,programed_by_id,program_note,program_date) VALUES(" + door_number + ",314,'Programmed by bridge system',getdate())";
                    }
                    else
                    {
                        //update the door as programmed by 314
                        sql = "UPDATE dbo.door_program set programed_by_id = 314,program_note = 'Programmed by bridge system',program_date = getdate() WHERE door_id = " + door_number;
                    }
                }
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    // cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
            Console.WriteLine(sql);
            //    Console.ReadLine();




            // Console.ReadLine(); //remove this otherwise when firing from the commandline it will hang a little
        }

        static void lineChanger(string newNumber, string fileName, int lineOfNumber)
        {
            string[] arrLine = File.ReadAllLines(fileName);
            arrLine[lineOfNumber - 1] = newNumber;
            File.WriteAllLines(fileName, arrLine);
        }


    }
}
