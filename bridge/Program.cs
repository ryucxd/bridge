using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;

namespace bridge
{
    class Program
    {
        static void Main(string[] args)
        {
            //below 2 need to be blank to run automated
            string door_number = "67097"; //params
            string quote_number = "80396-2-1";
            //door_number = args[0]; //uncomment these for automation
            //quote_number = args[1];//^^

            //wipe everything in that directory
            try
            {
                var directory = new DirectoryInfo(@"C:\temp\GTINPUT\") { Attributes = FileAttributes.Normal };
                foreach (var info in directory.GetFileSystemInfos("*", SearchOption.AllDirectories))
                {
                    info.Attributes = FileAttributes.Normal;
                }
                directory.Delete(true);
            }
            catch { }

            System.IO.Directory.CreateDirectory(@"C:\temp\GTINPUT\");


            //if the door is a thermal -- continue down the old path
            int thermal = 0;
            string door_type = "";

            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                string sql = "select coalesce(thermal,0) FROM dbo.door_type dt " +
                    "left join dbo.door d on dt.id = d.door_type_id " +
                    "where d.id = " + door_number;

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    thermal = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                }


                sql = "select door_type_description FROM dbo.door_type dt " +
                    "left join dbo.door d on dt.id = d.door_type_id " +
                    "where d.id = " + door_number;

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    door_type = cmd.ExecuteScalar().ToString();

                    //SqlDataAdapter da = new SqlDataAdapter(cmd);
                    //DataTable dt = new DataTable();

                    //da.Fill(dt);

                    //door_type = dt.Rows[0][0].ToString();
                    //infil_id = Convert.ToInt32(dt.Rows[0][1].ToString());
                }



                conn.Close();
            }

            if (thermal == -1 /*|| door_type.Contains(" SG ")*/)
            {
                string rev_number = quote_number + "- Rev 1"; //+ quote_number.Substring(quote_number.Length - 1);

                string startFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\" + "DataOutput " + quote_number + "- Door Designer.DO";//location
                                                                                                                                                                                  //string startFile = @"\\designsvr1\DROPBOX\" + "DataOutput " + quote_number + "- Door Designer.DO";
                                                                                                                                                                                  //string newFile = @"\\designsvr1\DROPBOX\" + "DataOutput " + quote_number + "- Door Designer2.DO";
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
                //test = test.Replace(@"\r\n\", @"\r\n0\"); //remove blank lines for a 0
                test = Regex.Replace(test, @"\r\n\r\n", Environment.NewLine + "0" + Environment.NewLine);
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
                string sql = "select dbo.door.id from dbo.door left join dbo.door_type on dbo.door.door_type_id = dbo.door_type.id where security_rating_level = 2 and dbo.door.id = " + door_number;
                int sr2 = 0;
                using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        string getData = Convert.ToString(cmd.ExecuteScalar());
                        //Console.WriteLine(getData);
                        //Console.ReadLine();
                        if (getData.Length > 4)
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

                    if (rowData != "112" && rowData != "111")
                    {
                        row94 = arrLine[93].ToString(); //-1 because count starts at 0
                        arrLine[93] = "0";
                        arrLine[87] = row94.ToString(); //88

                        row96 = arrLine[95].ToString(); // same 
                        arrLine[95] = "0";
                        arrLine[88] = row96.ToString(); //89
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
                    for (int i = 0; i < splitSecondColumn.Count(); i++)
                    {
                        xlWorksheetCSV.Cells[1][i + 1].Value2 = door_number.ToString();
                        xlWorksheetCSV.Cells[3][i + 1].Value2 = splitSecondColumn[i].ToString();
                    }



                    //temp = xlWorksheetCSV.Cells[2][0].Value2;

                    xlWorksheetCSV.SaveAs(newHardwareExcelFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
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
                            System.Data.DataTable dt = new System.Data.DataTable();
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
                            System.Data.DataTable dt = new System.Data.DataTable();
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

                    //update dbo.bridge_log
                    int max_bridge_id = 0;

                    sql = "select top 1 id from dbo.bridge_log where door_id = " + door_number + " order by id desc";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        max_bridge_id = (int)cmd.ExecuteScalar();

                    sql = "update dbo.bridge_log SET bridge_success = -1 WHERE id = " + max_bridge_id.ToString();
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        cmd.ExecuteNonQuery();

                    conn.Close();
                }
                Console.WriteLine(sql);
                // Console.ReadLine();
                // Console.ReadLine(); //remove this otherwise when firing from the commandline it will hang a little
            }
            else //this is NOT a thermal -- the idea here is to get everything from dbo.DWBridge and dump it into a GT Input
            {
                object misValue = System.Reflection.Missing.Value;

                string GT_input_location = @"\\designsvr1\apps\ALL DOORS\New Programmer Folder\GT INPUT";
                string new_GT_input_location = (@"C:\temp\GTINPUT\GT INPUT");

                if (door_type.Contains("SR2"))
                {

                    try
                    {
                        GT_input_location += " SR2.xlsm";
                        new_GT_input_location += " SR2.xlsm";

                        //^^ we need to copy and move this file before editing --temp for now
                        System.IO.File.Copy(GT_input_location, new_GT_input_location, true); //true = overwrite

                        //edit the huw excel sheet thing
                        Microsoft.Office.Interop.Excel.Application xlAppGTInput = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook xlWorkbookGTInput = xlAppGTInput.Workbooks.Open(new_GT_input_location, 0, false, 5, "", "", false,
                        Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        Microsoft.Office.Interop.Excel.Worksheet xlWorksheetGTInput = xlWorkbookGTInput.Sheets[1]; // assume it is the first sheet
                        Microsoft.Office.Interop.Excel.Range xlRangeCSV = xlWorksheetGTInput.UsedRange; // get the entire used range

                        //rename new_GT_input_location -- cant save it as the same name because of read/write issues
                        new_GT_input_location = @"C:\temp\GTINPUT\GT INPUT " + door_number + ".xlsm";


                        xlAppGTInput.DisplayAlerts = false;

                        using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                        {
                            conn.Open();

                            string sql = "select * " +
                                "FROM dbo.DWBridge dw " +
                                "left join dbo.door d on dw.SalesOrderNum = d.quote_number " +
                                "left join dbo.SALES_LEDGER s on d.customer_acc_ref = s.ACCOUNT_REF " +
                                "left join dbo.door_type dt on d.door_type_id = dt.id " +
                                "where d.id = " + door_number + " AND d.quote_number = '" + quote_number + "'";

                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                            {
                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                DataTable dt = new DataTable();

                                da.Fill(dt);

                                if (dt.Rows.Count == 0)
                                {
                                    Console.WriteLine("There is no record in DWBridge :(");
                                    Console.ReadLine();
                                    return;
                                }

                                //[col][row] //column 2 = B
                                xlWorksheetGTInput.Cells[2][1].Value2 = door_number;
                                xlAppGTInput.CalculateUntilAsyncQueriesDone();
                                xlWorksheetGTInput.Cells[2][2].Value2 = quote_number.ToString();
                                xlWorksheetGTInput.Cells[2][3].Value2 = dt.Rows[0]["NAME"].ToString();

                                sql = "select top 1 left(forename,1) + left(surname,1) FROM dbo.bridge_log b " +
                                    "left join [user_info].dbo.[user] u on b.staff_id = u.id " +
                                    "where door_id = " + door_number + " order by b.id desc";

                                using (SqlCommand cmdProgrammer = new SqlCommand(sql, conn))
                                {
                                    xlWorksheetGTInput.Cells[2][4].Value2 = cmdProgrammer.ExecuteScalar().ToString();
                                }

                                xlWorksheetGTInput.Cells[2][5].Value2 = dt.Rows[0]["quantity_same"].ToString();
                                xlWorksheetGTInput.Cells[2][6].Value2 = dt.Rows[0]["door_ref"].ToString();

                                if (dt.Rows[0]["double_y_n"].ToString() == "-1")
                                    xlWorksheetGTInput.Cells[2][8].Value2 = "Double A+B-Skin";
                                else
                                    xlWorksheetGTInput.Cells[2][8].Value2 = "Single A-Skin";

                                if (dt.Rows[0]["DoorStyle"].ToString() == "Double (Unequal Split)")
                                    xlWorksheetGTInput.Cells[2][9].Value2 = "Yes";


                                if (dt.Rows[0]["DoorStyle"].ToString().Contains("Double"))
                                {
                                    if (dt.Rows[0]["NAME"].ToString() == "JODAN CONTRACTS LTD")
                                    {
                                        xlWorksheetGTInput.Cells[2][13].Value2 = "Welded";
                                    }
                                    else
                                        xlWorksheetGTInput.Cells[2][13].Value2 = "Bolted";
                                }
                                else if (dt.Rows[0]["DoorStyle"].ToString().Contains("Single"))
                                {
                                    xlWorksheetGTInput.Cells[2][13].Value2 = "Welded";
                                }

                                xlWorksheetGTInput.Cells[2][15].Value2 = "Galv";

                                if (dt.Rows[0]["door_type_description"].ToString().Contains("Mortice"))
                                {
                                    xlWorksheetGTInput.Cells[2][18].Value2 = "Security Door Mortice";
                                }
                                else if (dt.Rows[0]["door_type_description"].ToString().Contains("Panic"))
                                {
                                    xlWorksheetGTInput.Cells[2][18].Value2 = "Security Door Panic";
                                }


                                xlWorksheetGTInput.Cells[2][20].Value2 = dt.Rows[0]["SOW"].ToString();
                                xlWorksheetGTInput.Cells[2][21].Value2 = dt.Rows[0]["SOH"].ToString();
                                //xlWorksheetGTInput.Cells[2][26].Value2 = dt.Rows[0]["hingeQty"].ToString();

                                //need to translate this
                                if (dt.Rows[0]["CillType"].ToString().Contains("Aluminium"))
                                    xlWorksheetGTInput.Cells[2][27].Value2 = "Standard Aluminium";
                                else if (dt.Rows[0]["CillType"].ToString() == "H Type")
                                    xlWorksheetGTInput.Cells[2][27].Value2 = "Cill H type";
                                else
                                    xlWorksheetGTInput.Cells[2][27].Value2 = dt.Rows[0]["CillType"].ToString();


                                //  xlWorksheetGTInput.Cells[2][28].Value2 = dt.Rows[0]["fixingType"].ToString();

                                //translate
                                if (dt.Rows[0]["hasJackingScrews"].ToString() == "Jacking Screws")
                                    xlWorksheetGTInput.Cells[2][29].Value2 = "Yes";


                                xlWorksheetGTInput.Cells[2][30].Value2 = dt.Rows[0]["fixingTo"].ToString();


                                //translate
                                if (dt.Rows[0]["Handing"].ToString().Contains("L"))
                                    xlWorksheetGTInput.Cells[2][31].Value2 = "Left Hand";
                                else if (dt.Rows[0]["Handing"].ToString().Contains("R"))
                                    xlWorksheetGTInput.Cells[2][31].Value2 = "Right Hand";

                                // xlWorksheetGTInput.Cells[2][32].Value2 = dt.Rows[0]["openingDirection"].ToString();



                                //center locks
                                sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["CentreLockStockCode"].ToString() + "'";
                                using (SqlCommand cmdCenterLock = new SqlCommand(sql, conn))
                                {
                                    var temp = cmdCenterLock.ExecuteScalar();
                                    if (temp != null)
                                    {
                                        xlWorksheetGTInput.Cells[2][34].Value2 = cmdCenterLock.ExecuteScalar().ToString();
                                    }
                                }

                                if (dt.Rows[0]["CentreLeverInside"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][36].Value2 = "Lever-Rose Fixed Assa 640 Un-sprung St.St.";

                                if (dt.Rows[0]["centrelockheight"].ToString().Length > 0)
                                    xlWorksheetGTInput.Cells[4][36].Value2 = dt.Rows[0]["centrelockheight"].ToString(); //special box

                                if (dt.Rows[0]["CentreLeverOutside"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][37].Value2 = "Lever-Rose Fixed Assa 640 Un-sprung St.St.";

                                if (dt.Rows[0]["CentreLockingInside"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][38].Value2 = "Yes";

                                if (dt.Rows[0]["CentreLockingOutside"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][39].Value2 = "Yes";


                                if (dt.Rows[0]["CentreInsideEscutcheonName"].ToString().Contains(" KEY / KEY ") && dt.Rows[0]["CentreOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                    xlWorksheetGTInput.Cells[2][40].Value2 = "Assa Full 31MM / 31MM SCP (Key o/s, Key i/s)";
                                else if (dt.Rows[0]["CentreInsideEscutcheonName"].ToString().Contains("KEY / TURN") && dt.Rows[0]["CentreOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                    xlWorksheetGTInput.Cells[2][40].Value2 = "Assa Half 31MM SCP (Key o/s, Thumbturn i/s)";
                                else if (dt.Rows[0]["CentreInsideEscutcheonName"].ToString().Contains("KEY / KEY ") && dt.Rows[0]["CentreOutsideEscutcheonName"].ToString().Contains("KEY / TURN"))
                                    xlWorksheetGTInput.Cells[2][40].Value2 = "Assa Half 31MM SCP (Thumbturn o/s, Key i/s)";

                                //door loop
                                if (dt.Rows[0]["DoorLoopType"].ToString().Contains("DL8"))
                                {
                                    xlWorksheetGTInput.Cells[2][42].Value2 = "Abloy DL8 Surface Mounted"; //same for the below??
                                    xlWorksheetGTInput.Cells[2][43].Value2 = "Active Leaf";
                                }
                                else if (dt.Rows[0]["DoorLoopType"].ToString().Contains("EA280"))
                                    xlWorksheetGTInput.Cells[2][42].Value2 = "Abloy EA280 Concealed";


                                //top lock
                                sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["TopLockStockCode"].ToString() + "'";
                                using (SqlCommand cmdTopLock = new SqlCommand(sql, conn))
                                {
                                    var temp = cmdTopLock.ExecuteScalar();
                                    if (temp != null)
                                    {
                                        xlWorksheetGTInput.Cells[2][45].Value2 = cmdTopLock.ExecuteScalar().ToString();

                                        if (dt.Rows[0]["ToplockingInside"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][47].Value2 = "Yes";
                                        if (dt.Rows[0]["TopLockingOutside"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][48].Value2 = "Yes";


                                        if (dt.Rows[0]["TopInsideEscutcheonName"].ToString().Contains(" KEY / KEY ") && dt.Rows[0]["TopOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                            xlWorksheetGTInput.Cells[2][49].Value2 = "Assa Full 31MM / 31MM SCP (Key o/s, Key i/s)";
                                        else if (dt.Rows[0]["TopInsideEscutcheonName"].ToString().Contains("KEY / TURN") && dt.Rows[0]["TopOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                            xlWorksheetGTInput.Cells[2][49].Value2 = "Assa Half 31MM SCP (Key o/s, Thumbturn i/s)";
                                        else if (dt.Rows[0]["TopInsideEscutcheonName"].ToString().Contains("KEY / KEY ") && dt.Rows[0]["TopOutsideEscutcheonName"].ToString().Contains("KEY / TURN"))
                                            xlWorksheetGTInput.Cells[2][49].Value2 = "Assa Half 31MM SCP (Thumbturn o/s, Key i/s)";

                                    }

                                }

                                //bot lock

                                sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["TopLockStockCode"].ToString() + "'";
                                using (SqlCommand cmdTopLock = new SqlCommand(sql, conn))
                                {
                                    var temp = cmdTopLock.ExecuteScalar();
                                    if (temp != null)
                                    {
                                        xlWorksheetGTInput.Cells[2][51].Value2 = cmdTopLock.ExecuteScalar().ToString();

                                        if (dt.Rows[0]["ToplockingInside"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][53].Value2 = "Yes";
                                        if (dt.Rows[0]["TopLockingOutside"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][54].Value2 = "Yes";
                                        //55 cylinderrrr
                                        if (dt.Rows[0]["BotInsideEscutcheonName"].ToString().Contains(" KEY / KEY ") && dt.Rows[0]["BottomOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                            xlWorksheetGTInput.Cells[2][55].Value2 = "Assa Full 31MM / 31MM SCP (Key o/s, Key i/s)";
                                        else if (dt.Rows[0]["BottomInsideEscutcheonName"].ToString().Contains("KEY / TURN") && dt.Rows[0]["BottomOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                            xlWorksheetGTInput.Cells[2][55].Value2 = "Assa Half 31MM SCP (Key o/s, Thumbturn i/s)";
                                        else if (dt.Rows[0]["BottomInsideEscutcheonName"].ToString().Contains("KEY / KEY ") && dt.Rows[0]["BottomOutsideEscutcheonName"].ToString().Contains("KEY / TURN"))
                                            xlWorksheetGTInput.Cells[2][55].Value2 = "Assa Half 31MM SCP (Thumbturn o/s, Key i/s)";

                                    }
                                }


                                //panics
                                sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["PanicDeviceStockCode"].ToString() + "'";
                                using (SqlCommand cmdPanic = new SqlCommand(sql, conn))
                                {
                                    var temp = cmdPanic.ExecuteScalar();
                                    if (temp != null)
                                    {
                                        xlWorksheetGTInput.Cells[2][57].Value2 = cmdPanic.ExecuteScalar().ToString();

                                        xlWorksheetGTInput.Cells[2][58].Value2 = xlWorksheetGTInput.Cells[2][59].Value2;


                                        sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["OADStockCode"].ToString() + "'";
                                        using (SqlCommand cmdOAD = new SqlCommand(sql, conn))
                                        {
                                            var temp2 = cmdOAD.ExecuteScalar();
                                            if (temp2 != null)
                                            {
                                                xlWorksheetGTInput.Cells[2][60].Value2 = cmdOAD.ExecuteScalar().ToString();
                                            }
                                        }

                                    }
                                }


                                //pushplate stuffs
                                if (dt.Rows[0]["PullHandleCode"].ToString() == "286")
                                {
                                    //if there is a pushplate then we use > "Pull Handle 19 x 300 Rose Mounted St.St. & 330mm x 76mm Push Plate"
                                    if (dt.Rows[0]["PushPlateType"].ToString().Length > 0)
                                        xlWorksheetGTInput.Cells[2][62].Value2 = "Pull Handle 19 x 300 Rose Mounted St.St. & 330mm x 76mm Push Plate";
                                    else
                                        xlWorksheetGTInput.Cells[2][62].Value2 = "Pull Handle 19 x 300 Rose Mounted St.St.";
                                }

                                //xlWorksheetGTInput.Cells[2][63].Value2 = dt.Rows[0]["pushPlateSide"].ToString(); //translate)
                                if (dt.Rows[0]["pushPlateSide"].ToString() == "Pull Side")
                                    xlWorksheetGTInput.Cells[2][63].Value2 = "Pullside";
                                else if (dt.Rows[0]["pushPlateSide"].ToString() == "Push Side")
                                    xlWorksheetGTInput.Cells[2][63].Value2 = "Pushside";
                                else if (dt.Rows[0]["pushPlateSide"].ToString() == "Both Side")
                                    xlWorksheetGTInput.Cells[2][63].Value2 = "Both sides";

                                if (dt.Rows[0]["pushPlateLeaves"].ToString() == "Active")
                                    xlWorksheetGTInput.Cells[2][64].Value2 = "1st Leaf";
                                else if (dt.Rows[0]["pushPlateLeaves"].ToString() == "Passive")
                                    xlWorksheetGTInput.Cells[2][64].Value2 = "2nd Leaf";
                                else if (dt.Rows[0]["pushPlateLeaves"].ToString() == "Active/Passive")
                                    xlWorksheetGTInput.Cells[2][64].Value2 = "Both Leafs";




                                //Closers

                                sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["CloserStockCode"].ToString() + "'";
                                using (SqlCommand cmdCloser = new SqlCommand(sql, conn))
                                {
                                    var temp = cmdCloser.ExecuteScalar();
                                    if (temp != null)
                                    {
                                        xlWorksheetGTInput.Cells[2][67].Value2 = cmdCloser.ExecuteScalar().ToString();

                                        xlWorksheetGTInput.Cells[2][68].Value2 = dt.Rows[0]["closerPullside"].ToString(); //closerPullside = 1 then pull side closerpushside = 1 then push
                                        if (dt.Rows[0]["closerPullside"].ToString() == "1" && dt.Rows[0]["closerpushside"].ToString() == "0")
                                            xlWorksheetGTInput.Cells[2][68].Value2 = "Pullside";
                                        else if (dt.Rows[0]["closerPullside"].ToString() == "0" && dt.Rows[0]["closerpushside"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][68].Value2 = "Pushside";

                                        if (dt.Rows[0]["closerOnActive"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][69].Value2 = "Yes";

                                        if (dt.Rows[0]["closerOnPassive"].ToString() == "1")
                                            xlWorksheetGTInput.Cells[2][71].Value2 = "Yes";
                                    }
                                }




                                if (dt.Rows[0]["StayLeaves"].ToString() == "Active")
                                    xlWorksheetGTInput.Cells[2][64].Value2 = "1st Leaf";
                                else if (dt.Rows[0]["StayLeaves"].ToString() == "Passive")
                                    xlWorksheetGTInput.Cells[2][64].Value2 = "2nd Leaf";
                                else if (dt.Rows[0]["StayLeaves"].ToString() == "Active/Passive")
                                    xlWorksheetGTInput.Cells[2][64].Value2 = "Both Leafs";

                                //end of closers




                                //stay
                                sql = "SELECT RTRIM(GT_input_name) FROM dbo.bridge_hardware bh " +
                                    "left join dbo.stock s on bh.stock_code = s.stock_code WHERE s.description = '" + dt.Rows[0]["StayType"].ToString() + "'";
                                using (SqlCommand cmdStay = new SqlCommand(sql, conn))
                                {
                                    if (string.IsNullOrEmpty(dt.Rows[0]["StayType"].ToString()))
                                    { }
                                    else
                                        xlWorksheetGTInput.Cells[2][74].Value2 = cmdStay.ExecuteScalar().ToString().Trim();
                                }

                                //leaf selector
                                if (dt.Rows[0]["LeafSelectorType"].ToString().Contains("MK2 SELECTOR EXTENDED  CATCH 152 ARM SAA"))
                                    xlWorksheetGTInput.Cells[2][77].Value2 = " c/w Extended Catch & Arm SAA (Wedge)";

                                if (dt.Rows[0]["SpyHoleType"].ToString().Length > 0)
                                    xlWorksheetGTInput.Cells[2][83].Value2 = "Zero 200 UL Door Viewer (Fire Rated)";


                                //vision / lourvre #1 ACTIVE
                                if (Convert.ToInt32(dt.Rows[0]["Active1VisionGlassThickness"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][90].Value2 = "Vision";
                                    xlWorksheetGTInput.Cells[2][91].Value2 = dt.Rows[0]["Active1VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][92].Value2 = dt.Rows[0]["Active1VisionLouvreWidth"].ToString();

                                    if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][93].Value2 = "Offset";
                                    else if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][93].Value2 = "Central";

                                    xlWorksheetGTInput.Cells[2][94].Value2 = dt.Rows[0]["Active1VisionLouvreDistanceFromFloor"].ToString();


                                }
                                else if (Convert.ToInt32(dt.Rows[0]["Active1VisionGlassThickness"].ToString()) == 0 &&
                                         Convert.ToInt32(dt.Rows[0]["Active1VisionLouvreHeight"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][90].Value2 = "Louver";
                                    xlWorksheetGTInput.Cells[2][91].Value2 = dt.Rows[0]["Active1VisionLouvreHeight"].ToString();

                                    //if this value is higher than the calculation to the right of it > set it to the calculation limit
                                    if (xlWorksheetGTInput.Cells[2][91].Value2 > xlWorksheetGTInput.Cells[3][91].Value2)
                                        xlWorksheetGTInput.Cells[2][91].Value2 = xlWorksheetGTInput.Cells[3][91].Value2;

                                    xlWorksheetGTInput.Cells[2][92].Value2 = dt.Rows[0]["Active1VisionLouvreWidth"].ToString();


                                    if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][93].Value2 = "Offset";
                                    else if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][93].Value2 = "Central";

                                    xlWorksheetGTInput.Cells[2][94].Value2 = dt.Rows[0]["Active1VisionLouvreDistanceFromFloor"].ToString();
                                }

                                //vision / lourvre #1 PASSIVE
                                if (Convert.ToInt32(dt.Rows[0]["Passive1VisionGlassThickness"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][99].Value2 = "Vision";
                                    xlWorksheetGTInput.Cells[2][100].Value2 = dt.Rows[0]["Passive1VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][101].Value2 = dt.Rows[0]["Passive1VisionLouvreWidth"].ToString();

                                    if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][102].Value2 = "Offset";
                                    else if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][102].Value2 = "Central";

                                    // xlWorksheetGTInput.Cells[2][103].Value2 = dt.Rows[0]["Passive1VisionLouvreDistanceFromFloor"].ToString();


                                }
                                else if (Convert.ToInt32(dt.Rows[0]["Passive1VisionGlassThickness"].ToString()) == 0 &&
                                         Convert.ToInt32(dt.Rows[0]["Passive1VisionLouvreHeight"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][90].Value2 = "Louver";
                                    xlWorksheetGTInput.Cells[2][91].Value2 = dt.Rows[0]["Passive1VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][92].Value2 = dt.Rows[0]["Passive1VisionLouvreWidth"].ToString();


                                    if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][93].Value2 = "Offset";
                                    else if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][93].Value2 = "Central";

                                    xlWorksheetGTInput.Cells[2][94].Value2 = dt.Rows[0]["Passive1VisionLouvreDistanceFromFloor"].ToString();
                                }

                                //vision / louvre #2 ACTIVE
                                if (Convert.ToInt32(dt.Rows[0]["Active2VisionGlassThickness"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][108].Value2 = "Vision";
                                    xlWorksheetGTInput.Cells[2][109].Value2 = dt.Rows[0]["Active2VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][110].Value2 = dt.Rows[0]["Active2VisionLouvreWidth"].ToString();

                                    if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][111].Value2 = "Offset";
                                    else if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][111].Value2 = "Central";

                                    //xlWorksheetGTInput.Cells[2][112].Value2 = dt.Rows[0]["Active1VisionLouvreDistanceFromFloor"].ToString();


                                }
                                else if (Convert.ToInt32(dt.Rows[0]["Active2VisionGlassThickness"].ToString()) == 0 &&
                                         Convert.ToInt32(dt.Rows[0]["Active2VisionLouvreHeight"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][108].Value2 = "Louver";
                                    xlWorksheetGTInput.Cells[2][109].Value2 = dt.Rows[0]["Active2VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][110].Value2 = dt.Rows[0]["Active2VisionLouvreWidth"].ToString();


                                    if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][111].Value2 = "Offset";
                                    else if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][111].Value2 = "Central";

                                    //if this value is higher than the calculation to the right of it > set it to the calculation limit
                                    if (xlWorksheetGTInput.Cells[2][109].Value2 > xlWorksheetGTInput.Cells[3][91].Value2)
                                        xlWorksheetGTInput.Cells[2][109].Value2 = xlWorksheetGTInput.Cells[3][91].Value2;


                                    xlWorksheetGTInput.Cells[2][112].Value2 = dt.Rows[0]["Active2VisionLouvreDistanceFromFloor"].ToString();

                                }

                                //vision / louvre #2 Passive
                                if (Convert.ToInt32(dt.Rows[0]["Passive2VisionGlassThickness"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][108].Value2 = "Vision";
                                    xlWorksheetGTInput.Cells[2][109].Value2 = dt.Rows[0]["Passive2VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][110].Value2 = dt.Rows[0]["Passive2VisionLouvreWidth"].ToString();

                                    if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][111].Value2 = "Offset";
                                    else if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][111].Value2 = "Central";

                                    xlWorksheetGTInput.Cells[2][112].Value2 = dt.Rows[0]["Passive1VisionLouvreDistanceFromFloor"].ToString();


                                }
                                else if (Convert.ToInt32(dt.Rows[0]["Passive2VisionGlassThickness"].ToString()) == 0 &&
                                         Convert.ToInt32(dt.Rows[0]["Passive2VisionLouvreHeight"].ToString()) > 0) //need to be int
                                {
                                    xlWorksheetGTInput.Cells[2][117].Value2 = "Louver";
                                    xlWorksheetGTInput.Cells[2][118].Value2 = dt.Rows[0]["Passive2VisionLouvreHeight"].ToString();
                                    xlWorksheetGTInput.Cells[2][119].Value2 = dt.Rows[0]["Passive2VisionLouvreWidth"].ToString();


                                    if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][120].Value2 = "Offset";
                                    else if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][120].Value2 = "Central";

                                    xlWorksheetGTInput.Cells[2][121].Value2 = dt.Rows[0]["Passive2VisionLouvreDistanceFromFloor"].ToString();


                                    sql = "select stock_code FROM dbo.STOCK where description = '" + dt.Rows[0]["TopLockingBolt"].ToString() + "'";
                                    using (SqlCommand cmdAny1 = new SqlCommand(sql))
                                    {
                                        sql = "select GT_input_name FROM dbo.bridge_hardware where stock_code = '" + cmdAny1.ExecuteScalar().ToString() + "'";
                                        using (SqlCommand cmdAny1StockCode = new SqlCommand(sql, conn))
                                        {
                                            xlWorksheetGTInput.Cells[2][130].Value2 = cmdAny1StockCode.ExecuteScalar().ToString().Trim();
                                        }
                                    }
                                }


                                xlWorksheetGTInput.Cells[2][126].Value2 = dt.Rows[0]["KickPlateSide"].ToString();

                                //translate >> dont need MM
                                xlWorksheetGTInput.Cells[2][127].Value2 = Regex.Match(dt.Rows[0]["kickPlateType"].ToString(), @"\d+").Value;

                                xlWorksheetGTInput.Cells[2][128].Value2 = dt.Rows[0]["KickPlateLeaves"].ToString();

                            }

                            int max_bridge_id = 0;
                            sql = "select top 1 id from dbo.bridge_log where door_id = " + door_number + " order by id desc";
                            using (SqlCommand cmdBridgeLog = new SqlCommand(sql, conn))
                                max_bridge_id = (int)cmdBridgeLog.ExecuteScalar();

                            sql = "update dbo.bridge_log SET gt_input_success = -1 WHERE id = " + max_bridge_id.ToString();
                            using (SqlCommand cmdBridgeLog = new SqlCommand(sql, conn))
                                cmdBridgeLog.ExecuteNonQuery();

                            conn.Close();
                        }


                        // Save the excel file under the captured location from the SaveFileDialog
                        xlWorkbookGTInput.SaveAs(new_GT_input_location);//@"C:\temp\GTINPUT\GT INPUTaaa");
                        xlAppGTInput.DisplayAlerts = true;
                        xlWorkbookGTInput.Close(true, misValue, misValue);
                        xlAppGTInput.Quit();

                        // Open the newly saved excel file
                        if (File.Exists(new_GT_input_location))
                            System.Diagnostics.Process.Start(new_GT_input_location);


                    }
                    catch
                    {
                        //update dbo.bridge_log
                        int max_bridge_id = 0;

                        using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                        {
                            conn.Open();
                            string sql = "select top 1 id from dbo.bridge_log where door_id = " + door_number + " order by id desc";
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                                max_bridge_id = (int)cmd.ExecuteScalar();

                            sql = "update dbo.bridge_log SET gt_input_success = 0 WHERE id = " + max_bridge_id.ToString();
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                                cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
                else if (door_type.Contains("SG") || door_type == "Single" || door_type == "Double" || door_type == "Double Leaf & Half")
                {
                    //try
                    //{

                    GT_input_location += ".xlsm";
                    new_GT_input_location += ".xlsm";

                    //^^ we need to copy and move this file before editing --temp for now
                    System.IO.File.Copy(GT_input_location, new_GT_input_location, true); //true = overwrite

                    //edit the huw excel sheet thing
                    Microsoft.Office.Interop.Excel.Application xlAppGTInput = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook xlWorkbookGTInput = xlAppGTInput.Workbooks.Open(new_GT_input_location, 0, false, 5, "", "", false,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Microsoft.Office.Interop.Excel.Worksheet xlWorksheetGTInput = xlWorkbookGTInput.Sheets[1]; // assume it is the first sheet
                    Microsoft.Office.Interop.Excel.Range xlRangeCSV = xlWorksheetGTInput.UsedRange; // get the entire used range

                    //rename new_GT_input_location -- cant save it as the same name because of read/write issues
                    new_GT_input_location = @"C:\temp\GTINPUT\GT INPUT " + door_number + ".xlsm";


                    xlAppGTInput.DisplayAlerts = false;

                    using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                    {
                        conn.Open();

                        string sql = "select * " +
                            "FROM dbo.DWBridge dw " +
                            "left join dbo.door d on dw.SalesOrderNum = d.quote_number " +
                            "left join dbo.SALES_LEDGER s on d.customer_acc_ref = s.ACCOUNT_REF " +
                            "left join dbo.door_type dt on d.door_type_id = dt.id " +
                            "where d.id = " + door_number + " AND d.quote_number = '" + quote_number + "'";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            DataTable dt = new DataTable();

                            da.Fill(dt);

                            if (dt.Rows.Count == 0)
                            {
                                Console.WriteLine("There is no record in DWBridge :(");
                                Console.ReadLine();
                                return;
                            }

                            //[col][row] //column 2 = B
                            xlWorksheetGTInput.Cells[2][1].Value2 = door_number;
                            xlAppGTInput.CalculateUntilAsyncQueriesDone();
                            xlWorksheetGTInput.Cells[2][2].Value2 = quote_number.ToString();
                            xlWorksheetGTInput.Cells[2][3].Value2 = dt.Rows[0]["NAME"].ToString();

                            sql = "select top 1 left(forename,1) + left(surname,1) FROM dbo.bridge_log b " +
                                "left join [user_info].dbo.[user] u on b.staff_id = u.id " +
                                "where door_id = " + door_number + " order by b.id desc";

                            using (SqlCommand cmdProgrammer = new SqlCommand(sql, conn))
                            {
                                xlWorksheetGTInput.Cells[2][4].Value2 = cmdProgrammer.ExecuteScalar().ToString();
                            }

                            xlWorksheetGTInput.Cells[2][5].Value2 = dt.Rows[0]["quantity_same"].ToString();
                            xlWorksheetGTInput.Cells[2][6].Value2 = dt.Rows[0]["door_ref"].ToString();

                            if (dt.Rows[0]["double_y_n"].ToString() == "-1")
                                xlWorksheetGTInput.Cells[2][8].Value2 = "Double A+B-Skin";
                            else
                                xlWorksheetGTInput.Cells[2][8].Value2 = "Single A-Skin";

                            if (dt.Rows[0]["DoorStyle"].ToString() == "Double (Unequal Split)")
                                xlWorksheetGTInput.Cells[2][9].Value2 = "Yes";


                            if (dt.Rows[0]["RebateType"].ToString() == "Single" && dt.Rows[0]["FrameType"].ToString() == "Non Tagged")
                                xlWorksheetGTInput.Cells[2][12].Value2 = "Single Rebate";
                            else if (dt.Rows[0]["RebateType"].ToString() == "Single" && dt.Rows[0]["FrameType"].ToString() == "Tagged")
                                xlWorksheetGTInput.Cells[2][12].Value2 = "Single Rebate Tagged";
                            if (dt.Rows[0]["RebateType"].ToString() == "Double" && dt.Rows[0]["FrameType"].ToString() == "Non Tagged")
                                xlWorksheetGTInput.Cells[2][12].Value2 = "Single Rebate";
                            else if (dt.Rows[0]["RebateType"].ToString() == "Double" && dt.Rows[0]["FrameType"].ToString() == "Tagged")
                                xlWorksheetGTInput.Cells[2][12].Value2 = "Double Rebate Tagged";
                            else if (dt.Rows[0]["FrameType"].ToString() == "Wrap Around")
                                xlWorksheetGTInput.Cells[2][12].Value2 = "Wrap Hinged";

                            if (dt.Rows[0]["DoorStyle"].ToString().Contains("Double"))
                            {
                                if (dt.Rows[0]["NAME"].ToString() == "JODAN CONTRACTS LTD")
                                {
                                    xlWorksheetGTInput.Cells[2][13].Value2 = "Welded";
                                }
                                else
                                    xlWorksheetGTInput.Cells[2][13].Value2 = "Bolted";
                            }
                            else if (dt.Rows[0]["DoorStyle"].ToString().Contains("Single"))
                            {
                                xlWorksheetGTInput.Cells[2][13].Value2 = "Welded";
                            }

                            xlWorksheetGTInput.Cells[2][14].Value2 = "45";

                            xlWorksheetGTInput.Cells[2][15].Value2 = "Galv";

                            xlWorksheetGTInput.Cells[2][16].Value2 = dt.Rows[0]["LeafMatlThickness"].ToString();
                            xlWorksheetGTInput.Cells[2][17].Value2 = dt.Rows[0]["FrameThicknessType"].ToString();



                            //get the infil >> if its fire rated then its >>fire door (bs 476)<<
                            if (dt.Rows[0]["infill_id"].ToString() == "6")
                            {
                                xlWorksheetGTInput.Cells[2][18].Value2 = "Fire Door (BS 476)";
                            }
                            else
                            {
                                if (dt.Rows[0]["door_type_description"].ToString().Contains("SG"))
                                {
                                    xlWorksheetGTInput.Cells[2][18].Value2 = "SG Door";
                                }
                                else
                                {
                                    xlWorksheetGTInput.Cells[2][18].Value2 = "General Purpose Door";
                                }

                            }

                            if (dt.Rows[0]["infill"].ToString() == "Dufaylite") // needs testing
                            {
                                xlWorksheetGTInput.Cells[2][19].Value2 = "Dufalite";
                            }
                            else if (dt.Rows[0]["infill"].ToString() == "Timber")
                            {
                                xlWorksheetGTInput.Cells[2][19].Value2 = "Wood";
                            }
                            else
                                xlWorksheetGTInput.Cells[2][19].Value2 = dt.Rows[0]["infill"].ToString();



                            xlWorksheetGTInput.Cells[2][20].Value2 = dt.Rows[0]["SOW"].ToString();
                            xlWorksheetGTInput.Cells[2][21].Value2 = dt.Rows[0]["SOH"].ToString();

                            xlWorksheetGTInput.Cells[2][22].Value2 = dt.Rows[0]["JambDepth"].ToString();
                            xlWorksheetGTInput.Cells[2][23].Value2 = dt.Rows[0]["JambWidth"].ToString();

                            if (door_type.Contains("SG")) // needs testing
                                xlWorksheetGTInput.Cells[2][24].Value2 = "Standard 80 Kilo Dogbolt 76 X 101.6 x3 St.St. (screw & glue)";
                            else
                                xlWorksheetGTInput.Cells[2][24].Value2 = "Standard 80 Kilo Dogbolt 76 X 101.6 x3 St.St.";

                            if (door_type.Contains("SG")) // needs testing
                                xlWorksheetGTInput.Cells[2][25].Value2 = "Standard 80 Kilo Dogbolt 76 X 101.6 x3 St.St. (screw & glue)";
                            else
                                xlWorksheetGTInput.Cells[2][25].Value2 = "Standard 80 Kilo Dogbolt 76 X 101.6 x3 St.St.";


                            //absolutely no idea wtf this is doing
                            xlWorksheetGTInput.Cells[2][26].Value2 = xlWorksheetGTInput.Cells[12][18].Value2; //needs big testing


                            //need to translate this
                            if (dt.Rows[0]["CillType"].ToString() == "Aluminium 564a")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Standard Aluminium";
                            else if (dt.Rows[0]["CillType"].ToString() == "A Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill A type";
                            else if (dt.Rows[0]["CillType"].ToString() == "B Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill B type";
                            else if (dt.Rows[0]["CillType"].ToString() == "C Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill C type";
                            else if (dt.Rows[0]["CillType"].ToString() == "D Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill D type";
                            else if (dt.Rows[0]["CillType"].ToString() == "E Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill E type";
                            else if (dt.Rows[0]["CillType"].ToString() == "H Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill H type";
                            else if (dt.Rows[0]["CillType"].ToString() == "J Type")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Cill J type";
                            else if (dt.Rows[0]["CillType"].ToString() == "AM3 Cill")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "AM3X Stormguard Aluminium Cill";
                            else if (dt.Rows[0]["CillType"].ToString() == "Aluminium 544A")
                                xlWorksheetGTInput.Cells[2][27].Value2 = "544A Double Ramped Aluminium Cill";
                            else if (dt.Rows[0]["CillType"].ToString() == "Head Jamb Cill" &&
                                (xlWorksheetGTInput.Cells[4][12].Value2 == "1" || xlWorksheetGTInput.Cells[4][12].Value2 == "2" ||
                                    xlWorksheetGTInput.Cells[4][12].Value2 == "7" || xlWorksheetGTInput.Cells[4][12].Value2 == "8")) //BIG TEST
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Head Jamb (Jtype 1,2,7,8)";
                            else if (dt.Rows[0]["CillType"].ToString() == "Head Jamb Cill" &&
                                    (xlWorksheetGTInput.Cells[4][12].Value2 == "3" || xlWorksheetGTInput.Cells[4][12].Value2 == "4" ||
                                     xlWorksheetGTInput.Cells[4][12].Value2 == "5" || xlWorksheetGTInput.Cells[4][12].Value2 == "6")) //BIG TEST
                                xlWorksheetGTInput.Cells[2][27].Value2 = "Head Jamb (Jtype 3,4,5,6)";
                            else
                                xlWorksheetGTInput.Cells[2][27].Value2 = dt.Rows[0]["CillType"].ToString();


                            //  xlWorksheetGTInput.Cells[2][28].Value2 = dt.Rows[0]["fixingType"].ToString();

                            if (dt.Rows[0]["fixingType"].ToString() == "Visible Fixings") // needs testing
                                xlWorksheetGTInput.Cells[2][28].Value2 = "Visable Fixings";
                            else
                                xlWorksheetGTInput.Cells[2][28].Value2 = dt.Rows[0]["fixingType"].ToString();

                            //translate
                            if (dt.Rows[0]["hasJackingScrews"].ToString() == "Jacking Screws")
                                xlWorksheetGTInput.Cells[2][29].Value2 = "Yes";


                            xlWorksheetGTInput.Cells[2][30].Value2 = dt.Rows[0]["fixingTo"].ToString();


                            //translate
                            if (dt.Rows[0]["Handing"].ToString().Contains("L"))
                                xlWorksheetGTInput.Cells[2][31].Value2 = "Left Hand";
                            else if (dt.Rows[0]["Handing"].ToString().Contains("R"))
                                xlWorksheetGTInput.Cells[2][31].Value2 = "Right Hand";

                            xlWorksheetGTInput.Cells[2][32].Value2 = dt.Rows[0]["openingDirection"].ToString();



                            //center locks
                            sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["CentreLockStockCode"].ToString() + "'";
                            using (SqlCommand cmdCenterLock = new SqlCommand(sql, conn))
                            {
                                var temp = cmdCenterLock.ExecuteScalar();
                                if (temp != null)
                                {
                                    xlWorksheetGTInput.Cells[2][34].Value2 = cmdCenterLock.ExecuteScalar().ToString();
                                }
                            }

                            if (dt.Rows[0]["CentreLeverInside"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][36].Value2 = "Lever-Rose Fixed Assa 640 Un-sprung St.St.";

                            if (dt.Rows[0]["centrelockheight"].ToString().Length > 0)
                                xlWorksheetGTInput.Cells[4][36].Value2 = dt.Rows[0]["centrelockheight"].ToString(); //special box

                            if (dt.Rows[0]["CentreLeverOutside"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][37].Value2 = "Lever-Rose Fixed Assa 640 Un-sprung St.St.";

                            if (dt.Rows[0]["CentreLockingInside"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][38].Value2 = "Yes";

                            if (dt.Rows[0]["CentreLockingOutside"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][39].Value2 = "Yes";


                            if (dt.Rows[0]["CentreInsideEscutcheonName"].ToString().Contains(" KEY / KEY ") && dt.Rows[0]["CentreOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                xlWorksheetGTInput.Cells[2][40].Value2 = "Assa Full 31MM / 31MM SCP (Key o/s, Key i/s)";
                            else if (dt.Rows[0]["CentreInsideEscutcheonName"].ToString().Contains("KEY / TURN") && dt.Rows[0]["CentreOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                xlWorksheetGTInput.Cells[2][40].Value2 = "Assa Half 31MM SCP (Key o/s, Thumbturn i/s)";
                            else if (dt.Rows[0]["CentreInsideEscutcheonName"].ToString().Contains("KEY / KEY ") && dt.Rows[0]["CentreOutsideEscutcheonName"].ToString().Contains("KEY / TURN"))
                                xlWorksheetGTInput.Cells[2][40].Value2 = "Assa Half 31MM SCP (Thumbturn o/s, Key i/s)";

                            //door loop
                            if (dt.Rows[0]["DoorLoopType"].ToString().Contains("DL8"))
                            {
                                xlWorksheetGTInput.Cells[2][42].Value2 = "Abloy DL8 Surface Mounted Door Loop"; //same for the below??
                                xlWorksheetGTInput.Cells[2][43].Value2 = "Active Leaf";
                            }
                            else if (dt.Rows[0]["DoorLoopType"].ToString().Contains("EA280"))
                                xlWorksheetGTInput.Cells[2][42].Value2 = "Abloy EA280 Concealed Door Loop";


                            //top lock
                            sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["TopLockStockCode"].ToString() + "'";
                            using (SqlCommand cmdTopLock = new SqlCommand(sql, conn))
                            {
                                var temp = cmdTopLock.ExecuteScalar();
                                if (temp != null)
                                {
                                    xlWorksheetGTInput.Cells[2][45].Value2 = cmdTopLock.ExecuteScalar().ToString();

                                    if (dt.Rows[0]["ToplockingInside"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][47].Value2 = "Yes";
                                    if (dt.Rows[0]["TopLockingOutside"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][48].Value2 = "Yes";



                                    if (dt.Rows[0]["TopInsideEscutcheonName"].ToString().Contains(" KEY / KEY ") && dt.Rows[0]["TopOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                        xlWorksheetGTInput.Cells[2][49].Value2 = "Assa Full 31MM / 31MM SCP (Key o/s, Key i/s)";
                                    else if (dt.Rows[0]["TopInsideEscutcheonName"].ToString().Contains("KEY / TURN") && dt.Rows[0]["TopOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                        xlWorksheetGTInput.Cells[2][49].Value2 = "Assa Half 31MM SCP (Key o/s, Thumbturn i/s)";
                                    else if (dt.Rows[0]["TopInsideEscutcheonName"].ToString().Contains("KEY / KEY ") && dt.Rows[0]["TopOutsideEscutcheonName"].ToString().Contains("KEY / TURN"))
                                        xlWorksheetGTInput.Cells[2][49].Value2 = "Assa Half 31MM SCP (Thumbturn o/s, Key i/s)";

                                }

                            }

                            //bot lock

                            sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["TopLockStockCode"].ToString() + "'";
                            using (SqlCommand cmdTopLock = new SqlCommand(sql, conn))
                            {
                                var temp = cmdTopLock.ExecuteScalar();
                                if (temp != null)
                                {
                                    xlWorksheetGTInput.Cells[2][51].Value2 = cmdTopLock.ExecuteScalar().ToString();

                                    if (dt.Rows[0]["ToplockingInside"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][53].Value2 = "Yes";
                                    if (dt.Rows[0]["TopLockingOutside"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][54].Value2 = "Yes";
                                    //55 cylinderrrr
                                    if (dt.Rows[0]["BotInsideEscutcheonName"].ToString().Contains(" KEY / KEY ") && dt.Rows[0]["BottomOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                        xlWorksheetGTInput.Cells[2][55].Value2 = "Assa Full 31MM / 31MM SCP (Key o/s, Key i/s)";
                                    else if (dt.Rows[0]["BottomInsideEscutcheonName"].ToString().Contains("KEY / TURN") && dt.Rows[0]["BottomOutsideEscutcheonName"].ToString().Contains("KEY / KEY "))
                                        xlWorksheetGTInput.Cells[2][55].Value2 = "Assa Half 31MM SCP (Key o/s, Thumbturn i/s)";
                                    else if (dt.Rows[0]["BottomInsideEscutcheonName"].ToString().Contains("KEY / KEY ") && dt.Rows[0]["BottomOutsideEscutcheonName"].ToString().Contains("KEY / TURN"))
                                        xlWorksheetGTInput.Cells[2][55].Value2 = "Assa Half 31MM SCP (Thumbturn o/s, Key i/s)";

                                }
                            }


                            //panics
                            sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["PanicDeviceStockCode"].ToString() + "'";
                            using (SqlCommand cmdPanic = new SqlCommand(sql, conn))
                            {
                                var temp = cmdPanic.ExecuteScalar();
                                if (temp != null)
                                {
                                    xlWorksheetGTInput.Cells[2][57].Value2 = cmdPanic.ExecuteScalar().ToString();

                                    xlWorksheetGTInput.Cells[2][58].Value2 = xlWorksheetGTInput.Cells[2][59].Value2;


                                    sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["OADStockCode"].ToString() + "'";
                                    using (SqlCommand cmdOAD = new SqlCommand(sql, conn))
                                    {
                                        var temp2 = cmdOAD.ExecuteScalar();
                                        if (temp2 != null)
                                        {
                                            xlWorksheetGTInput.Cells[2][60].Value2 = cmdOAD.ExecuteScalar().ToString();
                                        }
                                    }

                                }
                            }


                            //pushplate stuffs
                            if (dt.Rows[0]["PullHandleCode"].ToString() == "286")
                            {
                                //if there is a pushplate then we use > "Pull Handle 19 x 300 Rose Mounted St.St. & 330mm x 76mm Push Plate"
                                if (dt.Rows[0]["PushPlateType"].ToString().Length > 0)
                                    xlWorksheetGTInput.Cells[2][62].Value2 = "Pull Handle 19 x 300 Rose Mounted St.St. & 330mm x 76mm Push Plate";
                                else
                                    xlWorksheetGTInput.Cells[2][62].Value2 = "Pull Handle 19 x 300 Rose Mounted St.St.";
                            }

                            //xlWorksheetGTInput.Cells[2][63].Value2 = dt.Rows[0]["pushPlateSide"].ToString(); //translate)
                            if (dt.Rows[0]["pushPlateSide"].ToString() == "Pull Side")
                                xlWorksheetGTInput.Cells[2][63].Value2 = "Pullside";
                            else if (dt.Rows[0]["pushPlateSide"].ToString() == "Push Side")
                                xlWorksheetGTInput.Cells[2][63].Value2 = "Pushside";
                            else if (dt.Rows[0]["pushPlateSide"].ToString() == "Both Side")
                                xlWorksheetGTInput.Cells[2][63].Value2 = "Both sides";

                            if (dt.Rows[0]["pushPlateLeaves"].ToString() == "Active")
                                xlWorksheetGTInput.Cells[2][64].Value2 = "1st Leaf";
                            else if (dt.Rows[0]["pushPlateLeaves"].ToString() == "Passive")
                                xlWorksheetGTInput.Cells[2][64].Value2 = "2nd Leaf";
                            else if (dt.Rows[0]["pushPlateLeaves"].ToString() == "Active/Passive")
                                xlWorksheetGTInput.Cells[2][64].Value2 = "Both Leafs";




                            //Closers

                            sql = "SELECT GT_input_name FROM dbo.bridge_hardware WHERE stock_code = '" + dt.Rows[0]["CloserStockCode"].ToString() + "'";
                            using (SqlCommand cmdCloser = new SqlCommand(sql, conn))
                            {
                                var temp = cmdCloser.ExecuteScalar();
                                if (temp != null)
                                {
                                    xlWorksheetGTInput.Cells[2][67].Value2 = cmdCloser.ExecuteScalar().ToString();

                                    xlWorksheetGTInput.Cells[2][68].Value2 = dt.Rows[0]["closerPullside"].ToString(); //closerPullside = 1 then pull side closerpushside = 1 then push
                                    if (dt.Rows[0]["closerPullside"].ToString() == "1" && dt.Rows[0]["closerpushside"].ToString() == "0")
                                        xlWorksheetGTInput.Cells[2][68].Value2 = "Pullside";
                                    else if (dt.Rows[0]["closerPullside"].ToString() == "0" && dt.Rows[0]["closerpushside"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][68].Value2 = "Pushside";

                                    if (dt.Rows[0]["closerOnActive"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][69].Value2 = "Yes";

                                    if (dt.Rows[0]["closerOnPassive"].ToString() == "1")
                                        xlWorksheetGTInput.Cells[2][71].Value2 = "Yes";
                                }
                            }




                            if (dt.Rows[0]["StayLeaves"].ToString() == "Active")
                                xlWorksheetGTInput.Cells[2][64].Value2 = "1st Leaf";
                            else if (dt.Rows[0]["StayLeaves"].ToString() == "Passive")
                                xlWorksheetGTInput.Cells[2][64].Value2 = "2nd Leaf";
                            else if (dt.Rows[0]["StayLeaves"].ToString() == "Active/Passive")
                                xlWorksheetGTInput.Cells[2][64].Value2 = "Both Leafs";

                            //end of closers




                            //stay
                            sql = "SELECT RTRIM(GT_input_name) FROM dbo.bridge_hardware bh " +
                                "left join dbo.stock s on bh.stock_code = s.stock_code WHERE s.description = '" + dt.Rows[0]["StayType"].ToString() + "'";
                            using (SqlCommand cmdStay = new SqlCommand(sql, conn))
                            {
                                if (string.IsNullOrEmpty(dt.Rows[0]["StayType"].ToString()))
                                { }
                                else
                                    xlWorksheetGTInput.Cells[2][74].Value2 = cmdStay.ExecuteScalar().ToString().Trim();
                            }

                            //leaf selector
                            if (dt.Rows[0]["LeafSelectorType"].ToString().Contains("MK2 SELECTOR EXTENDED  CATCH 152 ARM SAA"))
                                xlWorksheetGTInput.Cells[2][77].Value2 = " c/w Extended Catch & Arm SAA (Wedge)";


                            //letterbox leaf 
                            if (dt.Rows[0]["LetterBoxActive"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][79].Value2 = "10\" Aluminium"; //not sure this is going to add the " into the correct spot 

                            if (dt.Rows[0]["LetterBoxPassive"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][79].Value2 = "10\" Aluminium"; //not sure this is going to add the " into the correct spot 

                            if (dt.Rows[0]["LetterBoxHeight"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][79].Value2 = "High";
                            else if (dt.Rows[0]["LetterBoxHeight"].ToString() == "2")
                                xlWorksheetGTInput.Cells[2][79].Value2 = "Low";


                            if (dt.Rows[0]["SpyHoleType"].ToString().Length > 0)
                                xlWorksheetGTInput.Cells[2][83].Value2 = "Zero 200 UL Door Viewer (Fire Rated)";


                            //round visions
                            if (dt.Rows[0]["RoundVisionNumberActive"].ToString() == "7")
                            {
                                xlWorksheetGTInput.Cells[2][85].Value2 = "Zero SP250";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberActive"].ToString() == "8")
                            {
                                xlWorksheetGTInput.Cells[2][85].Value2 = "Zero SP350";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberActive"].ToString() == "9")
                            {
                                xlWorksheetGTInput.Cells[2][85].Value2 = "Zero SP450";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberActive"].ToString() == "11")
                            {
                                xlWorksheetGTInput.Cells[2][85].Value2 = "Zero SP450 (Double Glazed)";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberActive"].ToString() == "10")
                            {
                                xlWorksheetGTInput.Cells[2][85].Value2 = "Zero SP550";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }

                            if (dt.Rows[0]["RoundVisionNumberPassive"].ToString() == "7")
                            {
                                xlWorksheetGTInput.Cells[2][86].Value2 = "Zero SP250";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberPassive"].ToString() == "8")
                            {
                                xlWorksheetGTInput.Cells[2][86].Value2 = "Zero SP350";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberPassive"].ToString() == "9")
                            {
                                xlWorksheetGTInput.Cells[2][86].Value2 = "Zero SP450";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberPassive"].ToString() == "11")
                            {
                                xlWorksheetGTInput.Cells[2][86].Value2 = "Zero SP450 (Double Glazed)";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }
                            else if (dt.Rows[0]["RoundVisionNumberPassive"].ToString() == "10")
                            {
                                xlWorksheetGTInput.Cells[2][86].Value2 = "Zero SP550";
                                xlWorksheetGTInput.Cells[2][88].Value2 = "Clear Laminate";
                            }

                            xlWorksheetGTInput.Cells[2][87].Value2 = dt.Rows[0]["RoundVisionHeight"].ToString();
                            //////////////////////////////////////////////////////////////////////////

                            //vision / lourvre #1 ACTIVE
                            if (Convert.ToInt32(dt.Rows[0]["Active1VisionGlassThickness"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][90].Value2 = "Vision";
                                xlWorksheetGTInput.Cells[2][91].Value2 = dt.Rows[0]["Active1VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][92].Value2 = dt.Rows[0]["Active1VisionLouvreWidth"].ToString();

                                if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][93].Value2 = "Offset";
                                else if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][93].Value2 = "Central";

                                xlWorksheetGTInput.Cells[2][94].Value2 = dt.Rows[0]["Active1VisionLouvreDistanceFromFloor"].ToString();


                            }
                            else if (Convert.ToInt32(dt.Rows[0]["Active1VisionGlassThickness"].ToString()) == 0 &&
                                     Convert.ToInt32(dt.Rows[0]["Active1VisionLouvreHeight"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][90].Value2 = "Louver";
                                xlWorksheetGTInput.Cells[2][91].Value2 = dt.Rows[0]["Active1VisionLouvreHeight"].ToString();

                                //if this value is higher than the calculation to the right of it > set it to the calculation limit
                                if (xlWorksheetGTInput.Cells[2][91].Value2 > xlWorksheetGTInput.Cells[3][91].Value2)
                                    xlWorksheetGTInput.Cells[2][91].Value2 = xlWorksheetGTInput.Cells[3][91].Value2;

                                xlWorksheetGTInput.Cells[2][92].Value2 = dt.Rows[0]["Active1VisionLouvreWidth"].ToString();


                                if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][93].Value2 = "Offset";
                                else if (dt.Rows[0]["Active1VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][93].Value2 = "Central";

                                xlWorksheetGTInput.Cells[2][94].Value2 = dt.Rows[0]["Active1VisionLouvreDistanceFromFloor"].ToString();
                            }

                            //vision / lourvre #1 PASSIVE
                            if (Convert.ToInt32(dt.Rows[0]["Passive1VisionGlassThickness"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][99].Value2 = "Vision";
                                xlWorksheetGTInput.Cells[2][100].Value2 = dt.Rows[0]["Passive1VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][101].Value2 = dt.Rows[0]["Passive1VisionLouvreWidth"].ToString();

                                if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][102].Value2 = "Offset";
                                else if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][102].Value2 = "Central";

                                // xlWorksheetGTInput.Cells[2][103].Value2 = dt.Rows[0]["Passive1VisionLouvreDistanceFromFloor"].ToString();


                            }
                            else if (Convert.ToInt32(dt.Rows[0]["Passive1VisionGlassThickness"].ToString()) == 0 &&
                                     Convert.ToInt32(dt.Rows[0]["Passive1VisionLouvreHeight"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][90].Value2 = "Louver";
                                xlWorksheetGTInput.Cells[2][91].Value2 = dt.Rows[0]["Passive1VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][92].Value2 = dt.Rows[0]["Passive1VisionLouvreWidth"].ToString();


                                if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][93].Value2 = "Offset";
                                else if (dt.Rows[0]["Passive1VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][93].Value2 = "Central";

                                xlWorksheetGTInput.Cells[2][94].Value2 = dt.Rows[0]["Passive1VisionLouvreDistanceFromFloor"].ToString();
                            }

                            //vision / louvre #2 ACTIVE
                            if (Convert.ToInt32(dt.Rows[0]["Active2VisionGlassThickness"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][108].Value2 = "Vision";
                                xlWorksheetGTInput.Cells[2][109].Value2 = dt.Rows[0]["Active2VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][110].Value2 = dt.Rows[0]["Active2VisionLouvreWidth"].ToString();

                                if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][111].Value2 = "Offset";
                                else if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][111].Value2 = "Central";

                                //xlWorksheetGTInput.Cells[2][112].Value2 = dt.Rows[0]["Active1VisionLouvreDistanceFromFloor"].ToString();


                            }
                            else if (Convert.ToInt32(dt.Rows[0]["Active2VisionGlassThickness"].ToString()) == 0 &&
                                     Convert.ToInt32(dt.Rows[0]["Active2VisionLouvreHeight"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][108].Value2 = "Louver";
                                xlWorksheetGTInput.Cells[2][109].Value2 = dt.Rows[0]["Active2VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][110].Value2 = dt.Rows[0]["Active2VisionLouvreWidth"].ToString();


                                if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][111].Value2 = "Offset";
                                else if (dt.Rows[0]["Active2VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][111].Value2 = "Central";

                                //if this value is higher than the calculation to the right of it > set it to the calculation limit
                                if (xlWorksheetGTInput.Cells[2][109].Value2 > xlWorksheetGTInput.Cells[3][91].Value2)
                                    xlWorksheetGTInput.Cells[2][109].Value2 = xlWorksheetGTInput.Cells[3][91].Value2;


                                xlWorksheetGTInput.Cells[2][112].Value2 = dt.Rows[0]["Active2VisionLouvreDistanceFromFloor"].ToString();

                            }

                            //vision / louvre #2 Passive
                            if (Convert.ToInt32(dt.Rows[0]["Passive2VisionGlassThickness"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][108].Value2 = "Vision";
                                xlWorksheetGTInput.Cells[2][109].Value2 = dt.Rows[0]["Passive2VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][110].Value2 = dt.Rows[0]["Passive2VisionLouvreWidth"].ToString();

                                if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][111].Value2 = "Offset";
                                else if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][111].Value2 = "Central";

                                xlWorksheetGTInput.Cells[2][112].Value2 = dt.Rows[0]["Passive1VisionLouvreDistanceFromFloor"].ToString();


                            }
                            else if (Convert.ToInt32(dt.Rows[0]["Passive2VisionGlassThickness"].ToString()) == 0 &&
                                     Convert.ToInt32(dt.Rows[0]["Passive2VisionLouvreHeight"].ToString()) > 0) //need to be int
                            {
                                xlWorksheetGTInput.Cells[2][117].Value2 = "Louver";
                                xlWorksheetGTInput.Cells[2][118].Value2 = dt.Rows[0]["Passive2VisionLouvreHeight"].ToString();
                                xlWorksheetGTInput.Cells[2][119].Value2 = dt.Rows[0]["Passive2VisionLouvreWidth"].ToString();


                                if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "1")
                                    xlWorksheetGTInput.Cells[2][120].Value2 = "Offset";
                                else if (dt.Rows[0]["Passive2VisionLouvreSetback"].ToString() == "0")
                                    xlWorksheetGTInput.Cells[2][120].Value2 = "Central";

                                xlWorksheetGTInput.Cells[2][121].Value2 = dt.Rows[0]["Passive2VisionLouvreDistanceFromFloor"].ToString();


                                sql = "select stock_code FROM dbo.STOCK where description = '" + dt.Rows[0]["TopLockingBolt"].ToString() + "'";
                                using (SqlCommand cmdAny1 = new SqlCommand(sql))
                                {
                                    sql = "select GT_input_name FROM dbo.bridge_hardware where stock_code = '" + cmdAny1.ExecuteScalar().ToString() + "'";
                                    using (SqlCommand cmdAny1StockCode = new SqlCommand(sql, conn))
                                    {
                                        xlWorksheetGTInput.Cells[2][130].Value2 = cmdAny1StockCode.ExecuteScalar().ToString().Trim();
                                    }
                                }
                            }


                            xlWorksheetGTInput.Cells[2][126].Value2 = dt.Rows[0]["KickPlateSide"].ToString();

                            //translate >> dont need MM
                            xlWorksheetGTInput.Cells[2][127].Value2 = Regex.Match(dt.Rows[0]["kickPlateType"].ToString(), @"\d+").Value;

                            xlWorksheetGTInput.Cells[2][128].Value2 = dt.Rows[0]["KickPlateLeaves"].ToString();


                            //signage 
                            if (dt.Rows[0]["SignageDescription"].ToString().Contains("FIRE ESCAPE KEEP CLEAR"))
                                xlWorksheetGTInput.Cells[2][141].Value2 = "Fire escape keep clear 76 dia";
                            else if (dt.Rows[0]["SignageDescription"].ToString().Contains("FIRE DOOR KEEP SHUT"))
                                xlWorksheetGTInput.Cells[2][141].Value2 = "Fire door keep shut 76 dia";
                            else if (dt.Rows[0]["SignageDescription"].ToString().Contains("FIRE EXIT KEEP CLEAR"))
                                xlWorksheetGTInput.Cells[2][141].Value2 = "Fire exit keep clear 76 dia";
                            else if (dt.Rows[0]["SignageDescription"].ToString().Contains("FIRE DOOR KEEP CLEAR"))
                                xlWorksheetGTInput.Cells[2][141].Value2 = "Fire door keep clear 76 dia";

                            if (dt.Rows[0]["SignageSide"].ToString() == "Push Side")
                                xlWorksheetGTInput.Cells[2][142].Value2 = "Pushside";
                            else if (dt.Rows[0]["SignageSide"].ToString() == "Pull Side")
                                xlWorksheetGTInput.Cells[2][142].Value2 = "Pullside";
                            else if (dt.Rows[0]["SignageSide"].ToString() == "Both Sides")
                                xlWorksheetGTInput.Cells[2][142].Value2 = "Both sides";
                            else if (dt.Rows[0]["SignageSide"].ToString() == "Both Side")
                                xlWorksheetGTInput.Cells[2][142].Value2 = "Both sides";


                            if (dt.Rows[0]["SignageLeaf"].ToString() == "Active")
                                xlWorksheetGTInput.Cells[2][143].Value2 = "1st Leaf";
                            else if (dt.Rows[0]["SignageLeaf"].ToString() == "Passive")
                                xlWorksheetGTInput.Cells[2][143].Value2 = "2nd Leaf";
                            else if (dt.Rows[0]["SignageLeaf"].ToString() == "Active/Passive")
                                xlWorksheetGTInput.Cells[2][143].Value2 = "Both Leafs";

                            xlWorksheetGTInput.Cells[2][144].Value2 = dt.Rows[0]["SignageDistanceFromBottom"].ToString();
                            xlWorksheetGTInput.Cells[2][145].Value2 = dt.Rows[0]["SignageBackset"].ToString();


                            //panels 1 , 2 and 3
                            int panel_count = 0;

                            if (dt.Rows[0]["Panel1RemFix"].ToString() == "1")
                            {
                                xlWorksheetGTInput.Cells[2][155].Value2 = "Fixed";
                                panel_count++;
                            }
                            else if (dt.Rows[0]["Panel1RemFix"].ToString() == "2")
                            {
                                xlWorksheetGTInput.Cells[2][155].Value2 = "Removable";
                                panel_count++;
                            }

                            if (dt.Rows[0]["Panel1Type"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][156].Value2 = "Overpanel";
                            else if (dt.Rows[0]["Panel1Type"].ToString() == "2")
                                xlWorksheetGTInput.Cells[2][156].Value2 = "Side Panel LHS pullside";
                            else if (dt.Rows[0]["Panel1Type"].ToString() == "3")
                                xlWorksheetGTInput.Cells[2][156].Value2 = "Side Panel RHS pullside";
                            else if (dt.Rows[0]["Panel1Type"].ToString() == "4")
                                xlWorksheetGTInput.Cells[2][156].Value2 = "Hinged overpanel";

                            xlWorksheetGTInput.Cells[2][157].Value2 = dt.Rows[0]["Panel1Width"].ToString();
                            xlWorksheetGTInput.Cells[2][158].Value2 = dt.Rows[0]["Panel1Height"].ToString();

                            if (dt.Rows[0]["Panel2RemFix"].ToString() == "1")
                            {
                                xlWorksheetGTInput.Cells[2][161].Value2 = "Fixed";
                                panel_count++;
                            }
                            else if (dt.Rows[0]["Panel2RemFix"].ToString() == "2")
                            {
                                xlWorksheetGTInput.Cells[2][161].Value2 = "Removable";
                                panel_count++;
                            }

                            if (dt.Rows[0]["Panel2Type"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][162].Value2 = "Overpanel";
                            else if (dt.Rows[0]["Panel2Type"].ToString() == "2")
                                xlWorksheetGTInput.Cells[2][162].Value2 = "Side Panel LHS pullside";
                            else if (dt.Rows[0]["Panel2Type"].ToString() == "3")
                                xlWorksheetGTInput.Cells[2][162].Value2 = "Side Panel RHS pullside";
                            else if (dt.Rows[0]["Panel2Type"].ToString() == "4")
                                xlWorksheetGTInput.Cells[2][162].Value2 = "Hinged overpanel";

                            xlWorksheetGTInput.Cells[2][163].Value2 = dt.Rows[0]["Panel2Width"].ToString();
                            xlWorksheetGTInput.Cells[2][164].Value2 = dt.Rows[0]["Panel2Height"].ToString();


                            if (dt.Rows[0]["Panel3RemFix"].ToString() == "1")
                            {
                                xlWorksheetGTInput.Cells[2][167].Value2 = "Fixed";
                                panel_count++;

                            }
                            else if (dt.Rows[0]["Panel3RemFix"].ToString() == "2")
                            {
                                xlWorksheetGTInput.Cells[2][167].Value2 = "Removable";
                                panel_count++;
                            }

                            if (dt.Rows[0]["Panel3Type"].ToString() == "1")
                                xlWorksheetGTInput.Cells[2][168].Value2 = "Overpanel";
                            else if (dt.Rows[0]["Panel3Type"].ToString() == "2")
                                xlWorksheetGTInput.Cells[2][168].Value2 = "Side Panel LHS pullside";
                            else if (dt.Rows[0]["Panel3Type"].ToString() == "3")
                                xlWorksheetGTInput.Cells[2][168].Value2 = "Side Panel RHS pullside";
                            else if (dt.Rows[0]["Panel3Type"].ToString() == "4")
                                xlWorksheetGTInput.Cells[2][168].Value2 = "Hinged overpanel";

                            xlWorksheetGTInput.Cells[2][169].Value2 = dt.Rows[0]["Panel3Width"].ToString();
                            xlWorksheetGTInput.Cells[2][170].Value2 = dt.Rows[0]["Panel3Height"].ToString();

                            //add the counted number of panels being used
                            if (panel_count > 0)
                                xlWorksheetGTInput.Cells[2][153].Value2 = panel_count;


                        }

                        int max_bridge_id = 0;
                        sql = "select top 1 id from dbo.bridge_log where door_id = " + door_number + " order by id desc";
                        using (SqlCommand cmdBridgeLog = new SqlCommand(sql, conn))
                            max_bridge_id = (int)cmdBridgeLog.ExecuteScalar();

                        sql = "update dbo.bridge_log SET gt_input_success = -1 WHERE id = " + max_bridge_id.ToString();
                        using (SqlCommand cmdBridgeLog = new SqlCommand(sql, conn))
                            cmdBridgeLog.ExecuteNonQuery();

                        conn.Close();
                    }


                    // Save the excel file under the captured location from the SaveFileDialog
                    xlWorkbookGTInput.SaveAs(new_GT_input_location);//@"C:\temp\GTINPUT\GT INPUTaaa");
                    xlAppGTInput.DisplayAlerts = true;
                    xlWorkbookGTInput.Close(true, misValue, misValue);
                    xlAppGTInput.Quit();

                    // Open the newly saved excel file
                    if (File.Exists(new_GT_input_location))
                        System.Diagnostics.Process.Start(new_GT_input_location);


                    //}
                    //catch
                    //{
                    //    //update dbo.bridge_log
                    //    int max_bridge_id = 0;

                    //    using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
                    //    {
                    //        conn.Open();
                    //        string sql = "select top 1 id from dbo.bridge_log where door_id = " + door_number + " order by id desc";
                    //        using (SqlCommand cmd = new SqlCommand(sql, conn))
                    //            max_bridge_id = (int)cmd.ExecuteScalar();

                    //        sql = "update dbo.bridge_log SET gt_input_success = 0 WHERE id = " + max_bridge_id.ToString();
                    //        using (SqlCommand cmd = new SqlCommand(sql, conn))
                    //            cmd.ExecuteNonQuery();
                    //        conn.Close();
                    //    }
                    //}

                }

            }
        }

        static void lineChanger(string newNumber, string fileName, int lineOfNumber)
        {
            string[] arrLine = File.ReadAllLines(fileName);
            arrLine[lineOfNumber - 1] = newNumber;
            File.WriteAllLines(fileName, arrLine);
        }


    }
}
