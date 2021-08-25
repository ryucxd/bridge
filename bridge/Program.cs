using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft;
using System.Data.SqlClient;

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
            string newFile = @"\\designsvr1\apps\Door Master\" + door_number + ".DO";
            string packingFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\Packing List " + rev_number + ".xlsx"; //should be the default file path for the session for everyone
            string engineerFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\Engineers Notes word " +  rev_number + ".docx";
            string newPackingLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\Packing List " + door_number + ".xlsx";
            string newEngineerLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\Engineer Notes " + door_number + ".docx";


            System.IO.Directory.CreateDirectory(@"\\designsvr1\apps\bridge_jobcard\" + door_number);
            //string fileName = "DataOutput " + quote_number + "- Door Designer.DO";

            //^^ we need to copy and move this file before editing 
            System.IO.File.Copy(startFile, newFile, true); //true = overwrite

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

            //also edit the packing list 

            //at some point we are going to move this excel sheet to another directory too

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(packingFile);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange; // get the entire used range

            if (File.Exists(newPackingLocation))
                File.Delete(newPackingLocation);


            xlWorksheet.Cells[5][7].Value2 = door_number.ToString();
            xlWorksheet.SaveAs(newPackingLocation);
            xlWorkbook.Close(true); //close the excel sheet
            xlApp.Quit(); //close everything excel related so that theres no errors when the door program tries to connect 
            if (File.Exists(newEngineerLocation))
                File.Delete(newEngineerLocation);
            File.Copy(engineerFile, newEngineerLocation,true); // also move this one over with a new name


            //check if there is a entry in dbo.door_program
            string sql = "select door_id FROM dbo.door_program WHERE door_id = " + door_number;
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
                    cmd.ExecuteNonQuery();
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
