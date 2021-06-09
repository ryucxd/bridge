using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft;

namespace bridge
{
    class Program
    {
        static void Main(string[] args)
        {
            string door_number = "67097"; //params
            string quote_number = "61586-1-1";//"60870"; //same 
            string rev_number =  quote_number + "- Rev " + quote_number.Substring(quote_number.Length - 1);
            //door_number = args[0];
            //quote_number = args[1];

            string startFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\" + "DataOutput " + quote_number + "- Door Designer.DO"; ;//location
            string newFile = @"\\designsvr1\apps\Door Master\" + door_number + ".DO";
            string packingFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\Packing List " + rev_number + ".xlsx"; //should be the default file path for the session for everyone
            string engineerFile = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\" + quote_number + @"\documents\Engineers Notes " + rev_number + ".xlsx";
            string newPackingLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\Packing List " + door_number + ".xlsx";
            string newEngineerLocation = @"\\designsvr1\apps\bridge_jobcard\" + door_number + @"\Engineer Notes " + door_number + ".xlsx";

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

            xlWorksheet.Cells[5][7].Value2 = door_number.ToString();
            xlWorksheet.SaveAs(newPackingLocation);
            xlWorkbook.Close(true); //close the excel sheet
            xlApp.Quit(); //close everything excel related so that theres no errors when the door program tries to connect 

            File.Copy(engineerFile, newEngineerLocation); // also move this one over with a new name







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
