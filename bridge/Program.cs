using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bridge
{
    class Program
    {
        static void Main(string[] args)
        {
            string door_number = "";//"670973"; //params
            string quote_number = "";//"60870"; //same
            door_number = args[0];
            quote_number = args[1];
           
            string file = @"\\designsvr1\SOLIDWORKS\DWDevelopment\Specifications\ " + quote_number  + @"\documents" + "DataOutput " + quote_number + ".DO"; //location
            string test = File.ReadAllText(file);
            test = test.Replace(quote_number, door_number);
            File.WriteAllText(file, test);
            int line_number = 224; //this is ALWAYS the beginning
            //vv will change
            for (int i = 0; i < door_number.Length; i++)
            {
                string singleDigit = door_number.Substring(i, 1);
                lineChanger(singleDigit, file, line_number);
                line_number = line_number + 1;
            }
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
