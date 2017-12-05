using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Aspose.Cells;
using System.IO;
using System.Windows.Forms;

namespace KeywordDrivenTest
{
    public static class LoadExcelFile
    { 


        public static string[,] ImportSheet(string fileName)
        {                                    
            // Creating a file stream containing the Excel file to be opened
            FileStream fstream = new FileStream(fileName, FileMode.Open,FileAccess.Read,FileShare.None);

            //Instantiating a Workbook object
            //Opening the Excel file through the file stream
            Workbook workbook = new Workbook(fstream);

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.Worksheets[0];

            //close the excel file
            fstream.Close();

            //Get raw and col number of excel file
            int col = worksheet.Cells.MaxDataColumn;
            int raw = worksheet.Cells.MaxDataRow;

            //define an array
            string[,] testArray = new string[raw+1,col+1];

            //send cell value to array
            for (int i = 0; i <= raw; i++)
                for (int j = 0; j <= col; j++)
                {                  
                   testArray[i, j] = (worksheet.Cells[i, j].Value ?? String.Empty).ToString();                               
                }
            return testArray;           
            
         }
    }        
}
