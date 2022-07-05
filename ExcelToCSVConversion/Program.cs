using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToCSVConversion
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //create list, move to for loop to check length up to 10,000, create new when max limit is reached
            

            
            createExcel();

            //print statements to keep track of program location
            Console.WriteLine("created excel object");
            Console.WriteLine("csvCreated");
            Console.ReadLine();
        }


        static void createExcel()
        {
            //initialize required excel variables
            _Excel.Application xlApp; 
            _Excel.Workbook xlWorkBook;
            _Excel.Worksheet xlWorkSheet;
            _Excel.Range range;
            int rw = 0;
            int cl = 0;

            //read in excel file
            xlApp = new _Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\rknig\Documents\Product.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (_Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //get range of excel file
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            //Console.WriteLine("cl = " + cl);

            List<Product> productList = new List<Product>();
            //for loop that stores product information from excel sheet, set to only first 5 products for testing purposes
            for (int i = 2; i <= 5; i++)
            {
                string[] product = new string[10];
                for (int j = 1; j <= cl; j++)
                {
                    //get contents of cell
                    string str = Convert.ToString((range.Cells[i, j] as _Excel.Range).Value2);
                    
                    //store in product array
                    product[j - 1] = str;
                }

                //if list has exceeded 10000, needed because csv can only store up to 10000 products
                if(productList.Count() == 10000)
                {
                    //send list to csv file maker
                    createCSV(productList);

                    //clear list
                    productList.Clear();
                }
                //create product from product object, store in productList
                Product product1 = new Product(product);
                productList.Add(product1);

             

            }
        }

        static void createCSV(List<Product> productList)
        {
            //variables for while loop
            bool keepGoing = true;
            int csvNum = 0;
            var csvPath = @"C:\Documents\csvoutput" + csvNum + ".csv";
            //create csvfile
            while (keepGoing)
            {
                //if path already exists, create new path variable, this allows for multiple csvs to be made, rather than overwritten
                //if excel sheet has more than 10000 products
                if(File.Exists(csvPath))
                {
                    csvNum++;
                    csvPath = @"C:\Documents\csvoutput" + csvNum + ".csv";
                }
                
            }

            //create steamWriter and csvWriter objects
            using(var streamwriter = new StreamWriter(csvPath))
            {
                using(var csvWriter = new CsvWriter(streamwriter, CultureInfo.InvariantCulture))
                {
                    //write to csv file using productList
                    csvWriter.WriteRecords(productList);
                    
                }
            }






        }


       
    }
}
