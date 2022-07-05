using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCSVConversion
{
    internal class Product
    {
        public string pid;
        public string cost;
        public string productId;
        public string mfrName;
        public string mfrPn;
        public string coo;
        public string shortDescrip;
        public string upc;
        public string uom;

        public string costToString;
        
        public Product(string[] product)
        {
            
            this.pid = product[0];

            //convert cost to int, add 20%, convert back to string
            this.cost = stringConvert(costToString, Convert.ToDouble(product[4]));
            
            this.productId = product[1];
            this.mfrName = product[2];
            this.mfrPn = product[3];
            this.coo = product[5];
            this.shortDescrip = product[6];
            this.upc = product[7];
            this.uom = product[8];

            
        }

        static string stringConvert(string costToString, double cost)
        {
            //add 20% to cost
            double tax = cost * .20;
            cost = cost + tax;

            //convert back to string
            costToString = cost.ToString();

            return (costToString);

        }
       }
}
