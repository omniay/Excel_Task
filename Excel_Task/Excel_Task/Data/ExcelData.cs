using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Task
{
   public class ExcelData
    {
        List<Object[]> celldata;
        public ExcelData()
        {
            celldata = new List<Object[]>()
            {
              new object[] {1, "Ahmed", "Zaki", "1st",1000},
              new object[] {2, "Mohamed", "Hegazy", "2nd",1000},
              new object[] {3, "Youssef", "Saad", "1st",1000}
            };
        }
        public List<Object[]> CellData
        {
            get
            {
                return celldata;
            }
        }
    }
}
