using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Entities
{
    public class Order
    {
        public int OrderID { get; set; }
        public string DesignName { get; set; }
        public string UserName { get; set; }
        public string BackColor { get; set; }
    }
}
