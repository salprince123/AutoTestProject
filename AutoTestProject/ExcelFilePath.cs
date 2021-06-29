using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTestProject
{
    public class ExcelFilePath
    {
        private string path { get; set; } 
        public ExcelFilePath()
        {
            path = "D:\\DataAutotest.xlsx";
        }
        public string showPath()
        {
            return path;
        }
    }
}
