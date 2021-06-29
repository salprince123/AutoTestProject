using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTestProject.Tool
{
    public class TableString
    {
        public String path { get; set; }
        public TableString(String path)
        {
            this.path = path;
        }
        public String getRowPath(String path, int number)
        {
            return $"{path}[{number}]";
        }
        public String getCellPath(String path, int row, int col)
        {
            return $"{path}[{row}]/td[{col}]";
        }
        public String getCellPath(int row, int col)
        {
            return $"{path}[{row}]/td[{col}]";
        }
    }
}
