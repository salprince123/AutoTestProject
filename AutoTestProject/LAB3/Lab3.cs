using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTestProject.LAB3
{
    public class Lab3
    {
        public int dayInMonth(int month, int year)
        {
            switch (month)
            {
                case 1: case 3: case 5: case 7: case 8: case 10: case 12: return 31;
                case 4: case 6: case 9: case 11: return 30;
                case 2:
                    {
                        if (year % 400 == 0)
                            return 29;
                        else if (year % 100 == 0)
                            return 28;
                        else if (year % 4 == 0)
                            return 29;
                        else return 28;
                    }
                default: return 0;

            }
        }
        public bool isValidDate(int day, int month, int year)
        {
            if (year < 1 && year > 12)
                return false;
            if (day < 1)
                return false;
            if (day > dayInMonth(day, year))
                return false;
            return true;
        }
        [Test]
        [TestCase(1,1,2021,true)]
        [TestCase(1, 1, 2022, true)]
        [TestCase(1, 1, 2023, true)]
        public void autoTestIsValidDate(int day, int month, int year, bool expected)
        {
            Assert.AreEqual(expected, isValidDate(day, month, year));
        }
        
    }
}
