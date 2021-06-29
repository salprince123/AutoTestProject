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
        [Test]
        [TestCase(1, 2000, 31)]
        [TestCase(2, 2000, 29)]
        [TestCase(4, 2000, 30)]
        [TestCase(13, 2000, 0)]
        [TestCase(0, 2021, 0)]
        [TestCase(1, 2021, 31)]
        [TestCase(2, 2021, 28)]
        [TestCase(4, 2021, 30)]
        [TestCase(13, 2021, 0)]

        public void autoTestDayInMonth(int month, int year, int expected)
        {
            Assert.AreEqual(expected, dayInMonth(month, year));
        }
        public bool isValidDate(int day, int month, int year)
        {
            if (year < 1 && year > 12)
                return false;
            if (day < 1)
                return false;
            if (day > dayInMonth(month, year))
                return false;
            return true;
        }
        [Test]
        [TestCase(29, 2, 2000, true)]
        [TestCase(29, 2, 2016, true)]
        [TestCase(29, 2, 2100, false)]
        [TestCase(29, 2, 2021, false)]
        [TestCase(29, 3, 2000, true)]
        [TestCase(29, 3, 2016, true)]
        [TestCase(29, 3, 2100, true)]
        [TestCase(29, 3, 2021, true)]
        [TestCase(29, 4, 2000, true)]
        [TestCase(29, 4, 2016, true)]
        [TestCase(29, 4, 2100, true)]
        [TestCase(29, 4, 2021, true)]
        [TestCase(30, 2, 2000, false)]
        [TestCase(30, 2, 2016, false)]
        [TestCase(30, 2, 2100, false)]
        [TestCase(30, 2, 2021, false)]
        [TestCase(30, 3, 2000, true)]
        [TestCase(30, 3, 2016, true)]
        [TestCase(30, 3, 2100, true)]
        [TestCase(30, 3, 2021, true)]
        [TestCase(30, 4, 2000, true)]
        [TestCase(30, 4, 2016, true)]
        [TestCase(30, 4, 2100, true)]
        [TestCase(30, 4, 2021, true)]
        [TestCase(31, 2, 2000, false)]
        [TestCase(31, 2, 2016, false)]
        [TestCase(31, 2, 2100, false)]
        [TestCase(31, 2, 2021, false)]
        [TestCase(31, 3, 2000, true)]
        [TestCase(31, 3, 2016, true)]
        [TestCase(31, 3, 2100, true)]
        [TestCase(31, 3, 2021, true)]
        [TestCase(31, 4, 2000, false)]
        [TestCase(31, 4, 2016, false)]
        [TestCase(31, 4, 2100, false)]
        [TestCase(31, 4, 2021, false)]
        public void autoTestIsValidDate(int day, int month, int year, bool expected)
        {
            Assert.AreEqual(expected, isValidDate(day, month, year));
        }

        

    }
}
