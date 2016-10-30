using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelDataGrabber;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelScheduleToolsTests;

namespace ExcelDataGrabber.Tests
{
    [TestClass]
    public class ExcelScheduleTests
    {
        [TestMethod]
        public void ReturnJobNumbersTest()
        {
            //Arrange
            var fakeExcel = Mock.returnFakeSchedule();

            //Act
            var actual = fakeExcel.ReturnJobNumbers().Count;
            int expected = 3;

            //Assert
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReturnUniqueJobNumbersTest()
        {
            //Arrange
            var fakeExcel = Mock.returnFakeSchedule();

            //Act
            var actual = fakeExcel.ReturnUniqueJobNumbers().Count;
            int expected = 2;

            //Assert
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ReturnDateTimeFromCellTest()
        {
            //Arrange
            var fakeExcel = Mock.returnFakeSchedule();

            //Act
            var actual = fakeExcel.ReturnDateTimeFromCell(2, "111112", "yy/MM/dd h:mm");
            DateTime expected = new DateTime(2017, 11, 21, 8, 0, 0);
            Assert.AreEqual(expected.Ticks, actual.Ticks);
        }

        [TestMethod()]
        public void GetJobNumberRowTest()
        {
            //Arrange
            var fakeExcel = Mock.returnFakeSchedule();

            //Act
            var actual = fakeExcel.GetJobNumberRow("111112");
            int expected = 1; 

            Assert.AreEqual(expected, actual);
        }
    }
}