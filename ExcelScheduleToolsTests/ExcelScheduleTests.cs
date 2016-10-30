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
    }
}