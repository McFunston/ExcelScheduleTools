using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelDataGrabber;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ExcelScheduleToolsTests;

namespace ExcelDataGrabber.Tests
{

    [TestClass]
    public class ExcelFileTests
    {
        [TestMethod]

        public void GetCellContentsTest()
        {
            //arrange
            var fakeExcel = Mock.returnFakeExcel();

            //Act
            string expected = "1111112";
            var actual = fakeExcel.GetCellContents(0, 2);

            //Assert
            //Assert.AreEqual(3, fakeExcel.RowCount);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void RowCountTest()
        {
            //Arrange
            var fakeExcel = Mock.returnFakeExcel();

            //Act
            var actual = fakeExcel.RowCount;
            int expected = 3;

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ColumnCountTest()
        {
            //Arrange
            var fakeExcel = Mock.returnFakeExcel();

            //Act
            var actual = fakeExcel.ColumnCount;
            int expected = 8;

            Assert.AreEqual(expected, actual);
        }
    }
}