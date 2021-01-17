using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Excel;
using ExShift.Tests.Setup;
using System.Collections.Generic;

namespace ExShift.Mapping.Tests
{
    [TestClass()]
    public class ExcelObjectMapperTests
    {
        [TestInitialize]
        public void Initalize()
        {
            ExcelObjectMapper.SetWorkbook(new Application().Workbooks.Add());
        }

        [TestMethod("Get all")]
        public void GetAllTest()
        {
            // Arrange
            ExcelObjectMapper eom = new ExcelObjectMapper();
            eom.Initialize();
            for (int i = 0; i < 5; i++)
            {
                eom.Persist(new PackageTestObject(i, i + 1));
            }

            // Act
            List<string> resultList = eom.GetAll<PackageTestObject>();

            // Assert
            Assert.AreEqual(5, resultList.Count);
        }
    }
}