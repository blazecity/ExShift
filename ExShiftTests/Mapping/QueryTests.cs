using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExShiftTests.Setup;
using ExShift.Tests.Setup;
using System.Collections.Generic;

namespace ExShift.Mapping.Tests
{
    [TestClass()]
    public class QueryTests : TestSetup
    {
        [TestInitialize]
        public void PersistObjects()
        {
            wb.Close(0);
            wb = app.Workbooks.Add();
            ExcelObjectMapper.SetWorkbook(wb);
            ExcelObjectMapper.Initialize();

            // Arrange
            PackageTestObject obj = new PackageTestObject(1, 2);
            foreach (PackageTestObjectNested nestedObj in obj.ListOfNestedObjects)
            {
                ExcelObjectMapper.Persist(nestedObj);
            }
            ExcelObjectMapper.Persist(obj.NestedObject);
            ExcelObjectMapper.Persist(obj);
            for (int i = 0; i < 5; i++)
            {
                ExcelObjectMapper.Persist(new PackageTestObject(i + 2, i + 3));
            }
        }

        [TestMethod("Only where clause")]
        public void SelectTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("BaseProperty = 'base_1'")
                                                                         .Run();

            // Assert
            Assert.AreEqual(6, resultList.Count);
        }

        [TestMethod("Search with AND-Operator")]
        public void SelectAndTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("BaseProperty = 'base_1'")
                                                                         .And("Property = 2")
                                                                         .Run();

            // Assert
            Assert.AreEqual(1, resultList.Count);
        }

        [TestMethod("Search with OR-Operator")]
        public void SelectOrTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("BaseProperty = 'base_1'")
                                                                         .Or("Property = 2")
                                                                         .Run();

            // Assert
            Assert.AreEqual(6, resultList.Count);
        }

        [TestMethod("Search with non-existant property")]
        public void SelectNonExistentPropertyTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("xy = 1")
                                                                         .Run();

            // Assert
            Assert.AreEqual(0, resultList.Count);
        }

        [TestMethod("Search with three query nodes")]
        public void SelectThreeQueryNodesTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("BaseProperty = 'base_1'")
                                                                         .Or("DerivedProperty = 3")
                                                                         .Or("Property = 2")
                                                                         .Run();

            // Assert
            Assert.AreEqual(6, resultList.Count);
        }

        [TestMethod("Search with five query nodes")]
        public void SelectFiveQueryNodesTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("Property = 2")
                                                                         .Or("BaseProperty = 'base_1'")
                                                                         .And("DerivedProperty = 3")
                                                                         .Or("AnotherBaseProperty = 'abp'")
                                                                         .And("Property = 3")
                                                                         .Run();

            // Assert
            Assert.AreEqual(1, resultList.Count);
        }
    }
}