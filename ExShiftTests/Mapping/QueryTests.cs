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
            // Arrange
            ExcelObjectMapper eom = new ExcelObjectMapper();
            eom.Initialize();
            for (int i = 0; i < 5; i++)
            {
                PackageTestObject testObject = new PackageTestObject(i, i + 1);
                foreach (PackageTestObjectNested nestedObj in testObject.ListOfNestedObjects)
                {
                    eom.Persist(nestedObj);
                }
                eom.Persist(testObject.NestedObject);
                eom.Persist(testObject);
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
            Assert.AreEqual(5, resultList.Count);
        }

        [TestMethod("Search with AND-Operator")]
        public void SelectAndTest()
        {
            // Act
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select()
                                                                         .Where("BaseProperty = 'base_1'")
                                                                         .And("PrimaryKey = 2")
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
                                                                         .Or("PrimaryKey = 2")
                                                                         .Run();

            // Assert
            Assert.AreEqual(5, resultList.Count);
        }
    }
}