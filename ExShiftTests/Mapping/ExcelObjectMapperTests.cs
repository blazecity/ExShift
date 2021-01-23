using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExShift.Tests.Setup;
using System.Collections.Generic;
using ExShiftTests.Setup;
using ExShift.Util;
using Microsoft.Office.Interop.Excel;
using System.Text.Json;

namespace ExShift.Mapping.Tests
{
    [TestClass()]
    public class ExcelObjectMapperTests : TestSetup
    {
        [TestMethod("Get all")]
        public void GetAllTest()
        {
            // Arrange
            for (int i = 0; i < 5; i++)
            {
                ExcelObjectMapper.Persist(new PackageTestObject(i, i + 1));
            }

            // Act
            IEnumerable<string> resultList = ExcelObjectMapper.GetAll<PackageTestObject>();
            byte counter = 0;
            foreach (string s in resultList)
            {
                counter++;
            }

            // Assert
            Assert.AreEqual(5, counter);
        }

        [TestMethod("Persist nested objects")]
        public void PersistWithNestedObjectsTest()
        {
            // Arrange
            PackageTestObject obj = new PackageTestObject(1, 2);
            foreach (PackageTestObjectNested nestedObj in obj.ListOfNestedObjects)
            {
                ExcelObjectMapper.Persist(nestedObj);
            }
            ExcelObjectMapper.Persist(obj.NestedObject);
            ExcelObjectMapper.Persist(obj);

            // Act
            string result = ExcelObjectMapper.Find<PackageTestObject>(obj.BaseProperty.ToString());
            ObjectPackager op = new ObjectPackager();
            PackageTestObject retrievedObject = op.Unpackage<PackageTestObject>(result);
            Dictionary<string, List<int>> index = ExcelObjectMapper.FindIndex<PackageTestObject>("Property");

            // Assert
            Assert.AreEqual(3, retrievedObject.ListOfNestedObjects.Count);
            Assert.AreEqual(1, index["1"][0]);

        }

        [TestMethod("Create index with integer")]
        public void CreateIndexIntTest()
        {
            // Arrange
            PackageTestObject obj = new PackageTestObject(1, 2);
            foreach (PackageTestObjectNested nestedObj in obj.ListOfNestedObjects)
            {
                ExcelObjectMapper.Persist(nestedObj);
            }
            ExcelObjectMapper.Persist(obj.NestedObject);
            ExcelObjectMapper.Persist(obj);

            // Act
            string propertyName = "DerivedProperty";
            ExcelObjectMapper.CreateIndex<PackageTestObject>(propertyName);
            Dictionary<string, List<int>> index = ExcelObjectMapper.FindIndex<PackageTestObject>(propertyName);

            // Assert
            Assert.AreEqual(1, index["2"][0]);
        }

        [TestMethod("Create index with string")]
        public void CreateIndexStringTest()
        {
            // Arrange
            PackageTestObject obj = new PackageTestObject(1, 2);
            foreach (PackageTestObjectNested nestedObj in obj.ListOfNestedObjects)
            {
                ExcelObjectMapper.Persist(nestedObj);
            }
            ExcelObjectMapper.Persist(obj.NestedObject);
            ExcelObjectMapper.Persist(obj);

            // Act
            string propertyName = "BaseProperty";
            ExcelObjectMapper.CreateIndex<PackageTestObject>(propertyName);
            Dictionary<string, List<int>> index = ExcelObjectMapper.FindIndex<PackageTestObject>(propertyName);

            // Assert
            Assert.AreEqual(1, index["base_1"][0]);
        }
    }
}