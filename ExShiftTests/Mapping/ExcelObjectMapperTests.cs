using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExShift.Tests.Setup;
using System.Collections.Generic;
using ExShiftTests.Setup;
using ExShift.Util;

namespace ExShift.Mapping.Tests
{
    [TestClass()]
    public class ExcelObjectMapperTests : TestSetup
    {
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
            IEnumerable<string> resultList = eom.GetAll<PackageTestObject>();
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
            ExcelObjectMapper eom = new ExcelObjectMapper();
            eom.Initialize();
            PackageTestObject obj = new PackageTestObject(1, 2);
            foreach (PackageTestObjectNested nestedObj in obj.ListOfNestedObjects)
            {
                eom.Persist(nestedObj);
            }
            eom.Persist(obj.NestedObject);
            eom.Persist(obj);

            // Act
            string result = eom.Find<PackageTestObject>(obj.BaseProperty.ToString());
            ObjectPackager op = new ObjectPackager(null);
            PackageTestObject retrievedObject = op.Unpackage<PackageTestObject>(result);

            // Assert
            Assert.AreEqual(3, retrievedObject.ListOfNestedObjects.Count);
        }
    }
}