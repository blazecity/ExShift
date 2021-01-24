﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExShift.Tests.Setup;
using System.Collections.Generic;
using ExShiftTests.Setup;

namespace ExShift.Mapping.Tests
{
    [TestClass()]
    public class ExcelObjectMapperTests : TestSetup
    {
        public ExcelObjectMapperTests() : base()
        {
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

        [TestMethod("Get all")]
        public void GetAllTest()
        {
            // Act
            IEnumerable<string> resultList = ExcelObjectMapper.GetAll<PackageTestObject>();
            byte counter = 0;
            foreach (string s in resultList)
            {
                counter++;
            }

            // Assert
            Assert.AreEqual(6, counter);
        }

        [TestMethod("Persist nested objects")]
        public void PersistWithNestedObjectsTest()
        {
            // Act
            PackageTestObject retrievedObject = ExcelObjectMapper.Find<PackageTestObject>("2");

            // Assert
            Assert.AreEqual(3, retrievedObject.ListOfNestedObjects.Count);
        }

        [TestMethod("Create index with integer")]
        public void CreateIndexIntTest()
        {
            // Act
            Dictionary<string, List<int>> index = ExcelObjectMapper.FindIndex<PackageTestObject>("DerivedProperty");

            // Assert
            Assert.AreEqual(1, index["2"][0]);
        }

        [TestMethod("Create index with string")]
        public void CreateIndexStringTest()
        {
            // Act
            string propertyName = "BaseProperty";
            ExcelObjectMapper.CreateIndex<PackageTestObject>(propertyName);
            Dictionary<string, List<int>> index = ExcelObjectMapper.FindIndex<PackageTestObject>(propertyName);

            // Assert
            Assert.AreEqual(1, index["base_1"][0]);
        }

        [TestMethod("Update entry")]
        public void UpdateEntryTest()
        {
            // Act
            PackageTestObject retrievedObject = ExcelObjectMapper.Find<PackageTestObject>("3");
            retrievedObject.Property = 99;
            ExcelObjectMapper.Update(retrievedObject);

            PackageTestObject objectAfterUpdate = ExcelObjectMapper.Find<PackageTestObject>("3");

            // Assert
            Assert.AreEqual(99, objectAfterUpdate.Property);
        }
    }
}