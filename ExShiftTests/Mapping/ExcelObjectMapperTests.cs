using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExShift.Tests.Setup;
using System.Collections.Generic;
using ExShiftTests.Setup;

namespace ExShift.Mapping.Tests
{
    [TestClass()]
    public class ExcelObjectMapperTests : TestSetup
    {
        [TestInitialize]
        public void TestInitialize()
        {
            wb.Close(0);
            wb = app.Workbooks.Add();
            ExcelObjectMapper.Initialize(wb);

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
            byte counter = 0;
            foreach (string s in ExcelObjectMapper.GetAll<PackageTestObject>())
            {
                counter++;
            }

            // Assert
            Assert.AreEqual(6, counter);
        }

        [TestMethod("Get all objects")]
        public void GetAllObjectsTest()
        {
            // Act
            byte counter = 0;
            foreach (PackageTestObject obj in ExcelObjectMapper.GetAllObjects<PackageTestObject>())
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

        [TestMethod("Persist generic types")]
        public void PersistGenericTypesTest()
        {
            // Arrange
            List<int> numbers = new List<int> { 1, 2, 34, 6, 8787 };
            GenericTestObjectPT<int> obj1 = new GenericTestObjectPT<int> { Pk = 1, List = numbers };
            List<PackageTestObject> testObjects = new List<PackageTestObject>(ExcelObjectMapper.GetAllObjects<PackageTestObject>());
            GTO<PackageTestObject> obj2 = new GTO<PackageTestObject> { Pk = 1, List =  testObjects};

            // Act
            ExcelObjectMapper.Persist(obj1);
            ExcelObjectMapper.Persist(obj2);

            GenericTestObjectPT<int> obj1AfterPersistence = ExcelObjectMapper.Find<GenericTestObjectPT<int>>(AttributeHelper.GetPrimaryKey(obj1));
            GTO<PackageTestObject> obj2AfterPersistence = ExcelObjectMapper.Find<GTO<PackageTestObject>>(AttributeHelper.GetPrimaryKey(obj2));
            Assert.IsNotNull(obj1AfterPersistence);
            Assert.IsNotNull(obj2AfterPersistence);
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

        [TestMethod("Delete entry")]
        public void DeleteEntryTest()
        {
            // Act
            PackageTestObject retrievedObject = ExcelObjectMapper.Find<PackageTestObject>("4");
            ExcelObjectMapper.Delete(retrievedObject);

            PackageTestObject objectAfterDeletion = ExcelObjectMapper.Find<PackageTestObject>("4");

            // Assert
            Assert.IsNull(objectAfterDeletion);
        }

        [TestMethod("Unique primary key")]
        public void UniquePrimaryKeyTest()
        {
            // Arrange
            wb.Close(0);
            wb = app.Workbooks.Add();
            ExcelObjectMapper.Initialize(wb);
            PackageTestObject obj1 = new PackageTestObject(1, 2);
            PackageTestObject obj2 = new PackageTestObject(1, 2);

            // Act
            foreach (PackageTestObjectNested nestedObj in obj1.ListOfNestedObjects)
            {
                ExcelObjectMapper.Persist(nestedObj);
            }
            ExcelObjectMapper.Persist(obj1.NestedObject);
            ExcelObjectMapper.Persist(obj1);
            ExcelObjectMapper.Persist(obj2);
            List<PackageTestObject> resultList = Query<PackageTestObject>.Select().Run();

            // Assert
            Assert.AreEqual(1, resultList.Count);
        }
    }
}