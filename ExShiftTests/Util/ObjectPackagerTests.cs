using ExShift.Mapping;
using ExShift.Tests.Setup;
using ExShiftTests.Setup;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExShift.Util.Tests
{
    [TestClass()]
    public class ObjectPackagerTests : TestSetup
    {

        [TestMethod("Serialize and Deserialize")]
        public void ObjectPackagerDeserializeTest()
        {
            PackageTestObject testObject = new PackageTestObject(1, 2);
            ExcelObjectMapper.Persist(testObject.NestedObject);
            foreach (PackageTestObjectNested obj in testObject.ListOfNestedObjects)
            {
                ExcelObjectMapper.Persist(obj);
            }
            ExcelObjectMapper.Persist(testObject);
            ObjectPackager op = new ObjectPackager();
            PackageTestObject deserializedObject = op.Unpackage<PackageTestObject>(ExcelObjectMapper.Find<PackageTestObject>(testObject.BaseProperty.ToString()));
            Assert.AreEqual(testObject, deserializedObject);
        }
    }
}