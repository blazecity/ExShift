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
            ExcelObjectMapper eom = new ExcelObjectMapper();
            eom.Initialize();
            PackageTestObject testObject = new PackageTestObject(1, 2);
            eom.Persist(testObject.NestedObject);
            foreach (PackageTestObjectNested obj in testObject.ListOfNestedObjects)
            {
                eom.Persist(obj);
            }
            eom.Persist(testObject);
            ObjectPackager op = new ObjectPackager(null);
            PackageTestObject deserializedObject = op.Unpackage<PackageTestObject>(eom.Find<PackageTestObject>(testObject.BaseProperty.ToString()));
            Assert.AreEqual(testObject, deserializedObject);
        }
    }
}