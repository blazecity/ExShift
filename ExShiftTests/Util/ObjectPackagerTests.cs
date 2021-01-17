using ExShift.Mapping;
using ExShift.Tests.Setup;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExShift.Util.Tests
{
    [TestClass()]
    public class ObjectPackagerTests
    {
        private Application app;
        private Workbook wb;

        [TestInitialize]
        public void Initalize()
        {
            app = new Application();
            wb = app.Workbooks.Add();
            ExcelObjectMapper.SetWorkbook(wb);
        }

        [TestCleanup]
        public void Cleanup()
        {
            wb.Close(0);
            app.Quit();
        }

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
            ObjectPackager op = new ObjectPackager(null, eom);
            PackageTestObject deserializedObject = op.Unpackage<PackageTestObject>(eom.Find<PackageTestObject>(testObject.BaseProperty.ToString()));
            Assert.AreEqual(testObject, deserializedObject);
        }
    }
}