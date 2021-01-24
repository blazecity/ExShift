using ExShift.Mapping;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Runtime.InteropServices;

namespace ExShiftTests.Setup
{
    [TestClass]
    public class TestSetup
    {
        protected static Application app;
        protected static Workbook wb;

        [AssemblyInitialize]
        public static void Setup(TestContext testContext)
        {
            app = new Application();
            wb = app.Workbooks.Add();
            ExcelObjectMapper.SetWorkbook(wb);
            ExcelObjectMapper.Initialize();
        }

        [AssemblyCleanup]
        public static void Teardown()
        {
            wb.Close(0);
            app.Quit();
            Marshal.ReleaseComObject(wb);
            wb = null;
            Marshal.ReleaseComObject(app);
            app = null;
        }
    }
}
