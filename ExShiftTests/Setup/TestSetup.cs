using ExShift.Mapping;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExShiftTests.Setup
{
    public class TestSetup
    {
        protected Application app;
        protected Workbook wb;

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
    }
}
