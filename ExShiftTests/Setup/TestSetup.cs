using ExShift.Mapping;
using Microsoft.Office.Interop.Excel;

namespace ExShiftTests.Setup
{
    public class TestSetup
    {
        protected Application app;
        protected Workbook wb;

        public TestSetup()
        {
            app = new Application();
            wb = app.Workbooks.Add();
            ExcelObjectMapper.SetWorkbook(wb);
            ExcelObjectMapper.Initialize();
        }

        ~TestSetup()
        {
            wb.Close(false);
            app.Quit();
        }
    }
}
