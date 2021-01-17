using ExShift.Util;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace ExShift.Mapping
{
    public class ExcelObjectMapper
    {
        private static Workbook workbook;

        public static void SetWorkbook(Workbook workbook)
        {
            ExcelObjectMapper.workbook = workbook;
        }

        private Worksheet CreateUnformattedTable(string name)
        {
            Worksheet ws = workbook.Worksheets.Add();
            ws.Name = name;
            return ws;
        }

        private Worksheet FindTable(string name)
        {
            try
            {
                return workbook.Worksheets[name];
            } 
            catch (Exception)
            {
                return null;
            }
        }

        public void Initialize()
        {
            Worksheet sysTable = CreateUnformattedTable("__sys");
            
            // Intialize ID counter
            sysTable.Cells[1, 1].Value = 1;

            // Initialize row counter
            sysTable.Cells[2, 1].Value = "{}";
        }

        private Worksheet CreatePersistenceTable(string name)
        {
            // Create sheet
            Worksheet table = CreateUnformattedTable(name);
            table.Application.ActiveWindow.FreezePanes = true;
            table.Visible = XlSheetVisibility.xlSheetHidden;

            // Initialize row counter
            ChangeRowCounter(name, 1);
            return table;
        }

        public Worksheet GetPersistenceTable<T>() where T : IPersistable
        {
            string tableName = typeof(T).Name;
            Worksheet ws = FindTable(tableName);
            if (ws == null)
            {
                return CreatePersistenceTable(tableName);
            }
            return ws;
        }

        public Worksheet GetPersistenceTable(string tableName)
        {
            Worksheet ws = FindTable(tableName);
            if (ws == null)
            {
                return CreatePersistenceTable(tableName);
            }
            return ws;
        }

        private int ChangeRowCounter(string tableName, int change)
        {
            Worksheet sysTable = FindTable("__sys");
            string jsonCounter = sysTable.Cells[2, 1].Value;
            Dictionary<string, int> dictCounter = JsonSerializer.Deserialize<Dictionary<string, int>>(jsonCounter);
            if (dictCounter.ContainsKey(tableName))
            {
                dictCounter[tableName] += change;
            }
            else
            {
                dictCounter[tableName] = 0;
            }
            sysTable.Cells[2, 1].Value = JsonSerializer.Serialize(dictCounter);
            return dictCounter[tableName];
        }

        public void Persist<T>(T obj) where T : IPersistable
        {
            ObjectPackager packager = new ObjectPackager(obj, this);
            string jsonPayload = packager.Package();
            Worksheet table = GetPersistenceTable<T>();
            int row = ChangeRowCounter(typeof(T).Name, 1);
            table.Cells[row, 1].Value = AttributeHelper.GetPrimaryKey(obj);
            table.Cells[row, 2].Value = jsonPayload;
            Range usedRange = table.UsedRange;
            usedRange.Sort(usedRange.Columns[1], XlSortOrder.xlAscending);
        }

        public string Find<T>(string primaryKey) where T : IPersistable
        {
            return Find(typeof(T).Name, primaryKey);
        }

        public string Find(string tableName, string primaryKey)
        {
            Worksheet table = GetPersistenceTable(tableName);
            return BinarySearch(table, primaryKey);
        }

        public List<string> GetAll<T>()
        {
            List<string> result = new List<string>();
            Range dataColumn = FindTable(typeof(T).Name).UsedRange.Columns[2];
            foreach (Range cell in dataColumn.Cells)
            {
                result.Add(cell.Value.ToString());
            }
            return result;
        }

        private string BinarySearch(Worksheet table, IComparable target)
        {
            Range primaryColumn = table.UsedRange.Columns[1];
            
            int left = 1;
            int right = primaryColumn.Rows.Count;
            int mid;
            string targetValue = target.ToString();

            while (left <= right)
            {
                mid = left + (right - left) / 2;
                string cellValue = primaryColumn.Cells[mid, 1].Value.ToString();
                if (cellValue == targetValue)
                {
                    return table.UsedRange.Cells[mid, 2].Value.ToString();
                }

                int compare = targetValue.CompareTo(cellValue);
                if (compare > 0) // shift right
                {
                    left = mid + 1;
                }
                else // shift left
                {
                    right = mid - 1;
                }
            }
            return "";
        }

        public void CreateIndex()
        {

        }

        public void Update(IPersistable obj)
        {

        }

        public void Delete(IPersistable obj)
        {
            
        }
    }
}
