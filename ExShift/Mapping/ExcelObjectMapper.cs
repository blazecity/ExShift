using ExShift.Util;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
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

        private static Worksheet CreateUnformattedTable(string name)
        {
            
            Worksheet ws = FindTable(name);
            if (ws == null)
            {
                ws = workbook.Worksheets.Add();
                ws.Name = name;
            }
            return ws;
            
        }

        private static Worksheet FindTable(string name)
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

        public static void Initialize()
        {
            Worksheet sysTable = CreateUnformattedTable("__sys");
            
            // Intialize ID counter
            sysTable.Cells[1, 1].Value = 1;

            // Initialize row counter
            sysTable.Cells[2, 1].Value = "{}";
        }

        private static Worksheet CreatePersistenceTable(string name)
        {
            // Create sheet
            Worksheet table = CreateUnformattedTable(name);
            table.Application.ActiveWindow.FreezePanes = true;
            table.Visible = XlSheetVisibility.xlSheetHidden;

            // Initialize row counter
            ChangeRowCounter(name, 1);
            return table;
        }

        public static Worksheet GetPersistenceTable<T>() where T : IPersistable
        {
            string tableName = typeof(T).Name;
            Worksheet ws = FindTable(tableName);
            if (ws == null)
            {
                return CreatePersistenceTable(tableName);
            }
            return ws;
        }

        public static Worksheet GetPersistenceTable(string tableName)
        {
            Worksheet ws = FindTable(tableName);
            if (ws == null)
            {
                return CreatePersistenceTable(tableName);
            }
            return ws;
        }

        private static int ChangeRowCounter(string tableName, int change)
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

        public static void Persist<T>(T obj) where T : IPersistable
        {
            ObjectPackager packager = new ObjectPackager();
            string jsonPayload = packager.Package(obj);
            Worksheet table = GetPersistenceTable<T>();
            int row = ChangeRowCounter(typeof(T).Name, 1);
            table.Cells[row, 1].Value = AttributeHelper.GetPrimaryKey(obj);
            table.Cells[row, 2].Value = jsonPayload;
            Range usedRange = table.UsedRange;
            usedRange.Sort(usedRange.Columns[1], XlSortOrder.xlAscending);
        }

        public static string Find<T>(string primaryKey) where T : IPersistable
        {
            return Find(typeof(T).Name, primaryKey);
        }

        public static string Find(string tableName, string primaryKey)
        {
            Worksheet table = GetPersistenceTable(tableName);
            return BinarySearch(table, primaryKey);
        }

        public static IEnumerable<string> GetAll<T>()
        {
            Range dataColumn = FindTable(typeof(T).Name).UsedRange.Columns[2];
            foreach (Range cell in dataColumn.Cells)
            {
                yield return cell.Value.ToString();
            }
        }

        private static string BinarySearch(Worksheet table, IComparable target)
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

        public static Worksheet FindIndex<T>(string property) where T : IPersistable
        {
            return FindTable("Idx_" + GetShortenedHash(typeof(T).Name + property));
        }

        private static string GetShortenedHash(string text)
        {
            byte[] encoded = Encoding.UTF8.GetBytes(text);
            SHA256 sha256 = SHA256.Create();
            byte[] hash = sha256.ComputeHash(encoded);
            char[] shortenedHash = new char[10];
            for (int i = 0; i < shortenedHash.Length; i++)
            {
                shortenedHash[i] = text[hash[i] % encoded.Length];
            }
            return new string(shortenedHash);
        }

        public static bool CreateIndex<T>(string property) where T : IPersistable
        {
            Type propertyType = typeof(T).GetProperty(property).PropertyType;

            // Check if type and property match
            if (propertyType == null)
            {
                return false;
            }

            // Check if there is already an index for the specified table and attribute
            if (FindIndex<T>(property) != null)
            {
                return true;
            }

            if (FindTable(typeof(T).Name) == null)
            {
                return false;
            }

            Worksheet indexTable = CreateUnformattedTable("Idx_" + GetShortenedHash(typeof(T).Name + property));

            Dictionary<dynamic, List<int>> index = new Dictionary<dynamic, List<int>>();
            int rowCounter = 1;
            foreach (string row in GetAll<T>())
            {
                JsonElement jsonElement = ObjectPackager.DeserializeTupel(row);
                JsonElement jsonProperty = jsonElement.GetProperty(property);
                dynamic key = ObjectPackager.ConvertJsonElement(propertyType, jsonProperty);
                if (!index.TryGetValue(key, out List<int> value))
                {
                    List<int> newValue = new List<int>
                    {
                        rowCounter
                    };
                    index.Add(key, newValue);
                }
                else
                {
                    value.Add(rowCounter);
                }
            }
            indexTable.Cells[1, 1].Value = JsonSerializer.Serialize(index);
            return true;
        }

        public static void Update(IPersistable obj)
        {

        }

        public static void Delete(IPersistable obj)
        {
            
        }
    }
}
