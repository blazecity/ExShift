﻿using ExShift.Util;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;
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
                // Create new data table
                ws = CreatePersistenceTable(tableName);

                // Create indizes
                List<PropertyInfo> indexProperties = AttributeHelper.GetPropertiesByAttribute<T>(typeof(Index));
                indexProperties.AddRange(AttributeHelper.GetPropertiesByAttribute<T>(typeof(PrimaryKey)));
                foreach (PropertyInfo indexProperty in indexProperties)
                {
                    CreateIndex<T>(indexProperty.Name);
                }
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
            // Persist object
            ObjectPackager packager = new ObjectPackager();
            string jsonPayload = packager.Package(obj);
            Worksheet table = GetPersistenceTable<T>();
            int row = ChangeRowCounter(typeof(T).Name, 1);
            table.Cells[row, 1].Value = jsonPayload;

            // Update indizes
            List<PropertyInfo> indexProperties = AttributeHelper.GetPropertiesByAttribute<T>(typeof(Index));
            indexProperties.AddRange(AttributeHelper.GetPropertiesByAttribute<T>(typeof(PrimaryKey)));
            foreach (PropertyInfo indexProperty in indexProperties)
            {
                UpdateIndex<T>(indexProperty.Name, indexProperty.GetValue(obj).ToString(), row);
            }
        }

        public static void UpdateIndex<T>(string propertyName, string key, int value) where T : IPersistable
        {
            Dictionary<string, List<int>> index = FindIndex<T>(propertyName);
            if (index.TryGetValue(key, out List<int> values))
            {
                values.Add(value);
            }
            else
            {
                values = new List<int>
                {
                    value
                };
                index.Add(key, values);
            }
            string tableName = "Idx_" + GetShortenedHash(typeof(T).Name + propertyName);
            Worksheet ws = FindTable(tableName);
            ws.Cells[1, 1].Value = JsonSerializer.Serialize(index);
        }

        public static string Find<T>(string primaryKey) where T : IPersistable
        {
            Worksheet table = GetPersistenceTable<T>();
            Dictionary<string, List<int>> index = FindIndex<T>(AttributeHelper.GetProperty<T>(typeof(PrimaryKey)).Name);
            int rowNumber = index[primaryKey][0];
            return table.Cells[rowNumber, 1].Value.ToString();
        }

        public static IEnumerable<string> GetAll<T>() where T : IPersistable
        {
            Range dataColumn = FindTable(typeof(T).Name).UsedRange.Columns[1];
            foreach (Range cell in dataColumn.Cells)
            {
                if (cell == null)
                {
                    yield break;
                }
                string cellValue = cell.Value;
                if (cellValue == null)
                {
                    yield break;
                }
                yield return cellValue.ToString();
            }
        }

        public static Dictionary<string, List<int>> FindIndex<T>(string property) where T : IPersistable
        {
            string tableName = "Idx_" + GetShortenedHash(typeof(T).Name + property);
            Worksheet indexTable = FindTable(tableName);
            if (indexTable == null)
            {
                return null;
            }
            string cellValue = indexTable.UsedRange[1, 1].Value.ToString();
            Dictionary<string, List<int>> index = JsonSerializer.Deserialize<Dictionary<string, List<int>>>(cellValue);
            return index;
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

            string tableName = "Idx_" + GetShortenedHash(typeof(T).Name + property);
            Worksheet indexTable = CreateUnformattedTable(tableName);
            Dictionary<string, List<int>> index = new Dictionary<string, List<int>>();
            int rowCounter = 1;
            foreach (string row in GetAll<T>())
            {
                JsonElement jsonElement = ObjectPackager.DeserializeTupel(row);
                JsonElement jsonProperty = jsonElement.GetProperty(property);
                string key = ObjectPackager.ConvertJsonElement(propertyType, jsonProperty).ToString();
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
