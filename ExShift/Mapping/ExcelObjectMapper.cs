using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ExShift.Mapping
{
    /// <summary>
    /// This class provides the functionality for handling data, such as 
    /// persiting, updating, deleting and searching (to a certain extent).
    /// </summary>
    public class ExcelObjectMapper
    {
        private static Workbook workbook;

        /// <summary>
        /// Sets the current Excel <see cref="Workbook"/> you need to work in.
        /// </summary>
        /// <param name="workbook"><see cref="Workbook"/></param>
        public static void SetWorkbook(Workbook workbook)
        {
            ExcelObjectMapper.workbook = workbook;
        }

        /// <summary>
        /// Creates a plain Excel <see cref="Worksheet"/>.
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>A new worksheet or an already existing one with the specified name.</returns>
        private static Worksheet CreateUnformattedTable(string name)
        {
            
            Worksheet ws = FindTable(name);
            if (ws == null)
            {
                ws = workbook.Worksheets.Add();
                ws.Name = name;
                ws.Visible = XlSheetVisibility.xlSheetHidden;
            }
            return ws;
            
        }

        /// <summary>
        /// Finds an Excel <see cref="Worksheet"/> with the given name.
        /// </summary>
        /// <param name="name">Worksheet name</param>
        /// <returns>
        /// <see cref="Worksheet"/> if one exists with the specified name 
        /// or else null if none is found.
        /// </returns>
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

        /// <summary>
        /// Sets up the internal helper worksheets.
        /// </summary>
        public static void Initialize()
        {
            string sysTableName = "__sys";
            Worksheet sysTable = FindTable(sysTableName);
            if (sysTable == null)
            {
                sysTable = CreateUnformattedTable("__sys");

                // Intialize ID counter
                sysTable.Cells[1, 1].Value = 1;

                // Initialize row counter
                sysTable.Cells[2, 1].Value = "{}";
            }
        }

        /// <summary>
        /// Creates a <see cref="Worksheet"/> for data.
        /// </summary>
        /// <param name="name">Worksheet name</param>
        /// <returns><see cref="Worksheet"/></returns>
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

        /// <summary>
        /// Gets a data table (<see cref="Worksheet"/>).
        /// Note that, if the table does not exist a new one will be created. 
        /// Also the indizes are initalized automatically.
        /// </summary>
        /// <typeparam name="T">Type of data which will be stored</typeparam>
        /// <returns>New or existing data table</returns>
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

        /// <summary>
        /// Increments or decrements the row counter by the specified amount.
        /// </summary>
        /// <param name="tableName">Worksheet name</param>
        /// <param name="change">Increment or decrement amount</param>
        /// <returns>Actualized row number of the first empty row</returns>
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

        /// <summary>
        /// Persist an object.
        /// </summary>
        /// <typeparam name="T">Type of object to be persisted</typeparam>
        /// <param name="obj">Object to persisted</param>
        public static void Persist<T>(T obj) where T : IPersistable
        {
            // Check for primary key (must be unique)
            Worksheet table = GetPersistenceTable<T>();
            string primaryKey = AttributeHelper.GetPrimaryKey(obj);
            PropertyInfo primaryKeyProperty = AttributeHelper.GetProperty<T>(typeof(PrimaryKey));
            Dictionary<string, List<int>> index = FindIndex<T>(primaryKeyProperty.Name);
            if (index.ContainsKey(primaryKey))
            {
                return;
            }

            // Persist object
            ObjectPackager packager = new ObjectPackager();
            string jsonPayload = packager.Package(obj);
            int row = ChangeRowCounter(typeof(T).Name, 1);
            table.Cells[row, 1].Value = jsonPayload;

            // Update indizes
            UpdateIndizes(obj, row);
        }

        /// <summary>
        /// Updates the index entry (in a specific index) when an object is updated.
        /// </summary>
        /// <typeparam name="T">Type of object which has been updated</typeparam>
        /// <param name="propertyName">Property which has changed</param>
        /// <param name="key">Property value</param>
        /// <param name="row">New row number</param>
        private static void UpdateIndex<T>(string propertyName, string key, int row) where T : IPersistable
        {
            Dictionary<string, List<int>> index = FindIndex<T>(propertyName);
            if (index.TryGetValue(key, out List<int> values))
            {
                values.Add(row);
            }
            else
            {
                values = new List<int>
                {
                    row
                };
                index.Add(key, values);
            }
            ResetIndex<T>(propertyName, index);
        }

        /// <summary>
        /// Deletes an index entry (in a specific index) when an object is deleted from the database.
        /// </summary>
        /// <typeparam name="T">Type of objects which has been deleted</typeparam>
        /// <param name="propertyName">Property name</param>
        /// <param name="key">Property value</param>
        /// <param name="row">Former row number</param>
        private static void DeleteIndexEntry<T>(string propertyName, string key, int row) where T : IPersistable
        {
            Dictionary<string, List<int>> index = FindIndex<T>(propertyName);
            if (index.TryGetValue(key, out List<int> values))
            {
                if (values.Contains(row))
                {
                    values.Remove(row);
                }

                if (values.Count == 0)
                {
                    index.Remove(key);
                }
            }
            ResetIndex<T>(propertyName, index);
        }

        /// <summary>
        /// Updates all index entries (for all indexed properties) when an object is updated.
        /// </summary>
        /// <typeparam name="T">Type of object which has been updated</typeparam>
        /// <param name="obj">Updated object</param>
        /// <param name="row">Row number of the updated entry</param>
        private static void UpdateIndizes<T>(T obj, int row) where T : IPersistable
        {
            List<PropertyInfo> indexProperties = AttributeHelper.GetPropertiesByAttribute<T>(typeof(Index));
            indexProperties.AddRange(AttributeHelper.GetPropertiesByAttribute<T>(typeof(PrimaryKey)));
            foreach (PropertyInfo indexProperty in indexProperties)
            {
                UpdateIndex<T>(indexProperty.Name, indexProperty.GetValue(obj).ToString(), row);
            }
        }

        /// <summary>
        /// Deletes all index entries (for all indexed properties) when an object is deleted from the database.
        /// </summary>
        /// <typeparam name="T">Type of objects which has been deleted</typeparam>
        /// <param name="obj">Object which has been deleted</param>
        /// <param name="row">Former row number</param>
        private static void DeleteIndexEntries<T>(T obj, int row) where T : IPersistable, new()
        {
            List<PropertyInfo> indexProperties = AttributeHelper.GetPropertiesByAttribute<T>(typeof(Index));
            indexProperties.AddRange(AttributeHelper.GetPropertiesByAttribute<T>(typeof(PrimaryKey)));
            foreach (T followingObject in YieldFollowingObjects<T>(row))
            {
                int oldRow = GetRowNumber<T>(AttributeHelper.GetPrimaryKey(followingObject));
                ResetRowInIndex<T>(followingObject, oldRow, oldRow - 1);
            }

            foreach (PropertyInfo indexProperty in indexProperties)
            {
                DeleteIndexEntry<T>(indexProperty.Name, indexProperty.GetValue(obj).ToString(), row);
            }
        }

        /// <summary>
        /// Yields all objects from a table after the specified row.
        /// </summary>
        /// <typeparam name="T">Type of object to retrieve</typeparam>
        /// <param name="row">Row (exclusive) to follow</param>
        /// <returns><see cref="IPersistable"/></returns>
        private static IEnumerable<T> YieldFollowingObjects<T>(int row) where T : IPersistable, new()
        {
            Range usedRange = FindTable(typeof(T).Name).UsedRange;
            ObjectPackager objectPackager = new ObjectPackager();
            if (usedRange.Rows.Count <= 1)
            {
                yield break;
            }
            for (int i = row + 1; i <= usedRange.Rows.Count; i++)
            {
                yield return objectPackager.Unpackage<T>(usedRange.Cells[i, 1].Value.ToString());
            }
        }

        /// <summary>
        /// Resets the row list in a index.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <param name="obj">Object to update</param>
        /// <param name="oldRow">Old row number</param>
        /// <param name="newRow">New row number</param>
        private static void ResetRowInIndex<T>(T obj, int oldRow, int newRow) where T : IPersistable, new()
        {
            List<PropertyInfo> indexProperties = AttributeHelper.GetPropertiesByAttribute<T>(typeof(Index));
            indexProperties.AddRange(AttributeHelper.GetPropertiesByAttribute<T>(typeof(PrimaryKey)));
            foreach (PropertyInfo indexProperty in indexProperties)
            {
                Dictionary<string, List<int>> index = FindIndex<T>(indexProperty.Name);
                string key = indexProperty.GetValue(obj).ToString();
                if (index.TryGetValue(key, out List<int> values))
                {
                    if (values.Contains(oldRow))
                    {
                        values.Remove(oldRow);
                        values.Add(newRow);
                    }
                }
                ResetIndex<T>(indexProperty.Name, index);
            }
        }

        /// <summary>
        /// Resets the current index with a new one.
        /// </summary>
        /// <typeparam name="T">Type which holds the property</typeparam>
        /// <param name="propertyName">Property name</param>
        /// <param name="index">New index as <see cref="Dictionary{string, List{int}}"/></param>
        private static void ResetIndex<T>(string propertyName, Dictionary<string, List<int>> index)
        {
            string tableName = "Idx_" + GetShortenedHash(typeof(T).Name + propertyName);
            Worksheet ws = FindTable(tableName);
            ws.Cells[1, 1].Value = JsonSerializer.Serialize(index);
        }

        /// <summary>
        /// Finds and returns an (deserialized) object based its primary key.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <param name="primaryKey">Primary key</param>
        /// <returns>Deserialized object</returns>
        public static T Find<T>(string primaryKey) where T : IPersistable, new()
        {
            ObjectPackager objectPackager = new ObjectPackager();
            string rawEntry = GetRawEntry<T>(primaryKey);
            if (rawEntry.Equals("-"))
            {
                return default;
            }
            return objectPackager.Unpackage<T>(rawEntry);
        }

        /// <summary>
        /// Finds and returns an (deserialized) object based its row number.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <param name="row">Row number</param>
        /// <returns>Deserialized object</returns>
        public static T Find<T>(int row) where T : IPersistable, new()
        {
            ObjectPackager objectPackager = new ObjectPackager();
            string rawEntry = GetRawEntry<T>(row);
            return objectPackager.Unpackage<T>(rawEntry);
            
        }

        /// <summary>
        /// Gets the raw JSON string of an entry based on its primary key.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <param name="row">Row number</param>
        /// <returns>Raw JSON string</returns>
        public static string GetRawEntry<T>(string primaryKey) where T : IPersistable, new()
        {
            int rowNumber = GetRowNumber<T>(primaryKey);
            if (rowNumber == -1)
            {
                return "-";
            }
            return GetRawEntry<T>(rowNumber);
        }

        /// <summary>
        /// Gets the raw JSON string of an entry based on its row number.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <param name="row">Row number</param>
        /// <returns>Raw JSON string</returns>
        public static string GetRawEntry<T>(int row) where T : IPersistable, new()
        {
            Worksheet table = GetPersistenceTable<T>();
            string cellValue = table.Cells[row, 1].Value.ToString();
            return cellValue;
        }

        /// <summary>
        /// Get the row number of an entry based on its primary key.
        /// </summary>
        /// <typeparam name="T">Type of object to search</typeparam>
        /// <param name="primaryKey">Primary key</param>
        /// <returns>Row number if found, else -1</returns>
        private static int GetRowNumber<T>(string primaryKey) where T : IPersistable, new()
        {
            Dictionary<string, List<int>> index = FindIndex<T>(AttributeHelper.GetProperty<T>(typeof(PrimaryKey)).Name);
            bool primaryKeyExists = index.TryGetValue(primaryKey, out List<int> rowNumbers);
            if (primaryKeyExists)
            {
                if (rowNumbers.Count == 0)
                {
                    return -1;
                }
                int rowNumber = rowNumbers[0];
                return rowNumber;
            }
            return -1;
        }

        /// <summary>
        /// Yields all entries from a table.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <returns>Yields all entries</returns>
        public static IEnumerable<string> GetAll<T>() where T : IPersistable
        {
            Worksheet table = FindTable(typeof(T).Name);
            if (table == null)
            {
                yield break;
            }
            Range dataColumn = table.UsedRange.Columns[1];
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

        /// <summary>
        /// Yields all entries from a table and deserializes them into objects.
        /// </summary>
        /// <typeparam name="T">Object type</typeparam>
        /// <returns>Yields all objects</returns>
        public static IEnumerable<T> GetAllObjects<T>() where T : IPersistable, new()
        {
            ObjectPackager objectPackager = new ObjectPackager();
            foreach (string jsonPayload in GetAll<T>())
            {
                T newObj = objectPackager.Unpackage<T>(jsonPayload);
                yield return newObj;
            }
        }

        /// <summary>
        /// Gets the index for the specified property.
        /// </summary>
        /// <typeparam name="T">Type which holds the property</typeparam>
        /// <param name="property">Property name</param>
        /// <returns>Index as <see cref="Dictionary{TKey, TValue}"/> but if none exists <c>null</c> is returned.</returns>
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

        /// <summary>
        /// Check if an index for the specified property exists.
        /// </summary>
        /// <typeparam name="T">Type which holds the property</typeparam>
        /// <param name="property">Property name</param>
        /// <returns><c>True</c> if index exists, else <c>false</c></returns>
        public static bool IsIndexed<T>(string property) where T : IPersistable
        {
            string tableName = "Idx_" + GetShortenedHash(typeof(T).Name + property);
            Worksheet indexTable = FindTable(tableName);
            if (indexTable == null)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Gets a hashed representation of a table name.
        /// </summary>
        /// <param name="text"><see cref="string"/> to hash</param>
        /// <returns>Hashed representation of the table name</returns>
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

        /// <summary>
        /// Creates a new index for the specified class property.
        /// </summary>
        /// <typeparam name="T">Type which holds the property</typeparam>
        /// <param name="property">Property name</param>
        /// <returns><c>true</c> if index creation was successful, else <c>false</c></returns>
        public static bool CreateIndex<T>(string property) where T : IPersistable
        {
            Type propertyType = typeof(T).GetProperty(property).PropertyType;

            // Check if type and property match
            if (propertyType == null)
            {
                return false;
            }

            // Check if there is already an index for the specified table and attribute
            if (IsIndexed<T>(property))
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

        /// <summary>
        /// Updates an exisiting entry.
        /// </summary>
        /// <typeparam name="T">Type of object to update</typeparam>
        /// <param name="obj">Object to update</param>
        public static void Update<T>(T obj) where T : IPersistable, new()
        {
            // Replace the existing record
            string primaryKey = AttributeHelper.GetPrimaryKey(obj);
            int rowNumber = GetRowNumber<T>(primaryKey);
            Worksheet dataTable = FindTable(obj.GetType().Name);
            ObjectPackager objectPackager = new ObjectPackager();
            dataTable.Cells[rowNumber, 1].Value = objectPackager.Package(obj);

            // Reorganize indizes
            UpdateIndizes(obj, rowNumber);
        }

        /// <summary>
        /// Deletes an entry from the table and triggers index reorganisation.
        /// </summary>
        /// <typeparam name="T">Type of objects which is deleted.</typeparam>
        /// <param name="obj">Object to delete</param>
        public static void Delete<T>(T obj) where T : IPersistable, new()
        {
            string primaryKey = AttributeHelper.GetPrimaryKey(obj);
            string tableName = obj.GetType().Name;
            int rowNumber = GetRowNumber<T>(primaryKey);

            // Reorganize index
            DeleteIndexEntries(obj, rowNumber);

            // Remove the exisiting record
            Worksheet dataTable = FindTable(tableName);
            dataTable.Rows[rowNumber].EntireRow.Delete();

            // Decrement row counter
            ChangeRowCounter(tableName, -1);
        }
    }
}
