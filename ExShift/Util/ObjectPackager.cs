using ExShift.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace ExShift.Util
{
    public class ObjectPackager
    {
        private readonly IPersistable obj;
        private readonly ExcelObjectMapper eom;
        public Dictionary<string, object> Properties { get; }

        public ObjectPackager(IPersistable obj, ExcelObjectMapper eom = null)
        {
            this.obj = obj;
            if (eom == null)
            {
                this.eom = new ExcelObjectMapper();
            }
            else
            {
                this.eom = eom;
            }
            Properties = new Dictionary<string, object>();
        }

        public ObjectPackager() 
        { 
        }

        public string Package()
        {
            List<PropertyInfo> properties = AttributeHelper.GetProperties(obj);
            foreach (PropertyInfo property in properties)
            {
                if (property.GetCustomAttribute(typeof(ForeignKey)) != null)
                {
                    if (property.GetCustomAttribute<MultiValue>() != null)
                    {
                        var items = property.GetValue(obj);
                        List<string> foreignKeys = new List<string>();
                        foreach (IPersistable item in items as IEnumerable<IPersistable>)
                        {
                            foreignKeys.Add(AttributeHelper.GetPrimaryKey(item));
                        }
                        Properties.Add(property.Name, foreignKeys);
                        continue;
                    }
                    ConvertForeignKey(property);
                }
                else
                {
                    Properties.Add(property.Name, property.GetValue(obj));
                }
            }
            return JsonSerializer.Serialize(Properties);
        }

        public T Unpackage<T>(string jsonPayload) where T : IPersistable, new()
        {
            Dictionary<string, JsonElement> resolvedDict = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(jsonPayload);
            Type type = typeof(T);
            List<PropertyInfo> propertyList = AttributeHelper.GetProperties(type);
            T newObject = new T();
            foreach (PropertyInfo property in propertyList)
            {
                if (property.GetCustomAttribute<ForeignKey>() != null)
                {
                    if (property.GetCustomAttribute<MultiValue>() != null)
                    {
                        Type genericType = property.PropertyType.GetGenericArguments()[0];
                        Type propertyTypeWithoutGenericType = property.PropertyType.GetGenericTypeDefinition();
                        Type listType = propertyTypeWithoutGenericType.MakeGenericType(genericType);
                        object newList = Activator.CreateInstance(listType);

                        JsonElement list = resolvedDict[property.Name];
                        for (int index = 0; index < list.GetArrayLength(); index++)
                        {
                            string json = eom.Find(genericType.Name, list[index].ToString());
                            ((IList)newList).Add(GetGenericMethod("Unpackage", genericType).Invoke(this, new string[] { json }));
                        }
                        property.SetValue(newObject, newList);
                        continue;
                    }
                    string[] parameters = new string[1];
                    string primaryKey = resolvedDict[property.Name].ToString();
                    parameters[0] = eom.Find(property.PropertyType.Name, primaryKey);
                    property.SetValue(newObject, GetGenericMethod("Unpackage", property.PropertyType).Invoke(this, parameters));
                    continue;
                }
                JsonElement jsonElement = resolvedDict[property.Name];
                Dictionary<Type, Func<string, dynamic>> actionTable = new Dictionary<Type, Func<string, dynamic>>
                {
                    {typeof(int), str => jsonElement.GetInt32() },
                    {typeof(double), str => jsonElement.GetDouble() },
                    {typeof(string), str => jsonElement.GetString() },
                    {typeof(bool), str => jsonElement.GetBoolean() }
                };
                property.SetValue(newObject, actionTable[property.PropertyType].Invoke(jsonElement.GetRawText()));
            }
            return newObject;
        }

        public JsonElement DeserializeTupel(string payload)
        {
            JsonDocument json = JsonDocument.Parse(payload);
            return json.RootElement;
        }

        private MethodInfo GetGenericMethod(string name, Type genericType, Type methodOrigin = null)
        {
            Type origin = methodOrigin;
            if (origin == null)
            {
                origin = GetType();
            }
            return origin.GetMethod(name).MakeGenericMethod(genericType);

        }

        private void ConvertForeignKey(PropertyInfo property)
        {
            IPersistable nestedObject = property.GetValue(obj) as IPersistable;
            Properties.Add(property.Name, AttributeHelper.GetPrimaryKey(nestedObject));
        }

        private void ResolveForeignKey()
        {

        }
    }
}
