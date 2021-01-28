using ExShift.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Text.Json;
using System.Linq;

namespace ExShift.Util
{
    public class ObjectPackager
    {
        public Dictionary<string, object> Properties { get; }

        public ObjectPackager()
        {
            Properties = new Dictionary<string, object>();
        }

        public string Package(IPersistable obj)
        {
            List<PropertyInfo> properties = new List<PropertyInfo>(obj.GetType().GetProperties());
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
                    ConvertForeignKey(obj, property);
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
            List<PropertyInfo> propertyList = new List<PropertyInfo>(type.GetProperties());
            T newObject = new T();
            foreach (PropertyInfo property in propertyList)
            {
                if (property.GetCustomAttribute<ForeignKey>() != null)
                {
                    MethodInfo findMethod = typeof(ExcelObjectMapper).GetMethod("GetRawEntry", new Type[] { typeof(string) });
                    if (property.GetCustomAttribute<MultiValue>() != null)
                    {
                        Type propertyTypeWithoutGenericType = property.PropertyType.GetGenericTypeDefinition();
                        Type genericType = property.PropertyType.GetGenericArguments()[0];
                        findMethod = findMethod.MakeGenericMethod(genericType);
                        Type listType = propertyTypeWithoutGenericType.MakeGenericType(genericType);
                        object newList = Activator.CreateInstance(listType);

                        JsonElement list = resolvedDict[property.Name];
                        for (int index = 0; index < list.GetArrayLength(); index++)
                        {
                            string json = findMethod.Invoke(null, new string[] { list[index].ToString() }) as string;
                            ((IList)newList).Add(GetGenericMethod("Unpackage", genericType).Invoke(this, new string[] { json }));
                        }
                        property.SetValue(newObject, newList);
                        continue;
                    }
                    string[] parameters = new string[1];
                    string primaryKey = resolvedDict[property.Name].ToString();
                    findMethod = findMethod.MakeGenericMethod(property.PropertyType);
                    parameters[0] = findMethod.Invoke(null, new string[] { primaryKey }) as string;
                    property.SetValue(newObject, GetGenericMethod("Unpackage", property.PropertyType).Invoke(this, parameters));
                    continue;
                }
                JsonElement jsonElement = resolvedDict[property.Name];
                property.SetValue(newObject, ConvertJsonElement(property.PropertyType, jsonElement));
            }
            return newObject;
        }

        public static JsonElement DeserializeTupel(string payload)
        {
            JsonDocument json = JsonDocument.Parse(payload);
            return json.RootElement;
        }

        public static dynamic ConvertJsonElement(Type dataType, JsonElement jsonEl)
        {
            Dictionary<Type, Func<JsonElement, dynamic>> actionTable = new Dictionary<Type, Func<JsonElement, dynamic>>
                {
                    {typeof(int), jsonElement => jsonElement.GetInt32() },
                    {typeof(double), jsonElement => jsonElement.GetDouble() },
                    {typeof(string), jsonElement => jsonElement.GetString() },
                    {typeof(bool), jsonElement => jsonElement.GetBoolean() }
                };
            return actionTable[dataType].Invoke(jsonEl);
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

        private void ConvertForeignKey(IPersistable obj, PropertyInfo property)
        {
            IPersistable nestedObject = property.GetValue(obj) as IPersistable;
            Properties.Add(property.Name, AttributeHelper.GetPrimaryKey(nestedObject));
        }
    }
}
