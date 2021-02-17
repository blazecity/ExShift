using ExShift.Mapping;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Text.Json;

namespace ExShift.Mapping
{
    /// <summary>
    /// This class is used for serializing objects into a raw string and vice versa.
    /// </summary>
    public class ObjectPackager
    {
        private readonly Dictionary<string, object> Properties;

        /// <summary>
        /// Constructor for a new <c>ObjectPackager</c> object.
        /// </summary>
        public ObjectPackager()
        {
            Properties = new Dictionary<string, object>();
        }

        /// <summary>
        /// Takes an <see cref="IPersistable"/> object and serializes it into a JSON string.
        /// <para>
        /// Nested objects will won't be serialized automatically. This has to be done first.
        /// Then in the parent object only the foreign key will be serialized.
        /// </para>
        /// </summary>
        /// <param name="obj"><see cref="IPersistable"/></param>
        /// <returns>JSON string</returns>
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

        /// <summary>
        /// Deserializes a JSON string into an object.
        /// <para>
        /// Note that also nested objects will also be deserialized and therefore a 
        /// reference is set pointing to it.
        /// </para>
        /// </summary>
        /// <typeparam name="T">Type of the object to be deserialized (highest level)</typeparam>
        /// <param name="jsonPayload">Serialized object (JSON)</param>
        /// <returns>Deserialized object</returns>
        public T Unpackage<T>(string jsonPayload) where T : IPersistable, new()
        {
            if (string.IsNullOrEmpty(jsonPayload) || string.IsNullOrWhiteSpace(jsonPayload) || jsonPayload == "-")
            {
                return default;
            }
            Dictionary<string, JsonElement> resolvedDict = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(jsonPayload);
            Type type = typeof(T);
            List<PropertyInfo> propertyList = new List<PropertyInfo>(type.GetProperties());
            T newObject = new T();
            foreach (PropertyInfo property in propertyList)
            {
                if (property.GetCustomAttribute<ForeignKey>() != null || property.GetCustomAttribute<MultiValue>() != null)
                {
                    MethodInfo findMethod = typeof(ExcelObjectMapper).GetMethod("GetRawEntry", new Type[] { typeof(string) });
                    if (property.GetCustomAttribute<MultiValue>() != null)
                    {
                        Type propertyTypeWithoutGenericType = property.PropertyType.GetGenericTypeDefinition();
                        Type genericType = property.PropertyType.GetGenericArguments()[0];
                        Type listType = propertyTypeWithoutGenericType.MakeGenericType(genericType);
                        object newList = Activator.CreateInstance(listType);
                        JsonElement list = resolvedDict[property.Name];
                        if (!genericType.IsPrimitive)
                        {
                            findMethod = findMethod.MakeGenericMethod(genericType);
                            for (int index = 0; index < list.GetArrayLength(); index++)
                            {
                                string json = findMethod.Invoke(null, new string[] { list[index].ToString() }) as string;
                                ((IList)newList).Add(GetGenericMethod("Unpackage", genericType).Invoke(this, new string[] { json }));
                            }
                        }
                        else
                        {
                            for (int index = 0; index < list.GetArrayLength(); index++)
                            {
                                ((IList)newList).Add(ConvertJsonElement(genericType, list[index]));
                            }
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

        /// <summary>
        /// Turns a JSON string into a <see cref="JsonElement"/> object.
        /// </summary>
        /// <param name="payload">JSON string</param>
        /// <returns><see cref="JsonElement"/></returns>
        public static JsonElement DeserializeTupel(string payload)
        {
            JsonDocument json = JsonDocument.Parse(payload);
            return json.RootElement;
        }

        /// <summary>
        /// Takes a <see cref="JsonElement"/> and returns its value.
        /// </summary>
        /// <param name="dataType">Data type of the value</param>
        /// <param name="jsonEl"><see cref="JsonElement"/></param>
        /// <returns></returns>
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

        /// <summary>
        /// Creates a generic method (helper method).
        /// </summary>
        /// <param name="name">Method name</param>
        /// <param name="genericType">Generic type</param>
        /// <param name="methodOrigin">Class with the given method</param>
        /// <returns><see cref="MethodInfo"/></returns>
        private MethodInfo GetGenericMethod(string name, Type genericType, Type methodOrigin = null)
        {
            Type origin = methodOrigin;
            if (origin == null)
            {
                origin = GetType();
            }
            return origin.GetMethod(name).MakeGenericMethod(genericType);
        }

        /// <summary>
        /// Adds the foreign key into the Properties <see cref="Dictionary{TKey, TValue}"/>
        /// instead of the whole (nested) object (helper method).
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="property"></param>
        private void ConvertForeignKey(IPersistable obj, PropertyInfo property)
        {
            IPersistable nestedObject = property.GetValue(obj) as IPersistable;
            Properties.Add(property.Name, AttributeHelper.GetPrimaryKey(nestedObject));
        }
    }
}
