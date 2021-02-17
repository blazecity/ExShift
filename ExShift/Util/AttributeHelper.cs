using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExShift.Mapping
{
    /// <summary>
    /// This class provides helping functionalities for working with attributes.
    /// </summary>
    public class AttributeHelper
    {
        /// <summary>
        /// Gets a list of properties, which are marked with the specified attribute.
        /// </summary>
        /// <param name="persistable">Object to check (can be default object)</param>
        /// <param name="attributeType">Type of the marking attribute</param>
        /// <returns></returns>
        public static List<PropertyInfo> GetPropertiesByAttribute<T>(Type attributeType)
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            List<PropertyInfo> multiValueProperties = new List<PropertyInfo>();
            foreach (PropertyInfo property in properties)
            {
                if (property.GetCustomAttribute(attributeType, false) != null)
                {
                    multiValueProperties.Add(property);
                }
            }
            return multiValueProperties;
        }

        /// <summary>
        /// Gets the <see cref="PropertyInfo"/> for a property with the specified attribute. 
        /// </summary>
        /// <typeparam name="T">Type to look for the property</typeparam>
        /// <param name="attributeType">Attribute type</param>
        /// <returns><see cref="PropertyInfo"/></returns>
        public static PropertyInfo GetProperty<T>(Type attributeType) where T : IPersistable
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            foreach (PropertyInfo property in properties)
            {
                if (property.GetCustomAttribute(attributeType) != null)
                {
                    return property;
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the primary key of an <see cref="IPersistable"/> object as a string.
        /// </summary>
        /// <param name="obj"><see cref="IPersistable"/> object</param>
        /// <returns>Primary key as string</returns>
        public static string GetPrimaryKey(IPersistable obj)
        {
            Type attributeHelperType = typeof(AttributeHelper);
            MethodInfo getPropertyMethod = attributeHelperType.GetMethod("GetProperty");
            MethodInfo executableMethod = getPropertyMethod.MakeGenericMethod(obj.GetType());
            PropertyInfo primaryKey = executableMethod.Invoke(null,  new object[] { typeof(PrimaryKey) }) as PropertyInfo;
            return primaryKey.GetValue(obj).ToString();
        }


        /// <summary>
        /// Gets the type argument of a generic type.
        /// </summary>
        /// <param name="type">Generic type</param>
        /// <returns>Type argument</returns>
        public static Type GetGenericArgument(Type type)
        {
            if (type.IsGenericType)
            {
                Type[] arguments = type.GetGenericArguments();
                return arguments[0];
            }
            return null;
        }

        /// <summary>
        /// Gets the name of a generic type without the number of number of type arguments.
        /// </summary>
        /// <param name="type">Generic type</param>
        /// <returns>Type name</returns>
        public static string GetGenericType(Type type)
        {
            Regex rgx = new Regex("`");
            return rgx.Split(type.Name)[0];
        }
    }
}
