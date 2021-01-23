using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExShift.Util
{
    class AttributeHelper
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

        public static string GetPrimaryKey(IPersistable obj)
        {
            Type attributeHelperType = typeof(AttributeHelper);
            MethodInfo getPropertyMethod = attributeHelperType.GetMethod("GetProperty");
            MethodInfo executableMethod = getPropertyMethod.MakeGenericMethod(obj.GetType());
            PropertyInfo primaryKey = executableMethod.Invoke(null,  new object[] { typeof(PrimaryKey) }) as PropertyInfo;
            return primaryKey.GetValue(obj).ToString();
        }
    }
}
