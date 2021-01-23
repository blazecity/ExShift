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
        public static List<PropertyInfo> GetPropertiesByAttribute(IPersistable persistable, Type attributeType)
        {
            PropertyInfo[] properties = persistable.GetType().GetProperties();
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

        public static PropertyInfo GetProperty(IPersistable persistable, Type attributeType)
        {
            PropertyInfo[] properties = persistable.GetType().GetProperties();
            foreach (PropertyInfo property in properties)
            {
                if (property.GetCustomAttribute(attributeType) != null)
                {
                    return property;
                }
            }
            return null;
        }

        public static List<PropertyInfo> GetProperties(Type type)
        {
            List<PropertyInfo> list = new List<PropertyInfo>();
            list.AddRange(type.GetProperties());
            return list;
        }

        public static List<PropertyInfo> GetProperties(IPersistable persistable)
        {
            return GetProperties(persistable.GetType());
        }

        public static string GetPrimaryKey(IPersistable obj)
        {
            PropertyInfo primaryKey = GetProperty(obj, typeof(PrimaryKey));
            return primaryKey.GetValue(obj).ToString();
        }
    }
}
