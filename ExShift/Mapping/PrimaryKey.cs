using System;

namespace ExShift.Mapping
{
    /// <summary>
    /// Attribute for marking a property as primary key.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = true)]
    public class PrimaryKey : Attribute
    {

    }
}
