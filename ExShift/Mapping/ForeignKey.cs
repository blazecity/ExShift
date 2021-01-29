using System;

namespace ExShift.Mapping
{
    /// <summary>
    /// Attribute for marking a property with nested object.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = true)]
    public class ForeignKey : Attribute
    {
    }
}
