using System;

namespace ExShift.Mapping
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true)]
    public class ForeignKey : Attribute
    {
    }
}
