using System;

namespace ExShift.Util
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true)]
    public class ForeignKey : Attribute
    {
    }
}
