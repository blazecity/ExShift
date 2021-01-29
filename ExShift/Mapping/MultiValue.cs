using System;

namespace ExShift.Mapping
{
    [AttributeUsage(AttributeTargets.Property)]
    public class MultiValue : Attribute
    {
        public MultiValue()
        {
        }
    }
}
