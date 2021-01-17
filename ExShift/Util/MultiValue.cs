using System;

namespace ExShift.Util
{
    [AttributeUsage(AttributeTargets.Property)]
    public class MultiValue : Attribute
    {
        public MultiValue()
        {
        }
    }
}
