using System;

namespace ExShift.Mapping
{
    /// <summary>
    /// Attibute for marking a property pointing to multiple objects (in lists).
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class MultiValue : Attribute
    {
        public MultiValue()
        {
        }
    }
}
