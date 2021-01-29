using System;

namespace ExShift.Mapping
{
    /// <summary>
    /// Attribute for marking a property for which an index should be maintained.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class Index : Attribute
    {
        public Index()
        {
        }
    }
}
