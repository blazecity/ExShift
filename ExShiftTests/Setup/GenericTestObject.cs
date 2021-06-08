using ExShift.Mapping;
using System.Collections.Generic;

namespace ExShiftTests.Setup
{
    public class GenericTestObjectPT<T> : IPersistable
    {
        [PrimaryKey]
        public int Pk { get; set; }

        [MultiValue]
        public List<T> List { get; set; }
    }

    public class GTO<T> : IPersistable
    {
        [PrimaryKey]
        public int Pk { get; set; }

        [MultiValue]
        [ForeignKey]
        public List<T> List { get; set; }
    }
}
