using ExShift.Util;
using System.Collections.Generic;

namespace ExShift.Tests.Setup
{
    public class PackageTestObject : PackageTestBaseClass
    {
        [Index]
        public int Property { get; set; }
        [PrimaryKey]
        public int DerivedProperty { get; set; }
        [ForeignKey]
        public PackageTestObjectNested NestedObject { get; set; }

        public PackageTestObject(int prop, int dp)
        {
            Property = prop;
            DerivedProperty = dp;
            NestedObject = new PackageTestObjectNested("nested_1");
        } 

        public PackageTestObject()
        {

        }

        public override bool Equals(object obj)
        {
            return obj is PackageTestObject @object &&
                   BaseProperty == @object.BaseProperty &&
                   Property == @object.Property &&
                   DerivedProperty == @object.DerivedProperty &&
                   EqualityComparer<PackageTestObjectNested>.Default.Equals(NestedObject, @object.NestedObject);
        }

        public override int GetHashCode()
        {
            int hashCode = 2054235610;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(BaseProperty);
            hashCode = hashCode * -1521134295 + Property.GetHashCode();
            hashCode = hashCode * -1521134295 + DerivedProperty.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<PackageTestObjectNested>.Default.GetHashCode(NestedObject);
            return hashCode;
        }
    }

    public class PackageTestObjectNested : IPersistable
    {
        [PrimaryKey]
        public string NestedProperty { get; set; }

        public PackageTestObjectNested(string property)
        {
            NestedProperty = property;
        }

        public PackageTestObjectNested()
        {

        }

        public override bool Equals(object obj)
        {
            return obj is PackageTestObjectNested nested &&
                   NestedProperty == nested.NestedProperty;
        }

        public override int GetHashCode()
        {
            return 615921751 + EqualityComparer<string>.Default.GetHashCode(NestedProperty);
        }
    }

    public class PackageTestBaseClass : IPersistable
    {
        [Index]
        public string BaseProperty { get; set; }
        [ForeignKey]
        [MultiValue]
        public List<PackageTestObjectNested> ListOfNestedObjects { get; set; }

        public PackageTestBaseClass()
        {
            BaseProperty = "base_1";
            ListOfNestedObjects = new List<PackageTestObjectNested>
            {
                new PackageTestObjectNested("nested_list_1"),
                new PackageTestObjectNested("nested_list_2"),
                new PackageTestObjectNested("nested_list_3")
            };
        }
    }
}
