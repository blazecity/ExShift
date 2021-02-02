# ExShift

ExShift is a OR mapping tool for storing objects in Excel sheets.

## How does it work?

Let's have a look at this example.


```
public class ClassA : IPersistable
{
  [PrimaryKey]
  public string PkA { get; set; }
  public string AnotherPropertyInA { get; set; }
  [Foreign Key]
  public ClassB ForeignObject { get; set; }
  
  public ClassA
  {
  }
}

public class ClassB : IPersistable
{
  [PrimaryKey]
  public int PkB { get; set; }
  [Index]
  public string AnotherPropertyInB { get; set; }
  
  public classB
  {
  }
}
```

```
ClassB b1 = new ClassB { PkB = 1, AnotherPropertyInB = "B Property" };
ClassA a1 = new ClassA { PkA = "primaryKeyA", AnotherPropertyInA = "A Property", ForeignObject = b1};
```

### General

All classes, for which objects will be persisted, need to have the following things:
- They need to implemented the `IPersistable` interface.
- One and only one property needs to be marked with the `PrimaryKey` attribute.
- Lists need to marked with `MultiValue` attribute.

Optionally, you can mark a property with the `Index` attribute. Then an index will be created automatically.

### Persting objects

```ExcelObjectMapper.Persist(b1)
ExcelObjectMapper.Persist(a1)
```
It is important to persist the Object `b1` first. The objects will be serialized into a **JSON** string. However for nested objects only a reference (via primary key) is stored.

Table for ClassA (Excel sheet):

| A                                                                                 |
|-----------------------------------------------------------------------------------|
| {"PkA": "primaryKeyA", "AnotherPropertyInA": "A Property", "ForeignObject": "b1"} |

Table for ClassB (Excel sheet):

| A                                             |
|-----------------------------------------------|
| {"PkB": 1, "AnotherPropertyInA":"B Property"} |

### Update
```ExcelObjectMapper.Update(IPersistable)```

### Delete
```ExcelObjectMapper.Delete(IPersistable)```

### Searching objects

For searching objects you have the two following possibilites:
- `ExcelObjectMapper.Find<T>(string primaryKey)` method
- `Query` class

#### How to use the `Query` class

For this let's take the example from the begining and let's say we want to find the object of `ClassA`, we persisted:

```
List<ClassA> resultList = Query<ClassA>.Select()
                                       .Run();
```

This would return all stored objects of `ClassA`. But you can also use the `Where(string)`, `And(string)` and `Or(string)` methods:

```
List<ClassA> resultList = Query<ClassA>.Select()
                                       .Where("PkA = 'primaryKeyA'")
                                       .And("AnotherPropertyInA = 'A Property'")
                                       .Run();
```

### Indizes

When a property is marked with the `Index` attribute, an index is created. This also automatically applies to the primary key properties. An index is simply a `Dictionary`, where the keys are the property values and the values of the dictionary are lists of integers, which point to the corresponding rows.

### Restrictions
- Currently it is not possible to search with the nested objects (similiar to `JOIN` in SQL). This is planned for a future release.
- Also you cannot change the primary key attribute once it has been set and persisted. The problem is the update of the foreign keys. But this will be implemented as well.
