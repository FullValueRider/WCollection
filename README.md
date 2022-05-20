# WCollection
 A wrapper which enhances the utility of Collection class and removed the oddity of Keys.

This class is a response to a botch job I did in a Stack Overflow answer where I provided a demonstrator as to what could be done by wrapping a Collection.  Unfortunately, the code was written in haste and had many issues about which I am now quite embarrased (https://stackoverflow.com/questions/71085037/redim-preserve-vba-alternatives/71086980?noredirect=1#comment127625279_71086980).

Hence the reason for this class.

The intent of the Wrapper is to make life easier for your typical MS Office user who uses VBA (and also eliminate boilerplate code when coding answers to Advent Of Code problems).  Consequently, the code is functional, not clever.  Its not trying to re-implement the net environment or a Net class.  It just tries to hold the hand of a programmer by doing things that might be considered useful.

## Preamble ##

The class is a self hosted factory.

New instances of the class should be created using the 'Deb' method.  This method name is an 'in-joke' for me as it is short for Debutante.

```
Dim myList as WCollection
Set myList = WCollection.Deb
```

As far as possible, methods return either a value or the instance of Me, to allow chaining of methods.

As of the intial upload there are 146 Rubberduck unit tests, all of which are passing.
The VBA code runs 'as is' in ttwinBasic so I've also included 32 and 64 bit compatible versions of the class in an ActiveX dll.  twinBasic makes this sooooooo easy to do.

### Some terminology/policy ###
<ul>
<li>WCollection stands for Wrapped Collection</li>
<li>WCollections use 1 based indexing
<li>WCollections can be enumerated using 'For Each'
<li>An Enumerable is any VBA entity that can be enumerated using 'For Each'
<li>All method parameters are prefixed by ip, op or iop due to the inadequacies of ByVal and ByRef 
<ul>
<li>ip - Input parameter only. should not be passed ByRef
<li>op - Output parameter only. should not be passed ByVal.  No input value expected
<li>iop - Provides an input value which may be mutated, should not be passed ByVal
</ul>
<li>Variables global to a Class are encapsulated in one of two private UDT.  A Properties UDT with a private variable p, or a State UDT with a variable name of s.
<li>Classes with a PredeclaredId are preferred to Modules.
<li>Classes tend to be written as self factories with methods that allow as much chaining as reasonable.
 
 
</ul>

## List of Methods in Alphabetical Order ##

<p>Add 
<p>AddRange 
<p>AddString 
<p>Clone
<p>Count
<p>Dedup
<p>Dequeue
<p>Enqueque
<p>EnqueueRange
<p>EnqueueString
<p>FilterIt
<p>First
<p>FirstIndex
<p>HasAnyItems
<p>HasItems
<p>HasNoItems
<p>HasOneItem
<p>HoldsItem
<p>IndexOf
<p>Insert
<p>InsertRange
<p>InserttString
<p>IsEquivalent
<p>IsSameOrder
<p>Item 'Default member
<p>Join
<p>LacksItem
<p>Last
<p>LastIndex
<p>LastIndexOf
<p>Mapit
<p>NewEnum
<p>Peek
<p>Pop
<p>Push
<p>PushRange
<p>PushString
<p>ReduceIt
<p>Remove
<p>RemoveAll
<p>RemoveAnyOf
<p>RemoveFirstOf
<p>RemoveLastOf
<p>SetItem  ' a chainable version of Set Item
<p>SetOfCommon
<p>SetOfHostOnly
<p>SetOfInputOnly
<p>SetOfNotCommon
<p>SetOfUnique
<p>Sort
<p>ToArray

## The user interface (API)  ##

In the code examples below, 'myList' is assumed to be a newly created WCollection, unless otherwise specified.

### Deb ###

```
'@Description("Deb (short for Debutante) is a class factory method which returns a new instance of WCollection, optionally populated with a param list, content of an enumerable, or a string (characters))
Public Function Deb(ParamArray ipItems() As Variant) As WCollection
```

If there is only one item in the paramarray then
<ul>
<li>If a string, add the characters of the string as individual items
<li>If an enumerable, add each item in the enumerable as an individial item
<li>If neither of the above, add the item as an individual item
<li>To add a single string or a single enumerable, encapsulate the string or enumerable in an array so that bullet point 2 applies
</ul>

##### Example #####
```
Dim myList as WCollection
' Create a list of 5 items 
' myList.Item(1) = 10
' myList(1) = 10, 
' mylist.FirstItem = 10
Set myList = WCollection.Deb(10, 20, 30, 40, 50)  

' Create a list of characters
' myList.Item(1) = "H"
' myList(1) = "H", 
' mylist.FirstItem = "H"
'Set myList =WCollection.Deb("Hello")

' Create a list from an Enumerable
' myList.Item(1) = 10
' myList(1) = 10, 
' mylist.FirstItem = 10
Dim myColl as collection
Set myColl = New Collection
With myColl
    .Add 10
    .Add 20
    .Add 30
    .Add 40
    .Add 50
End With
Set myList = WCollection.Deb(myColl)

'Create a list containing a single string
' myList.Item(1)="Hello"
' myList.Item(1)  = "Hello"
' myList.FirstItem = "Hello"
set myList =WCollection.Deb(Array("Hello"))

```

### Adding to a WCollection ###

There are three Add methods (but see also the equivalent Enqueue and Push methods)

#### Add ####

```
'@Description("Adds the items in the paramArray to the host instance")
Public Function Add(ParamArray ipItems() As Variant) As WCollection
```

##### Example #####
```
myList.Add 10,20,30,40, 50  ' Function return value is discarded
myList.Add(10).Add(20,30).Add(40,50) ' The return value of the last call to Add is discarded

Dim myList2 as WCollection
set myList2 = myList.Add(10).Add(20,30).Add(40,50) ' myList2 now points to myList
```

#### AddRange  ####

```
'@Description("Add Items from any object that can be enumerated using 'For Each'")
Public Function AddRange(ByVal ipEnumerable As Variant) As WCollection
```
##### Example #####
```
Dim myAL as Object
Set myAL = CreateObject("ArrayList")
With myAL
    .Add "Hello"
    .Add "World"
    .Add "Nice"
    .Add "To"
    .Add "Meet"
    .Add "You"
End With
myList.AddRange myAL  ' Function return value is discarded
```

#### AddString ####

```
'@Description("Add the characters  or substrings of a string as individual items. SUbstrings are generated if a separator is provided")
Public Function AddString _
( _
    ByVal ipString As String, _
    Optional ByVal ipSeparator As String = vbNullString _
) As WCollection
```
If the optional 'ipSeparator' string is provided then the array of substrings produced by VBA.SPlit(ipString,ipSeparator) is added
##### Example #####

To create a list of 11 characters
```
set myList = myList.AddString("Hello World")
```
To create an array of six substrings
```
' Sugar for myList.AddRange VBA.Split(ipString,ipSeparator) but without the flexibility of VBA.Split
myList.AddString "Hello World It's a nice day", " "
```



#### Clone ####

```
'@Description("Returns a a New WCollection that is a shallow copy of the Items in p.Coll")
Public Function Clone() As WCollection
```


#### Count ####

```
'@Description("Returns the number of items in the WCollection")
Public Property Get Count() As Long
```


#### Dedup ####

```

'@Description("Returns a New WCollection containing unique items")
Public Function Dedup() As WCollection

```


#### Dequeue ####
```


```


#### Enqueque ####
```

'@Description("Sugar for Add. Adds Queue terminology")
Public Function Enqueue(ParamArray ipItems() As Variant) As WCollection
```


#### EnqueueRange ####
```

'@Description("Sugar for AddRange. Adds Queue terminology")
Public Function EnqueueRange(ByVal ipEnumerable As Variant) As WCollection
```


#### EnqueueString ####

```

'@Description("Sugar for AddString.  Adds Queue terminology")
Public Function EnqueueString _
( _
    ByVal ipString As String, _
    Optional ByVal ipSeparator As String _
) As WCollection

```

#### FilterIt ####

```

'@Description("Returns a New WCollection based on the result of the ipCompareIt function"
Public Function FilterIt(ByVal ipCompareIt As ICmpIt) As WCollection

```

#### First ####

```


'@Description("Sugar for .Item(FirstIndex), Returns 'Empty' if the WCollection has no items")
Public Function First() As Variant

```



#### FirstIndex ####

```

'@Description("The Lbound of the WCollection, returns -1 is couunt is 0")
Public Function FirstIndex() As Long

```


#### HasAnyItems ####

```

'@Description("True if the WCollection has 1 or more itemp. Sugar for .Count>0")
Public Function HasAnyItems() As Boolean

```


#### HasItems ####

```

'@Description("True if the WCollection has 2 or more itemp. Sugar for .Count>1")
Public Function HasItems() As Boolean

```


#### HasNoItems ####

```

'@Description("True it the WCollection has zero items. SUgar for .Count = 0")
Public Function HasNoItems() As Boolean

```


#### HasOneItem ####

```

'@Description("True if the WCollection only holds one item. SUgar for .Count = 1")
Public Function HasOneItem() As Boolean

```


#### HoldsItem ####

```

'Description("The exists/contains function for WCollection")
Public Function HoldsItem(ByVal ipItem As Variant) As Boolean

```


#### IndexOf ####

```

'@Description("Returns the index of the first found item, If ipRTL is True the negative index is returned)
Public Function Indexof _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0, _
    Optional ByVal ipRTL As Boolean = False _
) As Long

```


#### Insert ####

```

'@Description("Inserts the items in the paramarray into the WCollection")
Public Function Insert(ByVal ipIndex As Long, ParamArray ipItems() As Variant) As WCollection

```


#### InsertRange ####

```


'@Description("Inserts the items in any object that can be enumerated using for each.")
Public Function InsertRange(ByVal ipIndex As Long, ByVal ipItems As Variant) As WCollection

```


#### InserttString ####

```


'@Description("Inserts a string as individual characters, or, if a separator is provided, as substrings")
Public Function InsertString _
( _
    ByVal ipIndex As Long, _
    ByRef ipString As String, _
    Optional ByVal ipSeparator As String = vbNullString _
) As WCollection

```


#### IsEquivalent ####

```

'@Description("True if the Wcollection and ipEnumerable contain the same items, irrespective of order")
Public Function IsEquivalent(ByVal ipEnumerable As Variant) As Boolean

```


#### IsSameOrder ####

```

'@Description("True if the items in the WCollection and ipEnumerable match when enumerated by index")
Public Function IsSameOrder(ByVal ipEnumerable As Variant) As Boolean

```


#### Item 'Default member ####

```


'@Description("Return the item at the specified index, Index may be negative")
'@DefaultMember
Public Property Get Item(ByVal ipIndex As Long) As Variant

'@Description("Adds items to the WCollection.  Accepts Values and Objects. Index May be negative")
Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)

```


#### Join ####

```


'@Description("Simplistic approach to returning the items as a single string, Will error if an item cannot be converted to a string using 'CStr'")
Public Function Join(Optional ByVal ipSeparator As String = ",") As String

```


#### LacksItem ####

```


'@Description("Because I totally dislike "Not HoldsItem" or "Not Exists" etc)
Public Function LacksItem(ByVal ipItem As Variant) As Boolean

```


#### Last ####

```


'@Description("Sugar for .Item(LastIndex). Returns 'Empty' if the WCollection has no items")
Public Function Last() As Variant

```


#### LastIndex ####

```


'@Description("The Ubound of the WCollection, returns -1 if count is 0")
Public Function LastIndex() As Long

```


#### LastIndexOf ####

```


'@Description("Returns the index of the first found item when searching from right to left, If ipRTL is True the negative index is returned)
Public Function LastIndexof _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0, _
    Optional ByVal ipRTL As Boolean = False _
) As Long

```


#### Mapit ####

```

'@Description("Returns a New WCollection where each item is the result of the ipMapIt function")
Public Function MapIt(ByVal ipMapIt As IMapIt) As WCollection

```


#### NewEnum ####

```

'@Description("Allow 'For Each' on the WCollection class")
'@Enumerator
Public Function NewEnum() As IEnumVARIANT

```


#### Peek ####

```


'@Description("Sugar for Item Get. Adds Stack/Queue terminology")
Public Function Peek(Optional ByVal ipIndex As Long = 1) As Variant

```


#### Pop ####

```


'@Description("Returns an array containing one or more items from LastIndex.  Adds Stack terminology")
Public Function Pop(Optional ByVal ipCount As Long = 1) As Variant

```


#### Push ####

```

'@Description("Sugar for Add.  Adds terminology for Stack")
Public Function Push(ParamArray ipItems() As Variant) As WCollection

```


#### PushRange ####

```


'@Description("Sugar for AddRange.  Adds terminology for Stack")
Public Function PushRange(ByVal ipEnumerable As Variant) As WCollection

```


#### PushString ####

```


'@Description("Sugar for AddString.  Adds Stack Terminology")
Public Function PushString _
( _
    ByVal ipString As String, _
    Optional ByVal ipSeparator As String _
) As WCollection

```


#### ReduceIt ####

```


'@Description("Returns a single value calculated by the ipReduceIt function")
Public Function ReduceIt(ByVal ipReduceIt As IReduceIt) As Variant

```


#### Remove ####

```


'@Description("Deletes One or more consecutive items from the WCollection")
Public Function Remove _
( _
    ByVal ipStartIndex As Long, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection

```


#### RemoveAll ####

```


'@Description("Returns a new empty collection, does not delete items from the current instance, if ipEndIndex is not supplied removes a single item")
Public Function RemoveAll() As WCollection

```


#### RemoveAnyOf ####

```


'@("Description("Removes all instances of the item from the specified range")
Public Function RemoveAnyOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection

```


#### RemoveFirstOf ####

```


'@Description("Removes the first found 'Item' from the specified range when searching from Left to Right")
Public Function RemoveFirstOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection

```


#### RemoveLastOf ####

```


'@Description(Removes the first item in the specified range when searching from Right to Left")
Public Function RemoveLastOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection

```


#### SetItem  ' a chainable version of Set Item ####

```


'@Description("Equivalent of Item(x)=xx, but returns the instance of Me to allow chaining")
Public Function SetItem(ByVal ipIndex As Long, ByVal ipItem As Variant) As WCollection

```


#### SetOfCommon ####

```


'@Description("Returns a new WCollection of all values in both WCollection and ipEnumerable")
Public Function SetOfCommon(ByVal ipEnumerable As Variant) As WCollection

```


#### SetOfHostOnly ####

```


'@Description("Returns a new WCollection of all values in WCollection only")
Public Function SetOfHostOnly(ByVal ipEnumerable As Variant) As WCollection

```


#### SetOfInputOnly ####

```


'@Description("Returns a new WCollection of all values in WCollection only")
Public Function SetOfInputOnly(ByVal ipEnumerable As Variant) As WCollection

```


#### SetOfNotCommon ####

```


'@Description("Returns a new WCollection of all values in WCollection only or ipEnumerable only")
Public Function SetOfNotCommon(ByRef ipEnumerable As Variant) As WCollection

```


#### SetOfUnique #####

```


'@Description("Returns a new WCollection of all unique values in WCollection and ipEnumerable")
Public Function SetOfUnique(ByRef ipEnumerable As Variant) As WCollection

```


#### Sort ####

```


'@Description("Sorts in ascending numeric or alphabetical order")
Public Function Sort() As WCollection
'QuickSort3 from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2909260#post2909260

```


#### ToArray ####

```


'@Description("Returns the content of the intenal collection as an array, Optional parameters allow a subsection of the collection to be selected.")
Public Function ToArray _
( _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As Variant

```

