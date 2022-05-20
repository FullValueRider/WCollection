VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A wrapper for the collection object to add flexibility"
'@IgnoreModule SuspiciousPredeclaredInstanceAccess
' This class has been developed using RubberDuck
' and consequently contains Rubberduck annotations
'@PredeclaredId
'@Exposed
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit


'@ModuleDescription("A wrapper for the collection object to add flexibility")

'Chaining
' Many methods of this class return the instance of Me.
' This allows methods of the WCOllection class to be chained.
' e.g.set  myList = WCollection.Deb.Add(1,2,3,4,5).AddRange(Array(1,2,3,4,5).Add(Array(1,2,3,4,5))

' Class UDT
' I typically use two private UDT in a Class Definition
' UDT State (variable s) for module level variables that describe the
' state of the class instance
' UDT Properties (variable p) for the backing variables for Class
' Properties
' The use of these enhances intellisense within the class/module/form
' Allows simpler exporting of state or properties if required.
' For WCollection we only need Properties

Const WCOLLECTION_ERROR                         As Long = 17 + vbObjectError
Const MY_NAME                                   As String = "WCollection"

Private Type Properties
    
    Coll                                        As Collection
    
End Type

Private p                                       As Properties

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Initialisation stuff
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Sub Class_Initialize()
    ' WCollection classes are initialised using WCollection.Deb, not New
    ' WCollection but New **IS** used internally as part of the factory method
    ' Don't blame me if you New a WCollection and things don't work as expected
End Sub

'@Description("Deb (short for Debutante) is a class factory method which returns a new instance of WCollection, optionally populated with a param list, content of an enumerable, or a string (characters))
Public Function Deb(ParamArray ipItems() As Variant) As WCollection
    ' This call is on the PredeclaredId of WCollection
    
   ' This is the only time New is required for WCollection
    With New WCollection
   
        Select Case True
     
            Case ArrayHasNoItems(CVar(ipItems))

                Set Deb = .ConstructInstance
         
            '2 or more items
            Case ArrayHasItems(CVar(ipItems))
         
                Set Deb = .ConstructInstanceFromEnumerable(CVar(ipItems))
            
            ' From here on we know ParamArray only has One Item
            Case IsString(ipItems(0))
         
                ' Because of the nature of ParamArray we cannot include the optional parameter
                ' for constructing as substrings.  See AddString method.
                Set Deb = .ConstructInstanceFromString(ipItems(0))
                
                
            Case VBA.IsArray(ipItems(0))
                
                ' for multi dimension arrays VBA Enumerates in Column Major Order
                Set Deb = .ConstructInstanceFromEnumerable(ipItems(0))
                    
                    
            Case IsEnumerableObject(ipItems(0))
                                 
                Set Deb = .ConstructInstanceFromEnumerable(ipItems(0))
                

            Case Else
            
                Err.Raise WCOLLECTION_ERROR, MY_NAME, "Unexpected constructor item"
             
             
         End Select
             
    End With
 
End Function


'@Description("Worker for the Deb function which return the newly created instance")
Friend Function ConstructInstance() As WCollection
Attribute ConstructInstance.VB_Description = "Worker for the Deb function which return the newly created instance"
    ' This call is in the new instance of WCollection created by Deb
    Set p.Coll = New Collection
    
    Set ConstructInstance = Me

End Function

'@Description("Worker for the Deb function which return the newly created instance")
Friend Function ConstructInstanceFromEnumerable(ByVal ipEnumerable As Variant) As WCollection
Attribute ConstructInstanceFromEnumerable.VB_Description = "Worker for the Deb function which return the newly created instance"
    ' This call is in the new instance of WCollection created by Deb
    Set p.Coll = New Collection
    
    Set ConstructInstanceFromEnumerable = AddRange(ipEnumerable)
    
End Function

'@Description("Worker for the Deb function which return the newly created instance")
Friend Function ConstructInstanceFromString(ByVal ipString As Variant) As WCollection
Attribute ConstructInstanceFromString.VB_Description = "Worker for the Deb function which return the newly created instance"
    ' This call is in the new instance of WCollection created by Deb
    Set p.Coll = New Collection
    
    Set ConstructInstanceFromString = AddString(ipString)
    
End Function

'@Description("Allow for Each on the WCollection class")
'@Enumerator
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_Description = "Allow for Each on the WCollection class"
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = p.Coll.[_NewEnum]
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' API
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

'@Description("Adds the items in the paramArray to p.coll")
Public Function Add(ParamArray ipItems() As Variant) As WCollection
Attribute Add.VB_Description = "Adds the items in the paramArray to p.coll"

    Dim myItem As Variant
    For Each myItem In ipItems
        p.Coll.Add myItem
    Next
    
    Set Add = Me
    
End Function

'@Description("Add Items from any object that can be enumerated using for each")
Public Function AddRange(ByVal ipEnumerable As Variant) As WCollection
Attribute AddRange.VB_Description = "Add Items from any object that can be enumerated using for each"
    ' This method will hapily add all items from a multidimensional array
    ' Items can be retrieved in the Array order by calculating
    ' the  Index using the array address to calcumate the column major order address
    GuardForIsNotEnumerable ipEnumerable
    
    Dim myItem As Variant
    For Each myItem In ipEnumerable
        p.Coll.Add myItem
    Next
    
    Set AddRange = Me
    
End Function

'@Description("Add the characters  or substtrings of string as individual items. SUbstrings are generated if a separator is provided")
Public Function AddString _
( _
    ByVal ipString As String, _
    Optional ByVal ipSeparator As String = vbNullString _
) As WCollection
Attribute AddString.VB_Description = "Add the characters  or substtrings of string as individual items. SUbstrings are generated if a separator is provided"
    
    GuardForIsNotString ipString
    
    If ipSeparator = vbNullString Then
    
        Set AddString = AddStringAsCharacters(ipString)
        
        
    Else
    
        Set AddString = AddStringAsSubstrings(ipString, ipSeparator)
        
        
    End If
    
End Function
'@Description("Worker for the AddString method")
Private Function AddStringAsCharacters(ByVal ipString As String) As WCollection
Attribute AddStringAsCharacters.VB_Description = "Worker for the AddString method"

    Dim myindex As Long
    For myindex = 1 To Len(ipString)
    
        p.Coll.Add VBA.Mid$(ipString, myindex, 1)
        
        
    Next
    
    Set AddStringAsCharacters = Me
    
End Function

'@Description("Worker for the AddString method")
Private Function AddStringAsSubstrings _
( _
    ByVal ipString As String, _
    ByVal ipSeparator As String _
) As WCollection
Attribute AddStringAsSubstrings.VB_Description = "Worker for the AddString method"
    Set AddStringAsSubstrings = AddRange(VBA.Split(ipString, ipSeparator))
End Function

'@Description("Returns a a New WCollection that is a shallow copy of the Items in p.Coll")
Public Function Clone() As WCollection
Attribute Clone.VB_Description = "Returns a a New WCollection that is a shallow copy of the Items in p.Coll"
    Set Clone = WCollection.Deb.AddRange(p.Coll)
End Function

'@DefaultMember
Public Property Get Item(ByVal ipIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0

    GuardForInvalidIndex ipIndex
    
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    
    Assign Item, p.Coll.Item(myindex)
    
End Property

'@Description("Adds items to the WCollection.  Accepts Values and Objects
Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)

    GuardForInvalidIndex ipIndex
    
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    
    p.Coll.Add ipItem, after:=myindex
    p.Coll.Remove myindex
    
    
End Property


'@Description("Equivalent of Item(x)=xx, but returns the instance of Me to allow chaining")
Public Function SetItem(ByVal ipIndex As Long, ByVal ipItem As Variant) As WCollection
Attribute SetItem.VB_Description = "Equivalent of Item(x)=xx, but returns the instance of Me to allow chaining"

    GuardForInvalidIndex ipIndex
    
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    
    p.Coll.Add ipItem, after:=myindex
    p.Coll.Remove myindex
    
    Set SetItem = Me
    
End Function


'Description("The exists/contains function for WCollection")
Public Function HoldsItem(ByVal ipItem As Variant) As Boolean
    HoldsItem = True
    Dim myItem As Variant
    For Each myItem In p.Coll
        If myItem = ipItem Then Exit Function
    Next
    HoldsItem = False
End Function

'@Description("Because I totally dislike "Not HoldsItem" or "Not Exists" etc)
Public Function LacksItem(ByVal ipItem As Variant) As Boolean
    LacksItem = Not HoldsItem(ipItem)
End Function


'@Description("Simplistic approach to returning the items as a single string, Will error if an item cannot be converted to a string using 'CStr'")
'@Ignore ShadowedDeclaration
Public Function Join(Optional ByVal ipSeparator As String = ",") As String
Attribute Join.VB_Description = "Simplistic approach to returning the items as a single string, Will error if an item cannot be converted to a string using 'CStr'"

    Join = vbNullString
    If p.Coll.Count = 0 Then Exit Function
    
    Dim myStr As String
    Dim myItem As Variant
    For Each myItem In p.Coll
    
        If Len(myStr) = 0 Then
            myStr = CStr(myItem)
        Else
            myStr = myStr & ipSeparator & CStr(myItem)
        End If
        
    Next
    
    Join = myStr

End Function


'@Description("Reverses the order of items")
Public Function Reverse() As WCollection
Attribute Reverse.VB_Description = "Reverses the order of items"

    If HasNoItems Then
        Set Reverse = WCollection.Deb
        Exit Function
    End If
    
    Dim myW As WCollection
    Set myW = WCollection.Deb
    Dim myindex As Long
    For myindex = LastIndex To FirstIndex Step -1
        '@Ignore FunctionReturnValueDiscarded
        myW.Add p.Coll.Item(myindex)
    Next
    
    Set Reverse = myW
    
End Function


'@Description("True if the WCollection has 2 or more itemp. Sugar for .Count>1")
Public Function HasItems() As Boolean
Attribute HasItems.VB_Description = "True if the WCollection has 2 or more itemp. Sugar for .Count>1"
    HasItems = p.Coll.Count > 1
End Function

'@Description("True it the WCollection has zero items. SUgar for .Count = 0")
Public Function HasNoItems() As Boolean
Attribute HasNoItems.VB_Description = "True it the WCollection has zero items. SUgar for .Count = 0"
    HasNoItems = p.Coll.Count = 0
End Function

'@Description("True if the WCollection has 1 or more itemp. Sugar for .Count>0")
Public Function HasAnyItems() As Boolean
Attribute HasAnyItems.VB_Description = "True if the WCollection has 1 or more itemp. Sugar for .Count>0"
    HasAnyItems = p.Coll.Count > 0
End Function

'@Description("True if the WCollection only holds one item. SUgar for .Count = 1")
Public Function HasOneItem() As Boolean
Attribute HasOneItem.VB_Description = "True if the WCollection only holds one item. SUgar for .Count = 1"
    HasOneItem = p.Coll.Count = 1
End Function

'@Description("Returns the index of the first found item, If ipLTR is false the negative index is returned)
Public Function Indexof _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0, _
    Optional ByVal ipRTL As Boolean = False _
) As Long

    GuardForInvalidIndex ipStartIndex
    GuardForInvalidIndex ipEndIndex
    
    Dim myStartIndex As Long
    myStartIndex = ResolveStartIndex(ipStartIndex)
    
    
    Dim myEndIndex As Long
    '@Ignore IIfSideEffect
    myEndIndex = ResolveEndIndex(ipEndIndex)
    
    ' In the case of negative indeces we always return
    ' items from Left To Right
    ' so negative indexes should resolve to positive indexes
    ' where firstindex is less than EndIndex
    ' Consequently, whilst negative indexes are allow, thet
    ' do not reverse the order
    ' in which items are rettireved or added to the WCollection
    GuardForInvalidIndexOrder myStartIndex, myEndIndex
    
    Indexof = -1
    
    Dim mySearchIndex As Long
    For mySearchIndex = myStartIndex To myEndIndex
    
        ' we keep a copy of myIndex because if the
        ' loop completes normall myIndex will be 1 greater
        ' than myEndIndex
        
        
        ' The comparison below is a bit simplistic but suffices
        ' for the intended use of this class (maniplulating excel data)
        If ipItem = p.Coll.Item(mySearchIndex) Then
            
            If ipRTL Then mySearchIndex = mySearchIndex - p.Coll.Count - 1
            Indexof = mySearchIndex
            Exit Function
            
        End If
        
    Next
    
End Function

'@Description("Returns the index of the first found item when searching from right to left, If ipLTR is false the negative index is returned)
Public Function LastIndexof _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0, _
    Optional ByVal ipRTL As Boolean = False _
) As Long
    Dim myStartIndex As Long
    myStartIndex = ResolveStartIndex(ipStartIndex)
    
    Dim myEndIndex As Long
    '@Ignore IIfSideEffect
    myEndIndex = ResolveEndIndex(ipEndIndex)
    
    LastIndexof = -1
    
    Dim mySearchIndex As Long
    For mySearchIndex = myEndIndex To myStartIndex Step -1
        ' The comparison below is a bit simplistic but suffices
        ' for the intended use of this class (maniplulating excel data)
        If ipItem = p.Coll.Item(mySearchIndex) Then
            
            If ipRTL Then mySearchIndex = mySearchIndex - p.Coll.Count - 1
            LastIndexof = mySearchIndex
            Exit Function
            
        End If
    Next
    
End Function

'@Description("Returns the content of the intenal collection as an array, Optional parameters allow a subsection of the collection to be selected.")
Public Function ToArray _
( _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As Variant
Attribute ToArray.VB_Description = "Returns the content of the intenal collection as an array, Optional parameters allow a subsection of the collection to be selected."
    
    GuardForInvalidIndex ipStartIndex
    GuardForInvalidIndex ipEndIndex
    
'    Dim myCount As Long
'    myCount = p.Coll.Count
    
    Dim myStartIndex As Long
    myStartIndex = ResolveStartIndex(ipStartIndex)
    
    Dim myEndIndex As Long
    myEndIndex = ResolveEndIndex(ipEndIndex)

    GuardForInvalidIndexOrder myStartIndex, myEndIndex
    
    ' Remember that startindex may not be 1 so
    ' we need to adjust end index accordingly
    Dim myArray As Variant
    ReDim myArray(0 To myEndIndex - myStartIndex)
    
    Dim myCollIndex As Long
    For myCollIndex = myStartIndex To myEndIndex
        
        '@Ignore UnassignedVariableUsage
        Assign myArray(myCollIndex - myStartIndex), p.Coll.Item(myCollIndex)
            
    Next
    
    '@Ignore UnassignedVariableUsage
    ToArray = myArray
End Function


'@Description("Inserts the items in the paramarray into the WCollection")
Public Function Insert(ByVal ipIndex As Long, ParamArray ipItems() As Variant) As WCollection
Attribute Insert.VB_Description = "Inserts the items in the paramarray into the WCollection"

    Set Insert = Me
    If ArrayHasNoItems(CVar(ipItems)) Then Exit Function
    
    GuardForInvalidIndex ipIndex
    
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    
    Dim myItem As Variant
    For Each myItem In ipItems
    
        p.Coll.Add myItem, before:=myindex
        myindex = myindex + 1
        
    Next
    
    Set Insert = Me
    
End Function

'@Description("Inserts the items in any object that can be enumerated using for each.")
Public Function InsertRange(ByVal ipIndex As Long, ByVal ipItems As Variant) As WCollection
Attribute InsertRange.VB_Description = "Inserts the items in any object that can be enumerated using for each."

    GuardForInvalidIndex ipIndex
    GuardForIsNotEnumerable ipItems
    
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    
    Dim myItem As Variant
    For Each myItem In ipItems
    
        p.Coll.Add myItem, before:=myindex
        myindex = myindex + 1
        
    Next
    
    Set InsertRange = Me
    
End Function

'@Description("Inserts a string as individual characters, or, if a separator is provided, as substrings")
Public Function InsertString _
( _
    ByVal ipIndex As Long, _
    ByRef ipString As String, _
    Optional ByVal ipSeparator As String = vbNullString _
) As WCollection
Attribute InsertString.VB_Description = "Inserts a string as individual characters, or, if a separator is provided, as substrings"

    GuardForInvalidIndex ipIndex
    
    Dim myLen As Long
    myLen = VBA.Len(ipString)
    
    If myLen = 0 Then
    
        Set InsertString = Me
        Exit Function
        
    End If
    
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    
    If VBA.Len(ipSeparator) = 0 Then
    
        Set InsertString = InsertStringAsCharacters(myindex, ipString)
        
        
    Else
    
        Set InsertString = InsertStringAsSubstrings(myindex, ipString, ipSeparator)
        
        
    End If
    
End Function

'@Description("Worker for method InsertString")
Private Function InsertStringAsCharacters _
( _
    ByVal ipIndex As Long, _
    ByVal ipString As String _
) As WCollection
Attribute InsertStringAsCharacters.VB_Description = "Worker for method InsertString"
    
    Dim myindex As Long
    myindex = ipIndex
    
    Dim myStringIndex As Long
    For myStringIndex = 1 To VBA.Len(ipString)
    
        '@Ignore UntypedFunctionUsage
        p.Coll.Add VBA.Mid(ipString, myStringIndex, 1), before:=myindex
        myindex = myindex + 1
        
        
    Next
    
    Set InsertStringAsCharacters = Me
    
End Function


'@("Worker for method InsertString")
Private Function InsertStringAsSubstrings _
( _
    ByVal ipIndex As Long, _
    ByVal ipString As String, _
    ByVal ipSeparator As String _
) As WCollection
    Set InsertStringAsSubstrings = InsertRange(ipIndex, VBA.Split(ipString, ipSeparator))
End Function

'@Description("Deletes One or more consecutive items from the WCollection")
Public Function Remove _
( _
    ByVal ipStartIndex As Long, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection
Attribute Remove.VB_Description = "Deletes One or more consecutive items from the WCollection"
    
    GuardForInvalidIndex ipStartIndex
    
    Dim myStartIndex As Long
    myStartIndex = ResolveStartIndex(ipStartIndex)
    
    ' if ipend index = 0 we only delete  one item, the item
    ' at myStartIndex
    Dim myEndIndex As Long
    If ipEndIndex = 0 Then
    
        myEndIndex = myStartIndex
        
    Else
    
        
        myEndIndex = ResolveEndIndex(ipEndIndex)
        
    End If
    
    GuardForInvalidIndexOrder myStartIndex, myEndIndex
    
    Dim myindex As Long
    ' The for loop is used as a counter only
    For myindex = myEndIndex To myStartIndex Step -1
        'If myStartIndex > p.Coll.Count Then Exit Function
        
        p.Coll.Remove myindex
        
    Next
    
    Set Remove = Me
    
End Function

'@Description("Returns a new empty collection, does not delete items from the current instance")
Public Function RemoveAll() As WCollection
Attribute RemoveAll.VB_Description = "Returns a new empty collection, does not delete items from the current instance"
    Set p.Coll = New Collection
    Set RemoveAll = Me
End Function

'@Description("Removes the first found 'Item' from the specified range when searching from Left to Right")
Public Function RemoveFirstOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection
Attribute RemoveFirstOf.VB_Description = "Removes the first found 'Item' from the specified range when searching from Left to Right"
    
    Dim myindex As Long
    myindex = Indexof(ipItem, ipStartIndex, ipEndIndex)
    If myindex > 0 Then p.Coll.Remove myindex
    Set RemoveFirstOf = Me
    
End Function

'@Description(Removes the first item in the specified range when searching from Right to Left")
Public Function RemoveLastOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection

    Dim myindex As Long
    myindex = LastIndexof(ipItem, ipStartIndex, ipEndIndex)
    If myindex > 0 Then p.Coll.Remove myindex
    Set RemoveLastOf = Me
    
End Function

'@("Description("Removes all instances of the item from the specified range")
Public Function RemoveAnyOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStartIndex As Long = 0, _
    Optional ByVal ipEndIndex As Long = 0 _
) As WCollection
    
    GuardForInvalidIndex ipStartIndex
    GuardForInvalidIndex ipEndIndex
    
    Dim myStartIndex As Long
    myStartIndex = ResolveStartIndex(ipStartIndex)
    
    Dim myEndIndex As Long
    'If myEndIndex = 0 Then myEndIndex = p.Coll.Count
    myEndIndex = ResolveEndIndex(ipEndIndex)
    
    GuardForInvalidIndexOrder myStartIndex, myEndIndex
    
    Dim myindex As Long
    For myindex = myEndIndex To myStartIndex Step -1
        
        If p.Coll.Item(myindex) = ipItem Then p.Coll.Remove myindex
        
    Next
    
    Set RemoveAnyOf = Me
    
End Function

'@Description("The Lbound of the WCollection, returns -1 is couunt is 0")
Public Function FirstIndex() As Long
Attribute FirstIndex.VB_Description = "The Lbound of the WCollection, returns -1 is couunt is 0"
    FirstIndex = IIf(p.Coll.Count = 0, -1, 1)
End Function

'@Description("The Ubound of the WCollection, returns -1 if count is 0")
Public Function LastIndex() As Long
Attribute LastIndex.VB_Description = "The Ubound of the WCollection, returns -1 if count is 0"
    '@Ignore IIfSideEffect
    LastIndex = IIf(p.Coll.Count = 0, -1, p.Coll.Count)
End Function

'@Description("Returns the number of items in the WCollection")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the WCollection"
    Count = p.Coll.Count
End Property

'@Description("Sugar for .Item(FirstIndex), Returns 'Empty' if the WCollection has no items")
Public Function First() As Variant
Attribute First.VB_Description = "Sugar for .Item(FirstIndex), Returns 'Empty' if the WCollection has no items"
    First = Empty
    If p.Coll.Count = 0 Then Exit Function
    Assign First, p.Coll.Item(FirstIndex)
End Function

'@Description("Sugar for .Item(LastIndex). Returns 'Empty' if the WCollection has no items")
Public Function Last() As Variant
Attribute Last.VB_Description = "Sugar for .Item(LastIndex). Returns 'Empty' if the WCollection has no items"
    Last = Empty
    If p.Coll.Count = 0 Then Exit Function
    Assign Last, p.Coll.Item(LastIndex)
End Function


'@Description("Returns a New WCollection containing unique items")
Public Function Dedup() As WCollection
Attribute Dedup.VB_Description = "Returns a New WCollection containing unique items"

    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    Dim myItem As Variant
    For Each myItem In p.Coll
    
        '@Ignore FunctionReturnValueDiscarded
        If myList.LacksItem(myItem) Then myList.Add myItem
        
    Next
    
    Set Dedup = myList
    
End Function
'@Description("Returns one 1 or more items from FirstIndex as an array.  Adds Queue terminology")
Public Function Dequeue(Optional ByVal ipCount As Long = 1) As Variant
Attribute Dequeue.VB_Description = "Returns one 1 or more items from FirstIndex as an array.  Adds Queue terminology"
   
    Dequeue = Empty 'Empty is better then Array()
    If ipCount < 1 Then Exit Function
    If p.Coll.Count = 0 Then Exit Function
    
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    Dim myCount As Long
    For myCount = 1 To ipCount
        
        '@Ignore FunctionReturnValueDiscarded, UnassignedVariableUsage
        myList.Add p.Coll.Item(1)
        p.Coll.Remove 1
        
        If Me.HasNoItems Then Exit For
        
    Next
    
    '@Ignore UnassignedVariableUsage
    Dequeue = myList.ToArray
    
End Function


'@Description("True if the Wcollection and ipEnumerable contain the same items, irrespective of order")
Public Function IsEquivalent(ByVal ipEnumerable As Variant) As Boolean
Attribute IsEquivalent.VB_Description = "True if the Wcollection and ipEnumerable contain the same items, irrespective of order"

    ' Taking the lazy way out for now
    IsEquivalent = IsSameOrder(WCollection.Deb(ipEnumerable).Sort)

End Function


'@Description("True if the items in the WCollection and ipEnumerable match when enumerated by index")
Public Function IsSameOrder(ByVal ipEnumerable As Variant) As Boolean
Attribute IsSameOrder.VB_Description = "True if the items in the WCollection and ipEnumerable match when enumerated by index"

    IsSameOrder = False
    GuardForIsNotEnumerable ipEnumerable
    
    Dim myEnumerable As Variant
    ' The Net Stack and Queue object do not have an Index method
    ' so we use the array equivalents
    If InStr(TypeName(ipEnumerable), "Stack,Queue") > 0 Then
        myEnumerable = ipEnumerable.ToArray
        
    Else
    
        Assign myEnumerable, ipEnumerable
        
    End If
    
    Dim myEnumerableCount As Long
    Select Case True
    
        Case VBA.IsArray(myEnumerable)
            myEnumerableCount = ArrayCount(ipEnumerable)
            
        Case HasCountMethod(myEnumerable)
            myEnumerableCount = myEnumerable.Count
            
    End Select
        
    If Me.Count <> myEnumerableCount Then Exit Function
        
   
    Dim myindex As Long
    myindex = 1
    Dim myItem As Variant
    For Each myItem In myEnumerable
        If myItem <> p.Coll.Item(myindex) Then Exit Function
        myindex = myindex + 1
    Next
    
    IsSameOrder = True
    
End Function
'@Description("Sugar for Add")
Public Function Enqueue(ParamArray ipItems() As Variant) As WCollection
Attribute Enqueue.VB_Description = "Sugar for Add"

    If ArrayHasNoItems(CVar(ipItems)) Then
        Set Enqueue = Me
        Exit Function
    End If
    
    Set Enqueue = AddRange(ipItems)
    
End Function

'@Description("Sugar for AddRange. Add Queue terminology")
Public Function EnqueueRange(ByVal ipEnumerable As Variant) As WCollection
Attribute EnqueueRange.VB_Description = "Sugar for AddRange. Add Queue terminology"
    Set EnqueueRange = AddRange(ipEnumerable)
End Function

'@Description("Sugar for AddString.  Adds Queue terminology")
Public Function EnqueueString _
( _
    ByVal ipString As String, _
    Optional ByVal ipSeparator As String _
) As WCollection
Attribute EnqueueString.VB_Description = "Sugar for AddString.  Adds Queue terminology"
    Set EnqueueString = AddString(ipString, ipSeparator)
End Function




'@Description("Sugar for Add.  Adds terminology for Stack")
Public Function Push(ParamArray ipItems() As Variant) As WCollection
Attribute Push.VB_Description = "Sugar for Add.  Adds terminology for Stack"

    Set Push = Me
    If ArrayHasNoItems(CVar(ipItems)) Then Exit Function
    
    Set Push = AddRange(CVar(ipItems))
    
End Function

'@Description("Sugar for AddRange.  Adds terminology for Stack")
Public Function PushRange(ByVal ipEnumerable As Variant) As WCollection
Attribute PushRange.VB_Description = "Sugar for AddRange.  Adds terminology for Stack"
    Set PushRange = AddRange(ipEnumerable)
End Function


'@Description("Sugar for AddString.  Adds Stack Terminology")
Public Function PushString _
( _
    ByVal ipString As String, _
    Optional ByVal ipSeparator As String _
) As WCollection
Attribute PushString.VB_Description = "Sugar for AddString.  Adds Stack Terminology"
    Set PushString = AddString(ipString, ipSeparator)
End Function



'@Description("Returns an array containing one or more items from LastIndex.  Adds Stack terminology")
Public Function Pop(Optional ByVal ipCount As Long = 1) As Variant
Attribute Pop.VB_Description = "Returns an array containing one or more items from LastIndex.  Adds Stack terminology"

    Pop = Empty
    If ipCount < 1 Then Exit Function
    If p.Coll.Count = 0 Then Exit Function
    
    Dim myList As WCollection
    Set myList = WCollection.Deb
 
    
    Dim myCount As Long
    For myCount = ipCount To 1 Step -1
        
        '@Ignore FunctionReturnValueDiscarded, UnassignedVariableUsage
        myList.Add p.Coll.Item(p.Coll.Count)
        p.Coll.Remove p.Coll.Count
        
        If Me.HasNoItems Then Exit For
        
    Next
    
    '@Ignore UnassignedVariableUsage
    Pop = myList.ToArray
    
End Function

'@Description("Sugar for Item Get. Adds Stack/Queue terminology")
Public Function Peek(Optional ByVal ipIndex As Long = 1) As Variant
Attribute Peek.VB_Description = "Sugar for Item Get. Adds Stack/Queue terminology"

    GuardForInvalidIndex ipIndex
    Dim myindex As Long
    myindex = ResolveStartIndex(ipIndex)
    Assign Peek, p.Coll.Item(myindex)
    
End Function

' Omit plngLeft & plngRight; they are used internally during recursion
Public Function Sort() As WCollection
    
    If Me.HasNoItems Then Exit Function
    
    'QuickSort3 from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2909260#post2909260
    Dim myArray As Variant
    myArray = Me.ToArray
    SnakeSort1 myArray
    Set Sort = WCollection.Deb(myArray)
    
End Function

' Set stuff
'@Description("Returns a new WCollection of all unique values in WCollection and ipEnumerable")
Public Function SetOfUnique(ByRef ipEnumerable As Variant) As WCollection
Attribute SetOfUnique.VB_Description = "Returns a new WCollection of all unique values in WCollection and ipEnumerable"


    If IsNotEnumerableObject(ipEnumerable) And Not VBA.IsArray(ipEnumerable) Then
    
        Set SetOfUnique = Me.Dedup
        
    Else
    
        Set SetOfUnique = WCollection.Deb.AddRange(p.Coll).AddRange(ipEnumerable).Dedup
        
    End If
    
End Function

'@Description("Returns a new WCollection of all values in both WCollection and ipEnumerable")
Public Function SetOfCommon(ByVal ipEnumerable As Variant) As WCollection
Attribute SetOfCommon.VB_Description = "Returns a new WCollection of all values in both WCollection and ipEnumerable"

    Dim myCommon As WCollection
    Set myCommon = WCollection.Deb
    
    Dim myItem As Variant
    For Each myItem In ipEnumerable
    
        '@Ignore FunctionReturnValueDiscarded
        If Me.HoldsItem(myItem) Then myCommon.Add myItem
    
    Next
    
    Set SetOfCommon = myCommon
    
End Function

'@Description("Returns a new WCollection of all values in WCollection only or ipEnumerable only")
Public Function SetOfNotCommon(ByRef ipEnumerable As Variant) As WCollection
Attribute SetOfNotCommon.VB_Description = "Returns a new WCollection of all values in WCollection only or ipEnumerable only"

    Dim myNotCommon As WCollection
    Set myNotCommon = WCollection.Deb(ipEnumerable).Dedup
    
    Dim myArray As Variant
    myArray = Me.Dedup.ToArray
    Dim myItem As Variant
    For Each myItem In myArray
    
        If myNotCommon.HoldsItem(myItem) Then
        
            '@Ignore FunctionReturnValueDiscarded
            myNotCommon.RemoveAnyOf myItem
            
        Else
        
            '@Ignore FunctionReturnValueDiscarded
            myNotCommon.Add myItem
            
        End If
        
    Next
    
    Set SetOfNotCommon = myNotCommon
    
End Function


'@Description("Returns a new WCollection of all values in WCollection only")
Public Function SetOfHostOnly(ByVal ipEnumerable As Variant) As WCollection
Attribute SetOfHostOnly.VB_Description = "Returns a new WCollection of all values in WCollection only"

    Dim myHostOnly As WCollection
    Set myHostOnly = Me.Clone
    
    Dim myItem As Variant
    For Each myItem In ipEnumerable
    
        If Me.HoldsItem(myItem) Then
        
            '@Ignore FunctionReturnValueDiscarded
            myHostOnly.RemoveAnyOf myItem
            
        End If
        
    Next
    
    Set SetOfHostOnly = myHostOnly
    
End Function

'@Description("Returns a new WCollection of all values in WCollection only")
Public Function SetOfInputOnly(ByVal ipEnumerable As Variant) As WCollection
Attribute SetOfInputOnly.VB_Description = "Returns a new WCollection of all values in WCollection only"

    Dim myInputOnly As WCollection
    Set myInputOnly = WCollection.Deb
    
    Dim myItem As Variant
    For Each myItem In ipEnumerable
    
        If Me.LacksItem(myItem) Then
        
            '@Ignore FunctionReturnValueDiscarded
            myInputOnly.Add myItem
            
        End If
        
    Next
    
    Set SetOfInputOnly = myInputOnly
    
End Function

' Some 'functional' stuff <distance sounds of thunder>
Public Function MapIt(ByVal ipMapIt As IMapIt) As WCollection
    
    Dim myW As WCollection
    Set myW = WCollection.Deb
    
    Dim myItem As Variant
    For Each myItem In p.Coll
        '@Ignore FunctionReturnValueDiscarded
        myW.Add ipMapIt.ExecMapIt(myItem)
    Next
    
    Set MapIt = myW
    
End Function


Public Function FilterIt(ByVal ipCompareIt As ICmpIt) As WCollection
    
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    Dim myItem As Variant
    For Each myItem In p.Coll
        '@Ignore FunctionReturnValueDiscarded
        If ipCompareIt.ExecCmpIt(myItem) Then myList.Add myItem
    Next
    
    Set FilterIt = myList
    
End Function


Public Function ReduceIt(ByVal ipReduceIt As IReduceIt) As Variant
    
    Dim myResult As Variant
    myResult = ipReduceIt.ExecReduceIt(p.Coll)
    
    Assign ReduceIt, myResult
    
End Function

'@Description("Takes a positive or negative index and returns a positive index")
Private Function ResolveStartIndex(ByVal ipStartIndex As Long) As Long
Attribute ResolveStartIndex.VB_Description = "Takes a positive or negative index and returns a positive index"
    ' This method is called after Guarding against invalid Indexes
    
    If ipStartIndex = 0 Then
    
        ResolveStartIndex = FirstIndex
    
    
    ElseIf ipStartIndex > 0 Then
    
        ResolveStartIndex = ipStartIndex
        
        
    Else
        ' remember myIndex is now negative
        ResolveStartIndex = LastIndex + ipStartIndex + 1
        
        
    End If
    
End Function

'@Description("Takes a positive or negative index and returns a positive index in the range of 1 to Count")
Private Function ResolveEndIndex(ByVal ipEndIndex As Long) As Long
Attribute ResolveEndIndex.VB_Description = "Takes a positive or negative index and returns a positive index in the range of 1 to Count"
    ' This method should be called after Guarding against invalid Indexes
    ' as it does not check for indexes outside of 1 to Count
    If ipEndIndex = 0 Then
    
        ResolveEndIndex = LastIndex
        
        
    ElseIf ipEndIndex > 0 Then
    
        ResolveEndIndex = ipEndIndex
        
        
    Else
        ' remember myIndex is now negative
        ResolveEndIndex = LastIndex + ipEndIndex + 1 '-1 which cancels out
        
        
    End If
    
End Function

Private Sub SnakeSort1(ByRef pvarArray As Variant)
    Dim i As Long
    '@Ignore HungarianNotation
    Dim iMin As Long
    '@Ignore HungarianNotation
    Dim iMax As Long
    '@Ignore HungarianNotation
    Dim lngIndex() As Long
    '@Ignore HungarianNotation
    Dim lngLevel As Long
    '@Ignore HungarianNotation
    Dim lngOldLevel As Long
    '@Ignore HungarianNotation
    Dim lngNewLevel As Long
    '@Ignore HungarianNotation
    Dim varMirror As Variant
    '@Ignore HungarianNotation
    Dim lngDirection As Long
    '@Ignore HungarianNotation
    Dim blnMirror As Boolean
    '@Ignore HungarianNotation
    Dim varSwap As Variant
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    ReDim lngIndex((iMax - iMin + 3) \ 2)
    lngIndex(0) = iMin
    i = iMin
    ' Initial loop: locate cutoffs for each ordered section
    Do Until i >= iMax
        Select Case lngDirection
            Case 1
                Do Until i = iMax
                    If pvarArray(i) > pvarArray(i + 1) Then Exit Do
                    i = i + 1
                Loop
            Case -1
                Do Until i = iMax
                    If pvarArray(i) < pvarArray(i + 1) Then Exit Do
                    i = i + 1
                Loop
            Case Else
                Do Until i = iMax
                    If pvarArray(i) <> pvarArray(i + 1) Then Exit Do
                    i = i + 1
                Loop
                If i = iMax Then lngDirection = 1
        End Select
        If lngDirection = 0 Then
            If pvarArray(i) > pvarArray(i + 1) Then
                lngDirection = -1
            Else
                lngDirection = 1
            End If
        Else
            lngLevel = lngLevel + 1
            lngIndex(lngLevel) = i * lngDirection
            lngDirection = 0
        End If
        i = i + 1
    Loop
    If Abs(lngIndex(lngLevel)) < iMax Then
        If lngDirection = 0 Then lngDirection = 1
        lngLevel = lngLevel + 1
        lngIndex(lngLevel) = i * lngDirection
    End If
    ' If the list is already sorted, exit
    If lngLevel <= 1 Then
        ' If sorted descending, reverse before exiting
        If lngIndex(lngLevel) < 0 Then
            For i = 0 To (iMax - iMin) \ 2
                varSwap = pvarArray(iMin + i)
                pvarArray(iMin + i) = pvarArray(iMax - i)
                pvarArray(iMax - i) = varSwap
            Next
        End If
        Exit Sub
    End If
    ' Main loop - merge section pairs together until only one section left
    ReDim varMirror(iMin To iMax)
    Do Until lngLevel = 1
        lngOldLevel = lngLevel
        For lngLevel = 1 To lngLevel - 1 Step 2
            If blnMirror Then
                '@Ignore UnassignedVariableUsage
                SnakeSortMerge varMirror, lngIndex(lngLevel - 1), lngIndex(lngLevel), lngIndex(lngLevel + 1), pvarArray
            Else
                '@Ignore UnassignedVariableUsage
                SnakeSortMerge pvarArray, lngIndex(lngLevel - 1), lngIndex(lngLevel), lngIndex(lngLevel + 1), varMirror
            End If
            lngNewLevel = lngNewLevel + 1
            lngIndex(lngNewLevel) = Abs(lngIndex(lngLevel + 1))
        Next
        If lngOldLevel Mod 2 = 1 Then
            If blnMirror Then
                For i = lngIndex(lngNewLevel) + 1 To iMax
                    '@Ignore UnassignedVariableUsage
                    pvarArray(i) = varMirror(i)
                Next
            Else
                For i = lngIndex(lngNewLevel) + 1 To iMax
                    varMirror(i) = pvarArray(i)
                Next
            End If
            lngNewLevel = lngNewLevel + 1
            lngIndex(lngNewLevel) = lngIndex(lngOldLevel)
        End If
        lngLevel = lngNewLevel
        lngNewLevel = 0
        blnMirror = Not blnMirror
    Loop
    ' Copy back to main array if necessary
    If blnMirror Then
        For i = iMin To iMax
            '@Ignore UnassignedVariableUsage
            pvarArray(i) = varMirror(i)
        Next
    End If
End Sub
 
'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
Private Sub SnakeSortMerge(ByRef pvarSource As Variant, ByRef plngLeft As Long, ByRef plngMid As Long, ByRef plngRight As Long, ByRef pvarDest As Variant)
    Dim L As Long
    Dim LMin As Long
    Dim LMax As Long
    Dim LStep As Long
    Dim R As Long
    Dim RMin As Long
    Dim RMax As Long
    Dim RStep As Long
    Dim O As Long
    
    If plngLeft <> 0 Then O = Abs(plngLeft) + 1
    If plngMid > 0 Then
        LMin = O
        LMax = Abs(plngMid)
        LStep = 1
    Else
        LMin = Abs(plngMid)
        LMax = O
        LStep = -1
    End If
    If plngRight > 0 Then
        RMin = Abs(plngMid) + 1
        RMax = Abs(plngRight)
        RStep = 1
    Else
        RMin = Abs(plngRight)
        RMax = Abs(plngMid) + 1
        RStep = -1
    End If
    L = LMin
    R = RMin
    Do
        If pvarSource(L) <= pvarSource(R) Then
            pvarDest(O) = pvarSource(L)
            If L = LMax Then
                For R = R To RMax Step RStep
                    O = O + 1
                    pvarDest(O) = pvarSource(R)
                Next
                Exit Do
            End If
            L = L + LStep
        Else
            pvarDest(O) = pvarSource(R)
            If R = RMax Then
                For L = L To LMax Step LStep
                    O = O + 1
                    pvarDest(O) = pvarSource(L)
                Next
                Exit Do
            End If
            R = R + RStep
        End If
        O = O + 1
    Loop
End Sub


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Guard Methods
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Raises an error if the absolute value of the index is not in the range 1 to Count")
Private Sub GuardForInvalidIndex(ByVal ipIndex As Long)
Attribute GuardForInvalidIndex.VB_Description = "Raises an error if the absolute value of the index is not in the range 1 to Count"

    If VBA.Math.Abs(ipIndex) >= 0 And VBA.Math.Abs(ipIndex) <= p.Coll.Count Then Exit Sub
    
    ReportError "Invalid Index:" & CStr(ipIndex)
         
End Sub

'@Description("Raises an error if the first index is greater hten the second index.")
Private Sub GuardForInvalidIndexOrder(ByVal ipFirstIndex As Long, ByVal ipSecondIndex As Long)
Attribute GuardForInvalidIndexOrder.VB_Description = "Raises an error if the first index is greater hten the second index."

    If ipFirstIndex <= ipSecondIndex Then Exit Sub
    
    ReportError "Invalid Index order: First index '" & CStr(ipFirstIndex) & "' is more than Second Index '" & CStr(ipSecondIndex) & "'"
    
End Sub

'@(Description(Raises an error is the argument is not an array and is not an object with a count method")
Private Sub GuardForIsNotEnumerable(ByVal ipEnumerable As Variant)
    
    If IsEnumerable(ipEnumerable) Then Exit Sub
    ReportError "Argument is not Enumerable"
    
End Sub

'@Description("Raises an error if the argument is not a string, does not allow objects with default string member")
Private Sub GuardForIsNotString(ByRef ipString As Variant)
Attribute GuardForIsNotString.VB_Description = "Raises an error if the argument is not a string, does not allow objects with default string member"

    If TypeName(ipString) = "String" Then Exit Sub
    
    ReportError "Argument is not a String"
    
End Sub

'@Description("True if the array is not allocated")
Private Sub GuardArrayHasNoItems(ByVal ipArray As Variant)
Attribute GuardArrayHasNoItems.VB_Description = "True if the array is not allocated"

    If ArrayHasItems(ipArray) Then Exit Sub
    
    ReportError "Array has no items"
    
End Sub


' This can be improved using vbWatchdog, or twinBasic so that
' the name of the failing method is also reported.
Private Sub ReportError(ByRef ipString As String)
    Err.Raise WCOLLECTION_ERROR, MY_NAME, ipString
    End
End Sub

    
