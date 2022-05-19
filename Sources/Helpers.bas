Attribute VB_Name = "Helpers"
Option Explicit

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Helper methods normally housed in other classes/modules
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

'@Description("Simplifies assignment when Item may be a value or object")
Public Sub Assign(ByRef opTo As Variant, ByRef ipFrom As Variant)
Attribute Assign.VB_Description = "Simplifies assignment when Item may be a value or object"
    
    If IsObject(ipFrom) Then
        
        Set opTo = ipFrom
        
        
    Else
        
        opTo = ipFrom
        
        
    End If
    
End Sub

' The next four methods are provided for examining
' ParamArrays but will work with any array.
' ParamArrays must be converted to a variant (CVar(<ParamArray>))
' before they can be forwarded to these next four methods.

'@Description("True if an Array has 1 or more items.")
Public Function ArrayHasAnyItems(ByRef ipArray As Variant) As Boolean
Attribute ArrayHasAnyItems.VB_Description = "True if an Array has 1 or more items."
    ' Taken from http://www.cpearson.com/excel/isarrayallocated.aspx
    On Error Resume Next
    
    ArrayHasAnyItems = _
        VBA.IsArray(ipArray) _
        And Not IsError(LBound(ipArray, 1)) _
        And LBound(ipArray, 1) <= UBound(ipArray, 1)
        
    On Error GoTo 0
    
End Function

'@Description("True if an array has not been initialised
Public Function ArrayHasNoItems(ByRef ipArray As Variant) As Boolean
    ArrayHasNoItems = Not ArrayHasAnyItems(ipArray)
End Function

'@Description("True if an Array contians only One Item")
Public Function ArrayHasOneItem(ByRef ipArray As Variant) As Boolean
Attribute ArrayHasOneItem.VB_Description = "True if an Array contians only One Item"
    ArrayHasOneItem = 1 = (UBound(ipArray) - LBound(ipArray) + 1)
End Function

'@Description("True if array has 2 or more items")
Public Function ArrayHasItems(ByRef ipArray As Variant) As Boolean
Attribute ArrayHasItems.VB_Description = "True if array has 2 or more items"
    ArrayHasItems = 1 < (UBound(ipArray) - LBound(ipArray) + 1)
End Function


'@Description("True if the object has a count method (presumed to be an enumerable object)")
Public Function HasCountMethod(ByVal ipObject As Variant) As Boolean
Attribute HasCountMethod.VB_Description = "True if the object has a count method (presumed to be an enumerable object)"

    On Error Resume Next
    '@Ignore VariableNotUsed
    Dim myCount As Long
    '@Ignore AssignmentNotUsed
    myCount = ipObject.Count
    HasCountMethod = Err.Number = 0
    On Error GoTo 0
    Err.Clear
    
End Function

'@Description("Because I dislike the Not <xxxx> construct in VBA.")
Public Function HasNoCountMethod(ByVal ipObject As Variant) As Boolean
Attribute HasNoCountMethod.VB_Description = "Because I dislike the Not <xxxx> construct in VBA."
    HasNoCountMethod = Not HasCountMethod(ipObject)
End Function


'@("True if item is string , ignores objects in case the default member returns a string")
Public Function IsString(ByVal ipString As Variant) As Boolean

    If VBA.IsObject(ipString) Then
    
        IsString = False
        
    Else
    
        IsString = VarType(ipString) = vbString
        
    End If
    
End Function

'@Description("True if the Item is an object and has a count method.  Excludes UDTs with count field.
Public Function IsEnumerableObject(ByRef ipEnumerableObject As Variant) As Boolean
    IsEnumerableObject = False
    If Not VBA.IsObject(ipEnumerableObject) Then Exit Function
    If HasNoCountMethod(ipEnumerableObject) Then Exit Function
    IsEnumerableObject = True
End Function

Public Function IsNotEnumerableObject(ByRef ipEnumerableObject As Variant) As Boolean
    IsNotEnumerableObject = Not IsEnumerableObject(ipEnumerableObject)
End Function


Public Function IsEnumerable(ByVal ipEnumerable As Variant) As Boolean
    IsEnumerable = True
    If VBA.IsArray(ipEnumerable) Then Exit Function
    If IsEnumerableObject(ipEnumerable) Then Exit Function
    IsEnumerable = False
End Function


Public Function IsNotEnumerable(ByVal ipEnumerable As Variant) As Boolean
    IsNotEnumerable = Not IsEnumerable(ipEnumerable)
End Function

'@Description("Returns the number of elements in an array")
Public Function ArrayCount(ByVal ipArray As Variant, Optional ByVal ipRank As Long = 0) As Long
Attribute ArrayCount.VB_Description = "Returns the number of elements in an array"
    
    ArrayCount = 0

    If ArrayHasNoItems(ipArray) Then Err.Raise 17 + vbObjectError, "WCollection Helpers", "The array has no items"
    
    Dim mySize As Long
    If ipRank = 0 Then ' Count all elements of the array
        mySize = 1
        Dim myRank As Long
        For myRank = 1 To ArrayRanks(ipArray)
            mySize = mySize * (UBound(ipArray, myRank) - LBound(ipArray, myRank) + 1)
        Next
    Else
        mySize = UBound(ipArray, ipRank) - LBound(ipArray, ipRank) + 1
    End If
            
    ArrayCount = mySize
    
        
End Function


'@Description("Returns the number of dimensions of an array. Return values >1:No of Ranks, 0: ")
Public Function ArrayRanks(ByVal ipArray As Variant) As Long
Attribute ArrayRanks.VB_Description = "Returns the number of dimensions of an array. Return values >1:No of Ranks, 0: "

    Dim myindex As Long
    For myindex = 1 To 60000
    
        '@Ignore UnhandledOnErrorResumeNext
        On Error Resume Next
        '@Ignore VariableNotUsed
        Dim myDummy As Long
        myDummy = UBound(ipArray, myindex)
        
        If Err.Number <> 0 Or myDummy = -1 Then
            
            Err.Clear
            Exit For
            
            
        End If
        
        On Error GoTo 0
        
        
    Next
    
    ArrayRanks = myindex - 1
    
End Function
