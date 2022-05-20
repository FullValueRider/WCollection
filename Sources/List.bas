VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Rubberduck annotations
'@PredeclaredId
'@Exposed
Option Explicit


'@ModuleDescription("A wrapper for the collection object to add flexibility")

Private Type State
    
    Coll                                        As Collection
    
End Type

Private s                                       As State

Private Sub Class_Initialize()
    Set s.Coll = New Collection
End Sub

Public Function Deb() As WCollection
    With New WCollection
        Set Deb = .ReadyToUseInstance
    End With
End Function


Friend Function ReadyToUseInstance() As WCollection
    Set ReadyToUseInstance = Me
End Function

Public Function NewEnum() As IEnumVARIANT
    Set NewEnum = s.Coll.[_NewEnum]
End Function

Public Function Add(ParamArray ipItems() As Variant) As WCollection
    Dim myItem As Variant
    For Each myItem In ipItems
        s.Coll.Add myItem
    Next
        Set Add = Me
End Function

Public Function AddRange(ByVal ipIterable As Variant) As WCollection
    Dim myitem As Variant
    For Each myitem In ipIterable
        s.Coll.Add myitem
    Next
    Set AddRange = Me
End Function


Public Function AddString(ByVal ipString As String) As WCollection
    Dim myIndex As Long
    For myIndex = 1 To Len(ipString)
        s.Coll.Add VBA.Mid$(ipString, myIndex, 1)
    Next
End Function


Public Function Clone() As WCollection
    Set Clone = WCollection.Deb.AddRange(s.Coll)
End Function
'@DefaultMember
Public Property Get Item(ByVal ipIndex As Long) As Variant
    If VBA.IsObject(s.Coll.Item(ipIndex)) Then
        Set Item = s.Coll.Item(ipIndex)
    Else
        Item = s.Coll.Item(ipIndex)
    End If
End Property

Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
    s.Coll.Add ipItem, after:=ipIndex
    s.Coll.Remove ipIndex
End Property

Public Property Set Item(ByVal ipindex As Long, ByVal ipitem As Variant)
    s.Coll.Add ipitem, after:=ipindex
    s.Coll.Remove ipindex
End Property


Public Function HoldsItem(ByVal ipItem As Variant) As Boolean
    HoldsItem = True
    Dim myItem As Variant
    For Each myItem In s.Coll
        If myItem = ipItem Then Exit Function
    Next
    HoldsItem = False
End Function

Public Function Join(Optional ByVal ipSeparator As String) As String
    
    If TypeName(s.Coll.Item(1)) <> "String" Then
        Join = "Items are not string type"
        Exit Function
    End If
    
    Dim myStr As String
    Dim myItem As Variant
    For Each myItem In s.Coll
        If Len(myStr) = 0 Then
            myStr = myItem
        Else
            myStr = myStr & ipSeparator
        End If
        
    Next

End Function

Public Function Reverse() As WCollection
    Dim myW As WCollection
    Set myW = WCollection.Deb
    Dim myIndex As Long
    For myIndex = LastIndex To FirstIndex Step -1
        myW.Add s.Coll.Item(myIndex)
    Next
    Set Reverse = myW
End Function

Public Function HasItems() As Boolean
    HasItems = s.Coll.Count > 0
End Function

Public Function HasNoItems() As Boolean
    HasNoItems = Not HasItems
End Function

Public Function Indexof(ByVal ipItem As Variant, Optional ipIndex As Long = -1) As Long
    Dim myIndex As Long
    For myIndex = IIf(ipIndex = -1, 1, ipIndex) To s.Coll.Count
        If ipItem = s.Coll.Item(myIndex) Then
            Indexof = myIndex
            Exit Function
        End If
    Next
End Function

Public Function LastIndexof(ByVal ipItem As Variant, Optional ipIndex As Long = -1) As Long
    Dim myIndex As Long
    For myIndex = LastIndex To IIf(ipIndex = -1, 1, ipIndex) Step -1
        If ipItem = s.Coll.Item(myIndex) Then
            LastIndexof = myIndex
            Exit Function
        End If
    Next
    LastIndexof = -1
End Function

Public Function LacksItem(ByVal ipItem As Variant) As Boolean
    LacksItem = Not HoldsItem(ipItem)
End Function


Public Function Insert(ByVal ipIndex As Long, ByVal ipItem As Variant) As WCollection
    s.Coll.Add ipItem, before:=ipIndex
    Set Insert = Me
End Function


Public Function Remove(ByVal ipIndex As Long) As WCollection
    s.Coll.Remove ipIndex
    Set Remove = Me
End Function

Public Function FirstIndex() As Long
    FirstIndex = 1
End Function

Public Function LastIndex() As Long
    LastIndex = s.Coll.Count
End Function

Public Function RemoveAll() As WCollection
    Dim myIndex As Long
    For myIndex = s.Coll.Count To 1 Step -1
        Remove myIndex
    Next
    Set RemoveAll = Me
End Function


Public Property Get Count() As Long
    Count = s.Coll.Count
End Property

Public Function ToArray() As Variant
    Dim myarray As Variant
    ReDim myarray(0 To s.Coll.Count - 1)
    Dim myItem As Variant
    Dim myIndex As Long
    myIndex = 0
    For Each myItem In s.Coll
        If VBA.IsObject(myItem) Then
            Set myarray(myIndex) = myItem
        Else
            myarray(myIndex) = myItem
        End If
        myIndex = myIndex + 1
    Next
    ToArray = myarray
End Function

Public Function RemoveFirstOf(ByVal ipItem As Variant) As WCollection
    Set RemoveFirstOf = Remove(Indexof(ipItem))
    Set RemoveFirstOf = Me
End Function

Public Function RemoveLastOf(ByVal ipItem As Variant) As WCollection
    Set RemoveLastOf = Remove(LastIndexof(ipItem))
    Set RemoveLastOf = Me
End Function

Public Function RemoveAnyOf(ByVal ipItem As Variant) As WCollection
    Dim myIndex As Long
    For myIndex = LastIndex To FirstIndex Step -1
        
        If s.Coll.Item(myIndex) = ipItem Then Remove myIndex
        
    Next
    Set RemoveAnyOf = Me
End Function

Public Function First() As Variant
    If VBA.IsObject(s.Coll.Item(FirstIndex)) Then
        Set First = s.Coll.Item(FirstIndex)
    Else
        First = s.Coll.Item(FirstIndex)
    End If
End Function

Public Function Last() As Variant
    If VBA.IsObject(s.Coll.Item(LastIndex)) Then
        Set Last = s.Coll.Item(LastIndex)
    Else
        Last = s.Coll.Item(LastIndex)
    End If
End Function

Public Function Enqueue(ByVal ipItem As Variant) As WCollection
    Add ipItem
    Set Enqueue = Me
End Function

Public Function Dequeue() As Variant
    If VBA.IsObject(s.Coll.Item(FirstIndex)) Then
        Set Dequeue = s.Coll.Item(FirstIndex)
    Else
        Dequeue = s.Coll.Item(FirstIndex)
    End If
    Remove 0
End Function

Public Function Push(ByVal ipitem As Variant) As WCollection
    Add ipitem
    Set Push = Me
End Function

Public Function Pop(ByVal ipitem As Variant) As Variant
    If VBA.IsObject(s.Coll.Item(FirstIndex)) Then
        Set Pop = s.Coll.Item(FirstIndex)
    Else
        Pop = s.Coll.Item(FirstIndex)
    End If
    Remove s.Coll.Count
End Function

Public Function Peek(ByVal ipIndex As Long) As Variant
    If VBA.IsObject(s.Coll.Item(FirstIndex)) Then
        Set Peek = s.Coll.Item(FirstIndex)
    Else
        Peek = s.Coll.Item(FirstIndex)
    End If
End Function