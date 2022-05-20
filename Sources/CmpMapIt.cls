VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmpMapIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'@PredeclaredId
'@Exposed

Implements ICmpIt


Private Const MY_NAME           As String = "cmpIt"

Private Type State

    Comparison                  As Comparison
    Reference                   As Variant
    Index                       As Variant
    MapIt                       As IMapIt
    
End Type

Private s                       As State

Private Function ICmpIt_ExecCmpIt(ByRef ipHostItem As Variant) As Boolean
    Assign ICmpIt_ExecCmpIt, ExecCmpIt(ipHostItem)
End Function

Private Function ICmpIt_Item() As Variant
    Assign ICmpIt_Item, Item
End Function


Private Function ICmpIt_Index() As Long
  ICmpIt_Index = Index
End Function

Private Function ICmpIt_Name() As String
  ICmpIt_Name = Name
End Function

'@Description("Tests the result of the Map function but returns the item if true, not the result of the map. Index applies to the result of the map function")
Public Function ExecCmpIt(ByVal ipHostItem As Variant) As Boolean
Attribute ExecCmpIt.VB_Description = "Tests the result of the Map function but returns the item if true, not the result of the map. Index applies to the result of the map function"
    
    ExecCmpIt = GetComparison(s.MapIt.ExecMapIt(ipHostItem))
    
    
End Function

'@Description("Tests the result of the Map function but returns the Host item if true, not the result of the map. Index applies to the result of the map function")
'@DefaultMember
Public Function Deb _
( _
    ByVal ipMapIt As IMapIt, _
    ByVal ipComparison As Comparison, _
    ByVal ipReference As Variant, _
    Optional ByVal ipIndex As Variant = Empty _
) As CmpMapIt
Attribute Deb.VB_Description = "Tests the result of the Map function but returns the Host item if true, not the result of the map. Index applies to the result of the map function"
Attribute Deb.VB_UserMemId = 0

    With New CmpMapIt
    
        Set Deb = .ConstructInstance(ipMapIt, ipComparison, ipReference, ipIndex)
        
    End With
    
End Function


Friend Function ConstructInstance _
( _
    ByVal ipMapIt As IMapIt, _
    ByVal ipComparison As Comparison, _
    ByVal ipReference As Variant, _
    ByVal ipIndex As Variant _
) As CmpMapIt
    
    s.Comparison = ipComparison
    s.Reference = ipReference
    s.Index = ipIndex
    Set s.MapIt = ipMapIt
    Set ConstructInstance = Me
    
End Function
    

Public Property Get Name() As String
  Name = MY_NAME
End Property

Public Function Index() As Long
    Index = s.Index
End Function

Public Function Item() As Variant
  Item = s.Reference
End Function

Public Function GetComparison(ByRef ipMappedItem As Variant) As Boolean

    Dim myItem As Variant

    If VBA.IsEmpty(s.Index) Then

        Assign myItem, ipMappedItem

    
    ElseIf TypeName(ipMappedItem) = "String" Then

        myItem = Mid$(ipMappedItem, s.Index, 1)


    Else

        On Error Resume Next

        Assign myItem, ipMappedItem(s.Index)
        If Err.Number <> 0 Then

           GetComparison = False
           On Error GoTo 0
           Exit Function

        End If

        On Error GoTo 0


    End If


    Select Case s.Comparison

        Case Comparison.IsEqualTo:                              GetComparison = myItem = s.Reference
        Case Comparison.IsNotEqualTo:                           GetComparison = myItem <> s.Reference
        Case Comparison.IsLessThan:                             GetComparison = myItem < s.Reference
        Case Comparison.IsLessThanOrEqualTo, IsNotMoreThan:     GetComparison = myItem <= s.Reference
        Case Comparison.IsMoreThan:                             GetComparison = myItem > s.Reference
        Case Comparison.IsMoreThanOrEqualTo, IsNotLessThan:     GetComparison = myItem >= s.Reference


    End Select

End Function
