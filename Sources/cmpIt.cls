VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpIt"
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
Public Function ExecCmpIt(ByRef ipHostItem As Variant) As Boolean
Attribute ExecCmpIt.VB_Description = "Tests the result of the Map function but returns the item if true, not the result of the map. Index applies to the result of the map function"
    
    
    Dim myItem As Variant
    
    If VBA.IsEmpty(s.Index) Then
    
        Assign myItem, ipHostItem
        
    'ToDo: This is a poor test, needs a better method which checks for interface type

        
    ElseIf TypeName(ipHostItem) = "String" Then
    
        myItem = Mid$(ipHostItem, s.Index, 1)
        
    
    Else
    
        On Error Resume Next
            
        Assign myItem, ipHostItem(s.Index)
        If Err.Number <> 0 Then
        
           ExecCmpIt = False
           On Error GoTo 0
           Exit Function
           
        End If
        
        On Error GoTo 0
        
        
    End If
    
    
    Select Case s.Comparison
    
        Case Comparison.IsEqualTo:                              ExecCmpIt = myItem = s.Reference
        Case Comparison.IsNotEqualTo:                           ExecCmpIt = myItem <> s.Reference
        Case Comparison.IsLessThan:                             ExecCmpIt = myItem < s.Reference
        Case Comparison.IsLessThanOrEqualTo, IsNotMoreThan:     ExecCmpIt = myItem <= s.Reference
        Case Comparison.IsMoreThan:                             ExecCmpIt = myItem > s.Reference
        Case Comparison.IsMoreThanOrEqualTo, IsNotLessThan:     ExecCmpIt = myItem >= s.Reference
        
        
    End Select
    
End Function

'@Description("Tests the result of the Map function but returns the Host item if true, not the result of the map. Index applies to the result of the map function")
'@DefaultMember
Public Function Deb(ByVal ipComparison As Comparison, ByVal ipReference As Variant, Optional ByVal ipIndex As Variant = Empty) As cmpIt
Attribute Deb.VB_Description = "Tests the result of the Map function but returns the Host item if true, not the result of the map. Index applies to the result of the map function"
Attribute Deb.VB_UserMemId = 0

    With New cmpIt
    
        Set Deb = .ConstructInstance(ipComparison, ipReference, ipIndex)
        
    End With
    
End Function


Friend Function ConstructInstance(ByVal ipComparison As Comparison, ByVal ipReference As Variant, ByVal ipIndex As Variant) As cmpIt
    
    s.Comparison = ipComparison
    s.Reference = ipReference
    s.Index = ipIndex
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

'Public Function GetComparison(ByRef ipHostItem As Variant) As Boolean
'
'    Dim myItem As Variant
'
'    If VBA.IsEmpty(s.Index) Then
'
'        Assign myItem, ipHostItem
'
'    'ToDo: This is a poor test, needs a better method which checks for interface type
'    ElseIf VBA.Left$(TypeName(ipItem), 2) = "mp" Then
'
'        Dim myIFace As IMapIt
'        Set myIFace = ipItem
'        Assign myItem, myIFace.ExecMapIt(ipItem)
'
'    ElseIf TypeName(ipHostItem) = "String" Then
'
'        myItem = Mid$(ipHostItem, s.Index, 1)
'
'
'    Else
'
'        On Error Resume Next
'
'        Assign myItem, ipHostItem(s.Index)
'        If Err.Number <> 0 Then
'
'           GetComparison = False
'           On Error GoTo 0
'           Exit Function
'
'        End If
'
'        On Error GoTo 0
'
'
'    End If
'
'
'    Select Case s.Comparison
'
'        Case Comparison.IsEqualTo:                              GetComparison = myItem = s.Reference
'        Case Comparison.IsNotEqualTo:                           GetComparison = myItem <> s.Reference
'        Case Comparison.IsLessThan:                             GetComparison = myItem < s.Reference
'        Case Comparison.IsLessThanOrEqualTo, IsNotMoreThan:     GetComparison = myItem <= s.Reference
'        Case Comparison.IsMoreThan:                             GetComparison = myItem > s.Reference
'        Case Comparison.IsMoreThanOrEqualTo, IsNotLessThan:     GetComparison = myItem >= s.Reference
'
'
'    End Select
'
'End Function
