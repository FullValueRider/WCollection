VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICmpIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Folder("CompareIts")
'@Interface
'@PredeclaredId
'@Exposed

Public Enum Comparison

    IsEqualTo
    IsNotEqualTo
    IsMoreThan
    IsMoreThanOrEqualTo
    IsNotLessThan
    IsLessThan
    IsLessThanOrEqualTo
    IsNotMoreThan
    
End Enum
    
Public Function ExecCmpIt(ByRef ipHostItem As Variant) As Boolean
End Function

Public Function Item() As Variant
End Function

Public Function Index() As Long
End Function

Public Function Name() As String
End Function
