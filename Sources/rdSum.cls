VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'@PredeclaredId
'@Exposed


Implements IReduceIt

Private Function IReduceIt_ExecReduceIt(ByVal ipHostCollection As Collection) As Variant
    IReduceIt_ExecReduceIt = ExecReduceIt(ipHostCollection)
End Function

Public Function ExecReduceIt(ByVal ipHostCollection As Collection) As Variant

    Dim myResult As Variant
    
    Dim myItem As Variant
    For Each myItem In ipHostCollection
    
        If (TypeName(myItem) <> "String") And (Not VBA.IsObject(myItem)) Then
        
        
            myResult = myResult + myItem
        
        End If
        
    Next
    
    ExecReduceIt = myResult
    
End Function


Public Function Deb() As rdSum
    With New rdSum
        Set Deb = .ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As rdSum
    Set ConstructInstance = Me
End Function
