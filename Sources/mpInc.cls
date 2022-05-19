VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed

Implements IMapIt
    
Private Const NotANumber            As String = "nan"

Private Type State
        
    Increment                       As Variant
        
        
End Type

Private s                           As State
'@Description("Add the value to each item in the WCollection. Optional Value is activated by 'mpMapIt()', just 'mpMapIt' will fail")
'@DefaultMember
Public Function Deb(Optional ByVal ipIncrement As Variant = 1) As mpInc
Attribute Deb.VB_Description = "Add the value to each item in the WCollection. Optional Value is activated by 'mpMapIt()', just 'mpMapIt' will fail"
Attribute Deb.VB_UserMemId = 0
        
    With New mpInc
        
        Set Deb = .ConstructInstance(ipIncrement)
            
    End With
        
End Function

Friend Function ConstructInstance(ByVal ipIncrement As Variant) As mpInc
        
    s.Increment = ipIncrement
    Set ConstructInstance = Me
        
End Function

'@Description("Adds the set value to each item in WCollection. If there is an error then "Nan" is returned)
Private Function IMapIt_ExecMapIt(ByVal ipHostItem As Variant) As Variant
    Assign IMapIt_ExecMapIt, ExecMapIt(ipHostItem)
End Function

Public Function ExecMapIt(ByVal ipHostItem As Variant) As Variant
   
    On Error Resume Next
    ' Assign requires ByRef parameters so we can't pass in an expression
    Dim myValue As Variant
    myValue = ipHostItem + s.Increment
    If Err.Number = 0 Then
        On Error GoTo 0
        Assign ExecMapIt, myValue
    Else
        On Error GoTo 0
        Assign ExecMapIt, NotANumber
    End If
            
End Function
