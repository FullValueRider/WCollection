VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpToChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
    
Implements IMapIt

Private Function IMapIt_ExecMapIt(ByVal ipHostItem As Variant) As Variant
    Assign IMapIt_ExecMapIt, ExecMapIt(ipHostItem)
End Function


Public Function ExecMapIt(ByVal ipHostItem As Variant) As Variant
        
    If VBA.VarType(ipHostItem) <> vbString Then
            
        Assign ExecMapIt, ipHostItem
    
    
    Else
        
        Assign ExecMapIt, WCollection.Deb(ipHostItem)
        
    
    End If
        
End Function


Public Function Deb() As mpToChars
    
    With New mpToChars
            
        Set Deb = .ConstructInstance()
            
    End With
    
End Function


Friend Function ConstructInstance() As mpToChars
    Set ConstructInstance = Me
End Function
