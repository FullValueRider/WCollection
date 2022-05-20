VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpLen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule DefaultMemberRequired
Option Explicit
'@PredeclaredId
'@Exposed

Implements IMapIt
    

'@Description("Add the value to each item in the WCollection. Optional Value is activated by 'mpMapIt()', just 'mpMapIt' will fail")
'@DefaultMember
Public Function Deb() As mpLen
Attribute Deb.VB_Description = "Add the value to each item in the WCollection. Optional Value is activated by 'mpMapIt()', just 'mpMapIt' will fail"
Attribute Deb.VB_UserMemId = 0
        
    With New mpLen
        
        Set Deb = .ConstructInstance()
            
    End With
        
End Function

Friend Function ConstructInstance() As mpLen
    Set ConstructInstance = Me
End Function

'@Description("Adds the set value to each item in WCollection. If there is an error then "Nan" is returned)
Private Function IMapIt_ExecMapIt(ByVal ipHostItem As Variant) As Variant
    Assign IMapIt_ExecMapIt, ExecMapIt(ipHostItem)
End Function

'@Description("Returns the length of the string or -1 if not a string")
Public Function ExecMapIt(ByVal ipHostItem As Variant) As Variant
Attribute ExecMapIt.VB_Description = "Returns the length of the string or -1 if not a string"
   
    If TypeName(ipHostItem) = "String" Then
    
        ExecMapIt = VBA.Len(ipHostItem)
        
    Else
    
        ExecMapIt = -1
        
    End If
    
            
End Function
