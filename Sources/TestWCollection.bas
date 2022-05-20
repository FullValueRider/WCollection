Attribute VB_Name = "TestWCollection"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("Deb")
Private Sub Test01_IsObjectIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    Dim myList As WCollection
    
    Dim myResult As Boolean
    'Act:
    Set myList = WCollection.Deb
    myResult = IsObject(myList)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 01 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Deb")
Private Sub Test02_WCollectionNameIsWCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "WCollection"
    Dim myList As WCollection
    
    Dim myResult As String
    'Act:
    Set myList = WCollection.Deb
    myResult = TypeName(myList)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 02 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Deb")
Private Sub Test03_DebIsCountZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 0
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 03 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Deb")
Private Sub Test04_DebWithParamArrayIsCountFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb(1, 2, 3, 4, 5)
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 04 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Deb")
Private Sub Test05_DebWithOneArrayCountIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb(Array(1, 2, 3, 4, 5))
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 05 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Deb")
Private Sub Test06_DebWithOneCollectionCountIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    Dim myList As WCollection
    
    Dim myTest As Collection
    Set myTest = New Collection
    
    With myTest
    
        .Add 1
        .Add 2
        .Add 3
        .Add 4
        .Add 5
        
    End With
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb(myTest)
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 06 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Deb")
Private Sub Test07_DebWithCollectionAndLongCountIsTwo()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 2
    Dim myList As WCollection
    
    Dim myTest As Collection
    Set myTest = New Collection
    
    With myTest
    
        .Add 1
        .Add 2
        .Add 3
        .Add 4
        .Add 5
        
    End With
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb(myTest, 42)
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 07 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Deb")
Private Sub Test08_DebWithStringHelloCountIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    
    Dim myList As WCollection
    Set myList = WCollection.Deb("Hello")
    
    Dim myResult As Long
    
    'Act:
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 08 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Add")
Private Sub Test09_AddParamArrayIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 09 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Add")
Private Sub Test10_AddCollectionIsOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 1
    
    
    Dim myTest As Collection
    Set myTest = New Collection
    
    With myTest
    
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        
    End With
    
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(myTest)
    
    Dim myResult As Long
    
    'Act:
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 10 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AddRange")
Private Sub Test11_AddRangeCollectionIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    
    Dim myTest As Collection
    Set myTest = New Collection
    
    With myTest
    
        .Add 1
        .Add 2
        .Add 3
        .Add 4
        .Add 5
        
    End With
    
    Dim myList As WCollection
    Dim myResult As Long
    
    'Act:
     Set myList = WCollection.Deb.AddRange(myTest)
     myResult = myList.Count
    
    'Assert:
     Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 11 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AddString")
Private Sub Test12_AddStringHelloIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    
    Dim myList As WCollection
    Dim myResult As Long
    'Act:
    
    Set myList = WCollection.Deb.AddString("Hello")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 12 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AddString")
Private Sub Test13a_AddSeparatedStringStringWithNoSeparatorCountIs26()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 26
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.AddString("Hello World Its A Nice Day")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 12 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AddString")
Private Sub Test13b_AddSeparatedStringWithSeparatorCountIsSix()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 6
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.AddString("Hello World Its A Nice Day", " ")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 13 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Clone")
Private Sub Test14_CloneUsingDoubles()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    Set myExpected = WCollection.Deb.AddRange(Array(1#, 2#, 3#, 4#, 5#))
   
    Dim myResult As WCollection
    'Act:
    Dim myList As WCollection
    Set myList = WCollection.Deb.AddRange(Array(1#, 2#, 3#, 4#, 5#))
    Set myResult = myList.Clone
    
    'Assert:
    Assert.AreEqual myExpected.Item(1), myResult.Item(1)
    Assert.AreEqual myExpected.Item(2), myResult.Item(2)
    Assert.AreEqual myExpected.Item(3), myResult.Item(3)
    Assert.AreEqual myExpected.Item(4), myResult.Item(4)
    Assert.AreEqual myExpected.Item(5), myResult.Item(5)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 14 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Clone")
Private Sub Test15_CloneEmptyWCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    Set myExpected = WCollection.Deb
   
    Dim myResult As WCollection
    'Act:
    Dim myList As WCollection
    Set myList = WCollection.Deb
    Set myResult = myList.Clone
    
    'Assert:
    Assert.AreEqual myExpected.Count, 0&
    Assert.AreEqual myResult.Count, 0&
   
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 15 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test16a_GetItemThreeOfFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 30
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    myResult = myList.Item(3)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 16 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test16b_GetItemMinusTwoOfFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 40
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    myResult = myList.Item(-2)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 16 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test17a_LetItemThreeOfFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 42
   
    Dim myResult As Long
    'Act:
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    myList.Item(3) = 42
    myResult = myList.Item(3)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 17 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test17b_LetItemThreeOfFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 42
   
    Dim myResult As Long
    'Act:
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    myList.Item(-2) = 42
    myResult = myList.Item(-2)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 17 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test17a_SetItemThreeOfFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 42
   
    Dim myResult As Long
    'Act:
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    myResult = myList.SetItem(3, 42).Item(3)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 17 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test17b_SetItemThreeOfFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 42
   
    Dim myResult As Long
    'Act:
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    myResult = myList.SetItem(-2, 42).Item(-2)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 17 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

''@TestMethod("WCollection")
'Private Sub Test18_GetItemDefaultMemberThreeOfFive()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 30
'    Dim myList As WCollection
'
'
'    Dim myResult As Long
'    'Act:
'    Set myList = WCollection.Deb.AddRange(Array(10, 20, 30, 40, 50))
'    myResult = myList(3)
'
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test 18 raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("WCollection")
'Private Sub Test19_LetItemDefaultMemberThreeOfFive()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 42
'
'    Dim myResult As Long
'    Dim myList As WCollection
'    'Act:
'    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
'    myList.Item(3) = 42
'    myResult = myList(3)
'
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test 19 raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

'@TestMethod("HoldsItem")
Private Sub Test20_HoldsItemTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    myResult = myList.HoldsItem(30)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 20 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("HoldsItem")
Private Sub Test21_HoldsItemFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    myResult = myList.HoldsItem(42)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 21 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LacksItem")
Private Sub Test22_LacksItemFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    myResult = myList.LacksItem(30)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 22 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LacksItem")
Private Sub Test23_LacksItemTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    myResult = myList.LacksItem(42)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 23 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Join")
Private Sub Test24_JoinEmptyWCOllection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = vbNullString
   
    Dim myResult As String
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb
    
    myResult = myList.Join
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 25 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Join")
Private Sub Test25_JoinStringWithsDefaultSeparator()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello,World,Its,A,Nice,Day"
   
    Dim myResult As String
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "World", "Its", "A", "Nice", "Day")
    
    myResult = myList.Join
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 25 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Join")
Private Sub Test26_JoinStringsSpecifiedSeparator()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World Its A Nice Day"
   
    Dim myResult As String
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "World", "Its", "A", "Nice", "Day")
    
    myResult = myList.Join(" ")
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 26 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Reverse")
Private Sub Test27_Reverse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As WCollection
    Set myExpected = WCollection.Deb(50, 40, 30, 20, 10)
   
    Dim myResult As WCollection
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    Set myResult = myList.Reverse
    
    'Assert:
    Assert.AreEqual myExpected(1), myResult(1)
    Assert.AreEqual myExpected(2), myResult(2)
    Assert.AreEqual myExpected(3), myResult(3)
    Assert.AreEqual myExpected(4), myResult(4)
    Assert.AreEqual myExpected(5), myResult(5)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 27 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test28_HasItemsFiveItemsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    myResult = myList.HasItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 28 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Has")
Private Sub Test29_HasItemsOneItemIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10)
    
    myResult = myList.HasItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 29 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test30_HasItemsNoItemsIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb
    
    myResult = myList.HasItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 30 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test31_HasOneItemFiveItemsIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    myResult = myList.HasOneItem
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 31 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Has")
Private Sub Test32_HasOneItemOneItemIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10)
    
    myResult = myList.HasOneItem
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 32 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test33_HasOneItemNoItemsIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb
    
    myResult = myList.HasOneItem
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 33 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test34_HasAnyItemsFiveItemsIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    myResult = myList.HasAnyItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 34 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Has")
Private Sub Test35_HasAnyItemsOneItemIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10)
    
    myResult = myList.HasAnyItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 35 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test36_HasAnyItemsNoItemsIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb
    
    myResult = myList.HasOneItem
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 36 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test37_HasAnyItemsFiveItemsIsFasle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    myResult = myList.HasNoItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 37 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Has")
Private Sub Test38_HasAnyItemsOneItemIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb.Add(10)
    
    myResult = myList.HasNoItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 38 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Has")
Private Sub Test39_HasAnyItemsNoItemsIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
   
    Dim myResult As Boolean
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb
    
    myResult = myList.HasNoItems
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 39 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IndexOf")
Private Sub Test40_Indexof30IndexIsThree()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 3
   
    Dim myResult As Long
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    myResult = myList.Indexof(30)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 40 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IndexOf")
Private Sub Test40_Indexof42IndexIsMinusOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = -1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    'Act:
    
    
    myResult = myList.Indexof(42)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 40 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IndexOf")
Private Sub Test41_Indexof30WithLimitsIndexIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Indexof(30, 4, 7)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 41 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IndexOf")
Private Sub Test42_Indexof42WithLimitsIndexIsMinusOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = -1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Indexof(42, 4, 7)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 42 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IndexOf")
Private Sub Test43_Indexof30WithStartOnlyIndexIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Indexof(30, 4)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 43 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IndexOf")
Private Sub Test44_Indexof30WithEndOnlyIndexIsOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Indexof(30, ipEndIndex:=7)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 44 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("LastIndexOf")
Private Sub Test45_LastIndexof30IndexIsThree()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 3
   
    Dim myResult As Long
    Dim myList As WCollection
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    
    myResult = myList.LastIndexof(30)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 45 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LastIndexOf")
Private Sub Test46_LastIndexof42IsMinusOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = -1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    'Act:
    
    
    myResult = myList.LastIndexof(42)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 46 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LastIndexOf")
Private Sub Test47_LastIndexof30WithLimitsIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.LastIndexof(30, 4, 7)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 47 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LastIndexOf")
Private Sub Test48_LastIndexof42WithLimitsIsMinusOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = -1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.LastIndexof(42, 4, 7)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 48 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LastIndexOf")
Private Sub Test49_LastIndexof30WithStartOnlyIndexIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 9
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.LastIndexof(30, 4)
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 49 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub Test50a_ToArrayDefaultRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 50 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub Test50b_ToArrayDefaultStartSpecifiedEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.ToArray(ipEndIndex:=4)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("ToArray")
Private Sub Test50c_ToArraySpecifiedStartDefaultEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(20, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.ToArray(4)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub Test50d_ToArraySpecifiedStartSpacifiedEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.ToArray(3, 7)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub Test50e_ToArrayNegativeSpecifiedStartNegativeSpecifiedEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.ToArray(-7, -3)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Insert")
Private Sub Test51a_InsertSingleItemPositiveIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 42, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Insert(5, 42).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Insert")
Private Sub Test51b_InsertSingleItemNegativeIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 42, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Insert(-5, 42).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Insert")
Private Sub Test51c_InsertMultipleItemsPositiveIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 42, 43, 44, 45, 46, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Insert(5, 42, 43, 44, 45, 46).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Insert")
Private Sub Test51d_InsertMultipleItemsNegativeIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 42, 43, 44, 45, 46, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Insert(-5, 42, 43, 44, 45, 46).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 51d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("InsertRange")
Private Sub Test52a_InsertRangePositiveIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 42, 43, 44, 45, 46, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Insert(5, 42, 43, 44, 45, 46).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 53a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("InsertRange")
Private Sub Test52b_InsertRangeNegativeIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 42, 43, 44, 45, 46, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.Insert(-5, 42, 43, 44, 45, 46).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 53b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("InsertRange")
Private Sub Test53c_InsertStringWithSeparatorPositiveIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.AddString("Hello There Day", " ")
    
    'Act:
    myResult = myList.InsertString(3, "World Its A Nice", " ").ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 53c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("InsertString")
Private Sub Test53d_InsertStringWithSeparatorNegativeIndex()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.AddString("Hello There Day", " ")
    
    'Act:
    myResult = myList.InsertString(-1, "World Its A Nice", " ").ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 53d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Remove")
Private Sub Test54a_RemoveSingleItemPositiveIndex()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "There", "World", "A", "Nice", "Day")
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add("Hello", "There", "World", "Its", "A", "Nice", "Day")
    
    'Act:
    myResult = myList.Remove(4).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 54a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Remove")
Private Sub Test54b_RemoveSingleItemNegativeIndex()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "There", "World", "A", "Nice", "Day")
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add("Hello", "There", "World", "Its", "A", "Nice", "Day")
    
    'Act:
    myResult = myList.Remove(-4).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 54b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Remove")
Private Sub Test54c_RemoveMultipleItemsPositiveIndex()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "There", "Nice", "Day")
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add("Hello", "There", "World", "Its", "A", "Nice", "Day")
    
    'Act:
    myResult = myList.Remove(3, 5).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 54c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Remove")
Private Sub Test54d_RemoveMultipleItemsMixedIndexes()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "There", "Nice", "Day")
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add("Hello", "There", "World", "Its", "A", "Nice", "Day")
    
    'Act:
    myResult = myList.Remove(3, -3).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 54d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Remove")
Private Sub Test55_RemoveAll()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 0
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add("Hello", "There", "World", "Its", "A", "Nice", "Day")
    
    'Act:
    myResult = myList.RemoveAll.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 55 raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveFirstOf")
Private Sub Test56a_RemoveFirstOfDefaultRange()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 10, 20, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveFirstOf(30).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 56a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveFirstOf")
Private Sub Test56b_RemoveFirstOfSpecifiedStartDefaultEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveFirstOf(30, 3).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 56b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("RemoveFirstOf")
Private Sub Test56c_RemoveFirstOfDefaultStartSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveFirstOf(30, ipEndIndex:=7).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 56b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveFirstOf")
Private Sub Test56d_RemoveFirstOfSpecifiedStartSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveFirstOf(30, 3, 6).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 56d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveFirstOf")
Private Sub Test56e_RemoveFirstOfNegativeSpecifiedStartNegativeSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveFirstOf(30, -5, -2).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 56e raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("RemoveFirstOf")
Private Sub Test57a_RemoveLastOfDefaultRange()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 30, 40, 50, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveLastOf(30).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 57a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveLastOf")
Private Sub Test57b_RemoveLastOfSpecifiedStartDefaultEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 30, 40, 50, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveLastOf(30, 5).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 57b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("RemoveLastOf")
Private Sub Test57c_RemoveLastOfDefaultStartSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveLastOf(30, ipEndIndex:=6).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 57c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveLastOf")
Private Sub Test57d_RemoveLastOfSpecifiedStartSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveLastOf(30, 2, 6).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 57d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveLastOf")
Private Sub Test57e_RemoveLastOfNegativeSpecifiedStartNegativeSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveLastOf(30, -7, -3).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 57e raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveAny")
Private Sub Test58a_RemoveAnyOfDefaultRange()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 40, 50)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveAnyOf(30).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 58a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveAny")
Private Sub Test58b_RemoveAnyOfSpecifiedStartDefaultEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(30, 30, 10, 20, 40, 50)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(30, 30, 10, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveAnyOf(30, 5).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 58b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("RemoveAny")
Private Sub Test58c_RemoveAnyOfDefaultStartSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 40, 50, 30, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveAnyOf(30, ipEndIndex:=6).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 58c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveAny")
Private Sub Test58d_RemoveAnyOfSpecifiedStartSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveAnyOf(30, 3, 7).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 58d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RemoveAny")
Private Sub Test58e_RemoveAnyOfNegativeSpecifiedStartNegativeSpecifiedEnd()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb(10, 30, 20, 30, 40, 50, 30, 30)
    
    'Act:
    myResult = myList.RemoveAnyOf(30, -6, -2).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 58e raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("FirstIndex")
Private Sub Test59a_FirstIndexEmptyWCollectionIsMinusOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = -1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    'Act:
    myResult = myList.FirstIndex
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 59a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FirstIndex")
Private Sub Test59b_FirstIndexFilledWCollectionIsOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.FirstIndex
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 59b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LastIndex")
Private Sub Test60a_LastIndexEmptyWCollectionIsMinusOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = -1
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    'Act:
    myResult = myList.LastIndex
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 60a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LastIndex")
Private Sub Test60b_LastIndexFilledWCollectionIsOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.LastIndex
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 60b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Count")
Private Sub Test61a_CountEmptyWCollectionIsMinusOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 0
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    'Act:
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 61a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Count")
Private Sub Test61b_CountFilledWCollectionIsOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
   
    Dim myResult As Long
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 61b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("First")
Private Sub Test62a_FirstItemEmptyCollection()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Empty
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    'Act:
    myResult = myList.First
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 62a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("First")
Private Sub Test62b_FirstItemFilledCollection()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 10
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.First
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 62a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Last")
Private Sub Test63a_LastItemEmptyCollection()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Empty
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb
    
    'Act:
    myResult = myList.First
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 62a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Last")
Private Sub Test63b_LastItemEmptyCollection()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Last
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 63braised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Enqueue")
Private Sub Test64a_EnqueueSingleParameter()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Enqueue(60).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 64a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Enqueue")
Private Sub Test64b_EnqueueMultipleParameters()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Enqueue(60, 70, 80, 90).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 64braised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Enqueue")
Private Sub Test64c_EnqueueCollectionCountIncrementsByOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 6
   
    Dim myTest As Collection
    Set myTest = New Collection
    
    With myTest
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        
    End With
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Enqueue(myTest).Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 64c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EnqueueRange")
Private Sub Test65a_EnqueueRangeSingleItemArray()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.EnqueueRange(Array(60)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 64a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EnqueueRange")
Private Sub Test65b_EnqueueMultipleItemArray()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.AddRange(Array(10, 20, 30, 40, 50))
    
    'Act:
    myResult = myList.EnqueueRange(Array(60, 70, 80, 90)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 65b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EnqueueString")
Private Sub Test66a_EnqueueStringHelloIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    
    Dim myList As WCollection
    Dim myResult As Long
    'Act:
    
    Set myList = WCollection.Deb.EnqueueString("Hello")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 66a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EnqueueString")
Private Sub Test66b_EnqueueStringSeparatedStringWithNoSeparatorCountIs26()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 26
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.EnqueueString("Hello World Its A Nice Day")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 66b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EnqueueString")
Private Sub Test66c_EnqueueSeparatedStringWithSeparatorCountIsSix()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 6
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.EnqueueString("Hello World Its A Nice Day", " ")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 66c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Dequeue")
Private Sub Test67a_DequeueSingleItem()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Dequeue
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 67a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Dequeue")
Private Sub Test67b_DequeueThreeItems()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Dequeue(3)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 67b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Dequeue")
Private Sub Test67c_DequeueZeroItems()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Empty
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Dequeue(0)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 67c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod ("Dequeue")
Private Sub Test67d_Dequeue10ItemsFromQueueOfFiveItems()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Dequeue(10)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 67d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Push")
Private Sub Test68a_PushSingleParameter()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Push(60).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 67a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Push")
Private Sub Test68b_PushMultipleParameters()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Push(60, 70, 80, 90).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 68b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Push")
Private Sub Test68c_PushCollectionCountIncrementsByOne()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 6
   
    Dim myTest As Collection
    Set myTest = New Collection
    
    With myTest
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        
    End With
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.Push(myTest).Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 68c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PushRange")
Private Sub Test69a_PushRangeSingleItemArray()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50)
    
    'Act:
    myResult = myList.PushRange(Array(60)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 69a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PushRange")
Private Sub Test69b_PushRangeMultipleItemArray()
     On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
   
    Dim myResult As Variant
    Dim myList As WCollection
    Set myList = WCollection.Deb.AddRange(Array(10, 20, 30, 40, 50))
    
    'Act:
    myResult = myList.PushRange(Array(60, 70, 80, 90)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 69b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PushString")
Private Sub Test70a_PushtringHelloIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 5
    
    Dim myList As WCollection
    Dim myResult As Long
    'Act:
    
    Set myList = WCollection.Deb.PushString("Hello")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 70a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PushString")
Private Sub Test70b_PushStringSeparatedStringWithNoSeparatorCountIs26()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 26
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 70b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PushString")
Private Sub Test70c_PushSeparatedStringWithSeparatorCountIsSix()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 6
    Dim myList As WCollection
    
    Dim myResult As Long
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Count
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 70c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Pop")
Private Sub Test71a_PopSingleItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Day")
    Dim myList As WCollection
    
    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Pop
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 71a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Pop")
Private Sub Test71b_PopMultipleItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Day", "Nice", "A")
    Dim myList As WCollection
    
    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Pop(3)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 71b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Pop")
Private Sub Test71c_PopZeroItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Empty
    Dim myList As WCollection
    
    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Pop(0)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 71c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Pop")
Private Sub Test71d_Pop10ItemsFromSix()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Day", "Nice", "A", "Its", "World", "Hello")
    Dim myList As WCollection
    
    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Pop(10)
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 71d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Peek")
Private Sub Test72a_PeekDefault()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello"
    Dim myList As WCollection
    
    Dim myResult As String
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Peek
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 72a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Peek")
Private Sub Test72b_PeekIndexThree()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Its"
    Dim myList As WCollection
    
    Dim myResult As String
    'Act:
    Set myList = WCollection.Deb.PushString("Hello World Its A Nice Day", " ")
    myResult = myList.Peek(3)
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 72b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Dedup")
Private Sub Test73a_DedupNoDuplicates()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50)
    Dim myList As WCollection
    
    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb(10, 20, 30, 40, 50)
    myResult = myList.Dedup.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 73a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Dedup")
Private Sub Test73b_DedupDuplicates()
    On Error GoTo TestFail
    
   'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 20, 40, 50)
    Dim myList As WCollection
    
    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb(10, 30, 30, 20, 30, 40, 20, 50)
    myResult = myList.Dedup.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 73b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Test74a_SortNumbers()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50)
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb(10, 30, 20, 50, 40)
    myResult = myList.Sort.ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 73b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
'
'@TestMethod("Sort")
Private Sub Test74b_SortStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("A", "Day", "Hello", "Its", "Nice", "World")
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.AddString("Hello World Its A Nice Day", " ")
    myResult = myList.Sort.ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 74b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfUnique")
Private Sub Test75A_SetOfUniqueNumbers()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 70, 60, 40, 50, 80)
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 70, 60)
    myResult = myList.SetOfUnique(Array(20, 40, 50, 60, 80)).ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 74b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfUnique")
Private Sub Test75b_SetOfUniqueStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Its", "A", "Terrible", "Day", "This", "Hello", "World", "Nice")
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.AddString("Its A Terrible Day This Day", " ")
    myResult = myList.SetOfUnique(WCollection.Deb.AddString("Hello World Its A Nice Day", " ").ToArray).ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 75b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfCommon")
Private Sub Test76A_CommonNumbers()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(20, 60)
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 70, 60)
    myResult = myList.SetOfCommon(Array(20, 40, 50, 60, 80)).ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 76a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfCommon")
Private Sub Test76b_CommonStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Its", "A", "Day")
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.AddString("Its A Terrible Day This Day", " ")
    myResult = myList.SetOfCommon(WCollection.Deb.AddString("Hello World Its A Nice Day", " ").ToArray).ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 76b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfHostOnly")
Private Sub Test77A_SettOfHostOnlyNumbers()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 30, 70)
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 70, 60)
    myResult = myList.SetOfHostOnly(Array(20, 40, 50, 60, 80)).Sort.ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 77a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfHostOnly")
Private Sub Test77b_SetOfHostOnlyStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = WCollection.Deb("Terrible", "This").Sort.ToArray
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.AddString("Its A Terrible Day This Day", " ")
    myResult = myList.SetOfHostOnly(WCollection.Deb.AddString("Hello World Its A Nice Day", " ").ToArray).Sort.ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 77b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfInputOnly")
Private Sub Test78A_SettOfInputOnlyNumbers()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(40, 50, 80)
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 70, 60)
    myResult = myList.SetOfInputOnly(Array(20, 40, 50, 60, 80)).Sort.ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 78a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SetOfInputOnly")
Private Sub Test78b_SetOfInputOnlyStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = WCollection.Deb("Hello", "World", "Nice").Sort.ToArray
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.AddString("Its A Terrible Day This Day", " ")
    myResult = myList.SetOfInputOnly(WCollection.Deb.AddString("Hello World Its A Nice Day", " ")).Sort.ToArray

    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 78b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MapIt")
Private Sub Test79a_MapIt_mpInc()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(11, 21, 31, 41, 51, 61, 71)
    Dim myList As WCollection

    Dim myResult As Variant
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50, 60, 70)
    myResult = myList.MapIt(mpInc()).ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 79a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MapIt")
Private Sub Test79b_MapIt_mpInc_WithNan()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(11, 21, 31, "nan", 51, 61, 71)
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, "Hello", 50, 60, 70)
    myResult = myList.MapIt(mpInc()).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 79b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MapIt")
Private Sub Test79c_MapIt_mpInc_5WithNan()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(15, 25, 35, "nan", 55, 65, 75)
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, "Hello", 50, 60, 70)
    myResult = myList.MapIt(mpInc(5)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 79c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("MapIt")
Private Sub Test79d_MapIt_mpToChars()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("H", "e", "l", "l", "o")
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "Hello", "Hello", "Hello", "Hello")
    myResult = myList.MapIt(mpToChars).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult(0).ToArray
    Assert.SequenceEquals myExpected, myResult(1).ToArray
    Assert.SequenceEquals myExpected, myResult(2).ToArray
    Assert.SequenceEquals myExpected, myResult(3).ToArray
    Assert.SequenceEquals myExpected, myResult(4).ToArray
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 79d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FilterIt")
Private Sub Test80a_CmpIt_IsEqualTo()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "Hello", "Hello")
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "Sailor", "Hello", "There", "Hello", "Dear")
    myResult = myList.FilterIt(cmpIt(Comparison.IsEqualTo, "Hello")).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 80a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("FilterIt")
Private Sub Test80b_CmpIt_IsEqualTo()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Sailor", "There", "Dear")
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "Sailor", "Hello", "There", "Hello", "Dear")
    myResult = myList.FilterIt(cmpIt(Comparison.IsNotEqualTo, "Hello")).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 80b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("FilterIt")
Private Sub Test80c_CmpIt_IsEqualTo_IndexTwo()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "Hello", "Hello", "Dear")
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "Sailor", "Hello", "There", "Hello", "Dear")
    myResult = myList.FilterIt(cmpIt(Comparison.IsEqualTo, "e", 2)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 80c raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("FilterIt")
Private Sub Test80d_CmpMapIt_mpLen_MoreThanFour()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "Sailor", "Hello", "There", "Hello")
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "Its", "Sailor", "A", "Hello", "Day", "There", "Hello", "Dear")
    myResult = myList.FilterIt(CmpMapIt(mpLen, Comparison.IsMoreThan, 4)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 80d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("FilterIt")
Private Sub Test80d_CmpMapIt_mpLen_IsNotLessThanFour()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("Hello", "Sailor", "Hello", "There", "Hello", "Dear")
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add("Hello", "Its", "Sailor", "A", "Hello", "Day", "There", "Hello", "Dear")
    myResult = myList.FilterIt(CmpMapIt(mpLen, Comparison.IsNotLessThan, 4)).ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 80d raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ReduceIt")
Private Sub Test81a_ReduceIt_Sum()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant
    myExpected = 210
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 30, 40, 50, 60)
    myResult = myList.ReduceIt(rdSum)
    
    'Assert:
    Assert.AreEqual myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 81a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsSameOrder")
Private Sub Test82a_IsSameOrderIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 40, 30, 50, 60)
    myResult = myList.Sort.IsSameOrder(Array(10, 20, 30, 40, 50, 60))
    
    'Assert:
    Assert.AreEqual myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 82a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsSameOrder")
Private Sub Test82b_IsSameOrderIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(10, 20, 40, 30, 50, 60)
    myResult = myList.IsSameOrder(Array(10, 20, 30, 40, 50, 60))
    
    'Assert:
    Assert.AreEqual myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 82b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsEquivalent")
Private Sub Test83a_IsEquivalentIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(20, 10, 40, 30, 60, 50)
    myResult = myList.Sort.IsEquivalent(Array(10, 20, 30, 40, 50, 60))
    
    'Assert:
    Assert.AreEqual myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 83a raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsEquivalent")
Private Sub Test83b_IsEquivalentIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    Dim myList As WCollection

    Dim myResult As Variant
    
    'Act:
    Set myList = WCollection.Deb.Add(20, 10, 40, 60, 50, 40)
    myResult = myList.IsEquivalent(Array(10, 20, 30, 40, 50, 60))
    
    'Assert:
    Assert.AreEqual myExpected, myResult


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test 83b raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

