VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormS 
   Caption         =   "T&TDG Structural Database Family Management"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20670
   OleObjectBlob   =   "UserFormStructure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lrB As Long
Public lrP As Long
Public lrAO As Long
Public lrAQ As Long
Public IsSubFolder As Boolean
Public FSO As Scripting.FileSystemObject
Public SourceFolder As Scripting.Folder, SubFolder As Scripting.Folder

Private Sub cboCate_Change()
    Worksheets("S").Range("C2") = Me.cboCate.Value
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R2C3,R4C3:R2000C4,2,FALSE),"""")"
    Me.txtACate.Value = Worksheets("S").Range("D2")
End Sub

Private Sub cboCheck_Click()
Application.ScreenUpdating = False
        Dim i As Long
        lrP = ActiveSheet.Range("P" & Rows.Count).End(xlUp).Row
        lrAO = ActiveSheet.Range("AO" & Rows.Count).End(xlUp).Row
           i = 4
        Do While i <= lrP
        If ActiveSheet.Range("W" & i).Value <> ActiveSheet.Range("BB" & i).Value Then
            MsgBox "Searching data and family in folder are difference in line " & i - 3 & " number."
            Me.opbCheck.Value = False
            Exit Sub
        Else
           i = i + 1
        End If
        Loop
            MsgBox "Searching data and family in database folder are MATCH"
            Me.opbCheck.Value = True
       'Sau khi data MATCH, kick chuot 2 dong moi nhay song song
 Application.ScreenUpdating = False
End Sub

Private Sub cboCopy_Click()
If Me.cboFamily.Value = "" Or Me.cboCate.Value = "" Then
Exit Sub
Else

Worksheets("S").Range("M4").Copy
End If
End Sub

Private Sub cboFamily_Change()
If Me.cboFamily.Value = "System" Then
   Me.txtName.Locked = False
Else: Me.cboFamily.Value = "Loadable"
   Me.txtName.Locked = True
End If
End Sub

Private Sub cboSelect_Change()
ListBox1.RowSource = ""

End Sub
Private Sub cboSort_Change()
ListBox1.RowSource = ""
End Sub

Private Sub cboSub_Change()
    Worksheets("S").Range("G2") = Me.cboSub.Value
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],R[2]C[-1]:R[1000]C,2,FALSE),"""")"
    Me.txtASub.Value = Worksheets("S").Range("H2")
End Sub

Private Sub cboType_Change()
    Worksheets("S").Range("E2") = Me.cboType.Value
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],R[2]C[-1]:R[1000]C,2,FALSE),"""")"
    Me.txtAType.Value = Worksheets("S").Range("F2")
End Sub

Private Sub cboUndo_Click()
Application.ScreenUpdating = False
'Dim findvalue As Range
On Error GoTo Existed

If Me.txtEditFile.Text = "" Or Me.ListBox2 = "" Then
   Exit Sub
Else
   Dim OldFile As String
   Dim NewFile As String
   
    OldFile = Worksheets("S").Range("AT5")
    NewFile = Worksheets("S").Range("AT4")
Existed:
    If Worksheets("S").Range("AT4").Value = Worksheets("S").Range("AT5").Value Then
  
     Exit Sub
    Else
    On Error Resume Next
    Name OldFile As NewFile
    End If
End If
On Error GoTo 0
Call Reget_file
MsgBox "Undo Successful!"
Me.cboUndo.Enabled = False
Me.opbCheck.Value = False
Application.ScreenUpdating = True
End Sub

Private Sub cmdClear_Click()
Application.ScreenUpdating = False
  Unload Me
  ThisWorkbook.Sheets("S").Unprotect Password:="ttdg"
  Worksheets("S").Activate
  UserFormS.Show
Application.ScreenUpdating = True

End Sub
Private Sub cmdDel_Click()
Application.ScreenUpdating = False
 If Me.cboFamily.Value = "" Or Me.cboCate.Value = "" Then
Exit Sub
Else
 ''''1.Password cho button
 Dim PassProtect As Variant
 PassProtect = InputBox(Prompt:="Please enter the password to unlock this button." & vbCrLf & "(Only delete when you are Administrator)", Title:="Warning")

If PassProtect = vbNullString Then Exit Sub

If PassProtect <> "ttdg" Then
MsgBox Prompt:="Incorrect password. Please try again.", Buttons:=vbOKOnly
  Exit Sub
  
Else
 Dim i As Long ''Ghi nho row da xoa de sau do quay tro lai
 lrP = Worksheets("S").Range("P" & Rows.Count).End(xlUp).Row
  i = Worksheets("S").Range("W4:W" & lrP).Find(What:=Me.txtName, LookIn:=xlValues).Row - 4
Call del
End If
Call cmdSearch_Click
End If
 ActiveWindow.ScrollColumn = 1
 Call SaveNew
 Me.opbCheck.Value = False
Application.ScreenUpdating = True
End Sub
Sub del()
Application.ScreenUpdating = False
''''2. Bat dau lenh delete
On Error GoTo cmdDel_Click_Error
 If Me.cboCate.Value = "" And Me.txtName.Value = "" Then
   Call MsgBox("Double click on Family so It can be deleted", vbInformation, "Delete Contact")
   Exit Sub
 End If
 
 'Thong bao chon MsgBox
 Dim kt As VbMsgBoxResult
 kt = MsgBox("You are about to delete this Family", vbYesNo)
 If kt = vbNo Then
    Exit Sub
 ElseIf kt = vbYes Then
 
Dim findvalue As Range
lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
Set findvalue = Worksheets("S").Range("I4:I" & lrB).Find(What:=Me.txtName, LookIn:=xlValues)
findvalue.Value = ""
findvalue.Offset(0, 2).Value = ""
findvalue.Offset(0, 1).Value = ""
findvalue.Offset(0, -1).Value = ""
findvalue.Offset(0, -2).Value = ""
findvalue.Offset(0, -3).Value = ""
findvalue.Offset(0, -4).Value = ""
findvalue.Offset(0, -5).Value = ""
findvalue.Offset(0, -6).Value = ""
findvalue.Offset(0, -7).Value = ""
findvalue.Offset(0, -8).Value = ""
clearList
Call SortData
 End If
On Error GoTo 0
Exit Sub
'if error occurs then show me exactly where the error occurs
cmdDel_Click_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdDel_Click"
Application.ScreenUpdating = True
End Sub
Sub clearList()
Application.ScreenUpdating = False
Me.cboFamily.Value = ""
Me.cboCate.Value = ""
Me.txtACate.Value = ""
Me.cboType.Value = ""
Me.txtAType.Value = ""
Me.cboSub.Value = ""
Me.txtASub.Value = ""
Me.txtName.Value = ""
Me.txtNumber.Value = ""
Application.ScreenUpdating = True
End Sub

Private Sub cmdEdit_Click()
Application.ScreenUpdating = False
If Me.cboFamily.Value = "" Or Me.cboCate.Value = "" Then
    Exit Sub
''*xet xem da co name bi trung
    Dim findName As Range
    lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
    Set findName = Worksheets("S").Range("I4:I" & lrB).Find(What:=Me.txtName, LookIn:=xlValues)
ElseIf Me.txtPreview.Value = Me.txtName.Value Then
    MsgBox "Family Name Existed"
    Exit Sub

Else
    Dim findvalue As Range
    lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
    Set findvalue = Sheets("S").Range("K4:K" & lrB).Find(What:=Me.txtID, LookIn:=xlValues)
''findvalue.value = Me.txtID ''''DANG CHO KHONG SUA ID, VI SUA THONG TIN CUA CHINH ID CU
findvalue.Offset(0, -1).Value = Me.txtNumber.Value
findvalue.Offset(0, -3).Value = Me.txtASub.Value
findvalue.Offset(0, -4).Value = Me.cboSub.Value
findvalue.Offset(0, -5).Value = Me.txtAType.Text
findvalue.Offset(0, -6).Value = Me.cboType.Value
findvalue.Offset(0, -7).Value = Me.txtACate.Value
findvalue.Offset(0, -8).Value = Me.cboCate.Value
findvalue.Offset(0, -9).Value = Me.cboFamily.Value
Call cmdSearch_Click
End If
 Call SortData
 '' Click ve listrow moi thao tac
   Dim i As Long
 On Error Resume Next
 lrP = Worksheets("S").Range("P" & Rows.Count).End(xlUp).Row
 i = Worksheets("S").Range("W4:W" & lrP).Find(What:=Me.txtPreview, LookIn:=xlValues).Row - 4
   If err.Number <> 0 Then
   MsgBox "You have not chosen any name to Edit"
   Exit Sub
   End If
 ListBox1.Selected(i) = True
 ''*add new sang sheet source
 Worksheets("S_Source").Activate
 Call SaveNew
 Worksheets("S").Activate
 Me.opbCheck.Value = False
 Application.ScreenUpdating = True
 
End Sub

Private Sub cmdFolder_Click()
Application.ScreenUpdating = False
Dim findvalue As Range
On Error GoTo Existed
lrAO = Worksheets("S").Range("AO" & Rows.Count).End(xlUp).Row
Set findvalue = Sheets("S").Range("AO4:AO" & lrAO).Find(What:=Me.txtEditFile, LookIn:=xlValues)
If Me.txtEditFile.Text = "" Then
   Exit Sub
Else
   Dim OldFile As String
   Dim NewFile As String
   Worksheets("S").Range("AS4:AS5").Value = Me.txtEditFile.Text
    OldFile = Worksheets("S").Range("AT2")
    NewFile = Worksheets("S").Range("AT3")
Existed:
    If Worksheets("S").Range("AT2").Value = Worksheets("S").Range("AT3").Value Then
    MsgBox "Name already has existed"
     Exit Sub
    Else
    On Error Resume Next
    Name OldFile As NewFile
    End If
End If
On Error GoTo 0
Call Reget_file
''''Chon ve row vua edit name
    On Error Resume Next
    Dim i As Long
    lrAO = ActiveSheet.Range("AO" & Rows.Count).End(xlUp).Row
    i = ActiveSheet.Range("AO4:AO" & lrAO).Find(What:=Me.txtEditFile, LookIn:=xlValues).Row - 4
     If err.Number <> 0 Then
      MsgBox "Check family name again"
      Exit Sub
     End If
    ListBox2.Selected(i) = True
''''''
Application.ScreenUpdating = True
Me.cboUndo.Enabled = True
Me.opbCheck.Value = False
End Sub

Private Sub cmdGetFile_Click()
Application.ScreenUpdating = False
''Xoa Get list file cu
    Worksheets("S").Range("AO:AO").ClearContents
    Dim myPath As String
    myPath = Me.txtPath.Text
    If myPath <> "" Then
        Set FSO = New Scripting.FileSystemObject
        If FSO.FolderExists(myPath) <> False Then
                Set SourceFolder = FSO.GetFolder(myPath)
                IsSubFolder = True
                Call ListFilesInFolder(SourceFolder, IsSubFolder)
        Else
        MsgBox "Selected Path Incorrect", vbInformation, "T&T Design Group"
        Exit Sub
        End If
    Else
    MsgBox "Folder Path Empty !!" & vbNewLine & vbNewLine & "", vbInformation, "T&T Design Group"
    Exit Sub
    End If
'*Sort va Filter du lieu trc khi dien
    Dim lrAO As Long
    lrAO = Worksheets("S").Range("AO" & Rows.Count).End(xlUp).Row
    Range("AO4:AO" & lrAO).Select
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Add Key:=Range("AO4:AO" & lrAO), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("S").Sort
        .SetRange Range("AO4:AO" & lrAO)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ''****KHAC NHAU O ADD2 VA ADD TRONG MS 365 VA MS 2016
    'ActiveWorkbook.Worksheets("S").Sort.SortFields.Add2 Key:=Range("AO4"), _
        'SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'*Dien du lieu vao Listrow
   ListBox2.List = Worksheets("S").Range("AN4:AO" & lrAO).Value
   Me.cboCheck.Enabled = True
   Me.opbCheck.Value = False
 Application.ScreenUpdating = True
End Sub
Sub Reget_file()
Application.ScreenUpdating = False
''Xoa Get list file cu
    Worksheets("S").Range("AO:AO").ClearContents
    Dim myPath As String
    myPath = Me.txtPath.Text
    If myPath <> "" Then
        Set FSO = New Scripting.FileSystemObject
        If FSO.FolderExists(myPath) <> False Then
                Set SourceFolder = FSO.GetFolder(myPath)
                IsSubFolder = True
                Call ListFilesInFolder(SourceFolder, IsSubFolder)
        Else
        MsgBox "Selected Path Incorrect", vbInformation, "T&T Design Group"
        Exit Sub
        End If
    Else
    MsgBox "Folder Path Empty !!" & vbNewLine & vbNewLine & "", vbInformation, "T&T Design Group"
    Exit Sub
    End If
'*Sort va Filter du lieu trc khi dien
    Dim lrAO As Long
    lrAO = Worksheets("S").Range("AO" & Rows.Count).End(xlUp).Row
    Range("AO4:AO" & lrAO).Select
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Add Key:=Range("AO4"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("S").Sort
        .SetRange Range("AO4:AO" & lrAO)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ''****KHAC NHAU O ADD2 VA ADD TRONG MS 365 VA MS 2016
    'ActiveWorkbook.Worksheets("S").Sort.SortFields.Add2 Key:=Range("AO4"), _
        'SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'*Dien du lieu vao Listrow
   ListBox2.List = Worksheets("S").Range("AN4:AO" & lrAO).Value
   Me.cboCheck.Enabled = True
   Me.opbCheck.Value = False
End Sub
Sub ListFilesInFolder(ByVal xFolderName As String, ByVal xIsSubfolders As Boolean)
Dim xFileSystemObject As Object
Dim xFolder As Object
Dim xSubFolder As Object
Dim xFile As Object
Dim rowIndex As Long
Set xFileSystemObject = CreateObject("Scripting.FileSystemObject")
Set xFolder = xFileSystemObject.GetFolder(xFolderName)
rowIndex = Application.ActiveSheet.Range("AO1000").End(xlUp).Row + 3
For Each xFile In xFolder.files
  Application.ActiveSheet.Cells(rowIndex, 41).Value = xFile.Name
  rowIndex = rowIndex + 1
Next xFile
If xIsSubfolders Then
  For Each xSubFolder In xFolder.SubFolders
    ListFilesInFolder xSubFolder.Path, True
  Next xSubFolder
End If
Set xFile = Nothing
Set xFolder = Nothing
Set xFileSystemObject = Nothing
End Sub
Private Sub cmdPath_Click()
Worksheets("S").Range("AO:AO") = ""
'declaring variable to store path
Dim get_Path As String
With Application.FileDialog(msoFileDialogFolderPicker)
If .Show <> 0 Then
get_Path = .SelectedItems(1)
End If
'MsgBox (Get_Path)
Worksheets("S").Cells(1, 46).Value = get_Path
txtPath.Text = get_Path
End With
Me.opbCheck.Value = False
End Sub

Private Sub cmdSearch_Click()
Application.ScreenUpdating = False
Me.opbCheck.Value = False
On Error GoTo errHandler:

 Worksheets("S").Range("M2") = Me.cboSelect.Value
 Worksheets("S").Range("M3") = Me.txtSearch.Value
 lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
 Worksheets("S").Range("B3:K" & lrB).AdvancedFilter Action:=xlFilterCopy, _
 CriteriaRange:=Range("S!Criteria"), _
 CopyToRange:=Range("S!Extract"), Unique:=False
 ''tim dong cuoi cua du lieu ket qua de dien vao listrow
 lrP = Worksheets("S").Range("P" & Rows.Count).End(xlUp).Row
 ListBox1.RowSource = Worksheets("S").Range("P4:Y" & lrP).Address
 Me.txtMax.Value = Worksheets("S").Range("L1")
 Me.txtNumbers.Value = Worksheets("S").Range("L2")
  
Exit Sub
errHandler:
 MsgBox "No Sort Field OR Sort Field and Search do not match"
 
 Application.ScreenUpdating = True
End Sub
Private Sub cmdSearch_Enter()
Call cmdSearch_Click
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub cmdAdd_Click()
Application.ScreenUpdating = False
If Me.cboFamily.Value = "" Or Me.cboCate.Value = "" Then
    Exit Sub
Else
'1. Dien du lieu moi
''BAO LOI KHI NHAP THONG TIN DUPLICATE
    Dim findName As Range
    lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
    Set findName = Worksheets("S").Range("I4:I" & lrB).Find(What:=Me.txtName, LookIn:=xlValues)
    On Error Resume Next
  If Me.txtName.Value = Me.txtPreview.Value Or findName.Value = Me.txtPreview.Value Then
    ''MsgBox "Family Name Existed"
    Exit Sub
  Else
    Dim Drng As Range
    Set Drng = Worksheets("S").Range("B4")
Drng.End(xlDown).Offset(1, 0).Value = Me.cboFamily.Value
Drng.End(xlDown).Offset(0, 1).Value = Me.cboCate.Value
Drng.End(xlDown).Offset(0, 2).Value = Me.txtACate.Value
Drng.End(xlDown).Offset(0, 3).Value = Me.cboType.Value
Drng.End(xlDown).Offset(0, 4).Value = Me.txtAType.Value
Drng.End(xlDown).Offset(0, 5).Value = Me.cboSub.Value
Drng.End(xlDown).Offset(0, 6).Value = Me.txtASub.Value
Drng.End(xlDown).Offset(0, 8).Value = Me.txtNumber.Value
Drng.End(xlDown).Offset(0, 9).Value = Worksheets("S").Range("L3").Value + 1 ''*ID luon + 1 (RAT QUAN TRONG)
' 2.COPY CONG THUC CUA O BEN TREN
    Drng.End(xlDown).Offset(-1, 7).Copy
    Drng.End(xlDown).Offset(0, 7).PasteSpecial Paste:=xlPasteFormulas

    Call SortData
    Call cmdSearch_Click
  End If
''''3.CHON VE DONG MOI ADD NEW
 Dim i As Long
 lrP = Worksheets("S").Range("P" & Rows.Count).End(xlUp).Row
 i = Worksheets("S").Range("W4:W" & lrP).Find(What:=Worksheets("S").Range("M4"), LookIn:=xlValues).Row - 4
 ListBox1.Selected(i) = True
 ''DIEN DU LIEU O DONG VUA MOI ADDNEW VAO
    ''Dien thong tin vao cac text box
    Me.cboFamily.Value = Me.ListBox1.Value
    Me.cboCate.Value = Me.ListBox1.Column(1)
    Me.txtACate.Value = Me.ListBox1.Column(2)
    Me.cboType.Value = Me.ListBox1.Column(3)
    Me.txtAType.Value = Me.ListBox1.Column(4)
    Me.cboSub.Value = Me.ListBox1.Column(5)
    Me.txtASub.Value = Me.ListBox1.Column(6)
    Me.txtName.Value = Me.ListBox1.Column(7)
    Me.txtNumber.Value = Me.ListBox1.Column(8)
    Me.txtID.Value = Me.ListBox1.Column(9)
 
End If
 Call SaveNew
 Me.cmdEdit.Enabled = False ''sau khi add khoa edit vi muon edit phai chon name
 Me.opbCheck.Value = False
Application.ScreenUpdating = True
End Sub
Sub SortData()

   ''1. Sort IndataS
   lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
    Worksheets("S").Range("A3:K" & lrB).Select
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Add Key:=Range("B4:B" & lrB), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    ActiveWorkbook.Worksheets("S").Sort.SortFields.Add Key:=Range("I4:I" & lrB), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("S").Sort
        .SetRange Range("A3:K" & lrB)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  ''''2. Sort Outdata
  lrP = Worksheets("S").Range("P" & Rows.Count).End(xlUp).Row
    Worksheets("S").Range("P3:Y" & lrP).Select
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Add Key:=Range("P4:P" & lrP), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    ActiveWorkbook.Worksheets("S").Sort.SortFields.Add Key:=Range("W4:W" & lrP), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("S").Sort
        .SetRange Range("P3:Y" & lrP)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
''''3.Sort phan system family
''''Tim dong cuoi cung cua system family
    Dim n As Long
    ActiveSheet.Range("BA1").Value = "=MATCH(""Loadable"",B:B,0)-1"
    n = ActiveSheet.Range("BA1").Value
''''Sort vung du lieu System Family
    ActiveSheet.Range("B3:K" & n).Select ''Sheet A va S la K, sheet MEPF la M
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("C4:C" & n), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("B3:K" & n)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub
Private Sub ListBox1_Click()
If Me.opbCheck.Value = False Then
    Exit Sub
Else
    Dim i As Long
    i = ListBox1.ListIndex
    ListBox2.Selected(i) = True
End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.ScreenUpdating = False
Me.cmdAdd.Enabled = True
Me.cmdEdit.Enabled = True
''Dien thong tin vao cac box
Me.cboFamily.Value = Me.ListBox1.Value
Me.cboCate.Value = Me.ListBox1.Column(1)
Me.txtACate.Value = Me.ListBox1.Column(2)
Me.cboType.Value = Me.ListBox1.Column(3)
Me.txtAType.Value = Me.ListBox1.Column(4)
Me.cboSub.Value = Me.ListBox1.Column(5)
Me.txtASub.Value = Me.ListBox1.Column(6)
Me.txtName.Value = Me.ListBox1.Column(7)
Me.txtNumber.Value = Me.ListBox1.Column(8)
Me.txtID.Value = Me.ListBox1.Column(9)
Application.ScreenUpdating = True
End Sub

Private Sub ListBox2_Click()
If Me.opbCheck.Value = False Then
    Exit Sub
Else
    Dim i As Long
    i = ListBox2.ListIndex
    ListBox1.Selected(i) = True
End If
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.ScreenUpdating = False
''Dien thong tin vao txtbox
Dim i As Long
With Me.ListBox2
    For i = 0 To .ListCount - 1
      If .Selected(i) Then
        Me.txtEditFile.Value = ListBox2.List(i, 1)
        Worksheets("S").Range("AQ4") = ListBox2.List(i, 1)
        Worksheets("S").Range("AQ5") = ListBox2.List(i, 1)
        Exit For
      End If
    Next
End With

Application.ScreenUpdating = True
End Sub
Sub PreviewName()
''PREVIEW EDIT NAME
If Me.cboFamily.Value = "System" Then
Me.txtPreview.Value = Me.txtName.Value

ElseIf Me.cboFamily.Value = "" And Me.txtACate.Value = "" And Me.cboType.Value _
= "" And Me.txtAType.Value = "" And Me.cboSub.Value = "" And Me.txtASub.Value = "" Then
  Me.txtPreview.Value = ""
''A.Xet TH Number <10
ElseIf Me.txtNumber.Value < 10 Then
  ''TH txtAsub bang - hoac bang ""
      If Me.txtASub.Value = "-" Or Me.txtASub.Value = "" Then
      Me.txtPreview.Value = "S_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_0" & Me.txtNumber.Value
      Else
      Me.txtPreview.Value = "S_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtASub.Value & "_0" & Me.txtNumber.Value
      End If
Else
''B.TH con lai Number>10
      If Me.txtASub.Value = "-" Or Me.txtASub.Value = "" Then
      Me.txtPreview.Value = "S_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtNumber.Value
      Else
      Me.txtPreview.Value = "S_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtASub.Value & "_" & Me.txtNumber.Value
      End If
End If
End Sub

Private Sub txtACate_Change()
Application.ScreenUpdating = False
Call PreviewName
Application.ScreenUpdating = True
End Sub

Private Sub txtASub_Change()
Application.ScreenUpdating = False
Call PreviewName
Application.ScreenUpdating = True
End Sub

Private Sub txtAType_Change()
Application.ScreenUpdating = False
Call PreviewName
Application.ScreenUpdating = True
End Sub

Private Sub txtEditFile_Change()

End Sub

Private Sub txtName_Change()
PreviewName
End Sub

Private Sub txtNumber_Change()
PreviewName
End Sub
Private Sub txtPreview_Change()
Worksheets("S").Range("M4").Value = Me.txtPreview.Value
End Sub

Private Sub UserFormS_Initialize()
lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
Worksheets("S").Range("B4:K" & lrB).Name = "IndataS"
ListBox1.RowSource = "IndataS"

End Sub
Private Sub txtSearch_Change()
Call cmdSearch_Click
End Sub

Private Sub txtSearch_Enter()
Call cmdSearch_Click
End Sub

Private Sub UserForm_Activate()
'KHOI DONG DIEN SAN LINK PATH
Me.txtPath.Value = Worksheets("S").Range("AT1").Value
Me.cboUndo.Enabled = False
Me.cboCheck.Enabled = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Application.ScreenUpdating = False
 ''1. Copy du lieu sang khu vuc khac
    Worksheets("S").Activate
    Worksheets("S").Columns("B:K").Select
    Selection.Copy
    Worksheets("S").Columns("AC:AL").PasteSpecial xlPasteValues
 ''2. Xoa truong du lieu copy tu Folder doi ten vao
 lrAO = Worksheets("S").Range("AO" & Rows.Count).End(xlUp).Row
 Worksheets("S").Range("AO4:AO" & lrAO).ClearContents
 Worksheets("S").Range("AQ4") = ""
 Worksheets("S").Range("AS4") = ""
 
     ''Tro ve man hinh chinh
Call FORMAT
Application.ScreenUpdating = True
ThisWorkbook.Sheets("S").Protect Password:="ttdg"

End Sub

Sub FORMAT()
    Worksheets("S").Columns("L:Z").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    Worksheets("S").Columns("AN:AT").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    ''''To mau do neu co ID va Family Name nao bi duplicate
    lrB = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    ActiveSheet.Range("I4:I" & lrB).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    ActiveSheet.Range("K4:K" & lrB).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ''Dinh dang cot ID
    ActiveSheet.Range("K4:K2000").Select
    Selection.NumberFormat = """S""####"
     ActiveWindow.ScrollColumn = 1
End Sub
Sub SaveNew()
lrB = Worksheets("S").Range("B" & Rows.Count).End(xlUp).Row
'' CATEGORY
    Worksheets("S").Activate
    Worksheets("S").Range("C4:D" & lrB).Copy
    Worksheets("S_Source").Activate
    Worksheets("S_Source").Range("B4").PasteSpecial xlPasteValues
    Dim lrBs As Long
    lrBs = Worksheets("S_Source").Range("B" & Rows.Count).End(xlUp).Row - 1
    Worksheets("S_Source").Range("B4:C" & lrBs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
''TYPE
    Worksheets("S").Activate
    Worksheets("S").Range("E4:F" & lrB).Copy
    Worksheets("S_Source").Activate
    Worksheets("S_Source").Range("E4").PasteSpecial xlPasteValues
    Dim lrEs As Long
    lrEs = Worksheets("S_Source").Range("E" & Rows.Count).End(xlUp).Row - 1
    Worksheets("S_Source").Range("E4:F" & lrEs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
''SUBTYPE
    Worksheets("S").Activate
    Worksheets("S").Range("G4:H" & lrB).Copy
    Worksheets("S_Source").Activate
    Worksheets("S_Source").Range("H4").PasteSpecial xlPasteValues
    Dim lrHs As Long
    lrHs = Worksheets("S_Source").Range("H" & Rows.Count).End(xlUp).Row - 1
    Worksheets("S_Source").Range("H4:I" & lrHs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    Worksheets("S").Activate
End Sub

