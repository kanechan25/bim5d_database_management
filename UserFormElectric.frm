VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormE 
   Caption         =   "T&TDG Electrical Database Family Management"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22890
   OleObjectBlob   =   "UserFormElectric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormE"
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
    Worksheets("E").Range("C2") = Me.cboCate.Value
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],R[2]C[-1]:R[1000]C,2,FALSE),"""")"
    Me.txtACate.Value = Worksheets("E").Range("D2")
End Sub

Private Sub cboCheck_Click()
Application.ScreenUpdating = False
''''BB CHO AO (SHEET MEPF)
    Range("BB4").Select
    ActiveCell.Formula2R1C1 = _
        "=IFS(RC[-11]="""","""",LEN(RC[-11])-LEN(SUBSTITUTE(RC[-11],""."",""""))=3,LEFT(RC[-11],vitri(""."",RC[-11],3)-1),LEN(RC[-11])-LEN(SUBSTITUTE(RC[-11],""."",""""))=2,LEFT(RC[-11],vitri(""."",RC[-11],2)-1),TRUE,LEFT(RC[-11],vitri(""."",RC[-11],1)-1))"
    Range("BB4").Select
    Selection.AutoFill Destination:=Range("BB4:BB2000"), Type:=xlFillDefault
''''CHECK LIST
        Dim i As Long
        lrP = ActiveSheet.Range("P" & Rows.Count).End(xlUp).Row
        lrAQ = ActiveSheet.Range("AQ" & Rows.Count).End(xlUp).Row
           i = 4
        Do While i <= lrP
        If ActiveSheet.Range("Y" & i).Value <> ActiveSheet.Range("BB" & i).Value Then
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
Worksheets("E").Range("O4").Copy
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

Private Sub cboSub1_Change()
    Worksheets("E").Range("G2") = Me.cboSub1.Value
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],R[2]C[-1]:R[1000]C,2,FALSE),"""")"
    Me.txtASub1.Value = Worksheets("E").Range("H2")
End Sub

Private Sub cboSub2_Change()
    Worksheets("E").Range("I2") = Me.cboSub2.Value
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],R[2]C[-1]:R[1000]C,2,FALSE),"""")"
    Me.txtASub2.Value = Worksheets("E").Range("J2")
End Sub

Private Sub cboType_Change()
    Worksheets("E").Range("E2") = Me.cboType.Value
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],R[2]C[-1]:R[1000]C,2,FALSE),"""")"
    Me.txtAType.Value = Worksheets("E").Range("F2")
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
   
    OldFile = Worksheets("E").Range("AV5")
    NewFile = Worksheets("E").Range("AV4")
Existed:
    If Worksheets("E").Range("AV4").Value = Worksheets("E").Range("AV5").Value Then
  
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
  ThisWorkbook.Sheets("E").Unprotect Password:="ttdg"
  Worksheets("E").Activate
  UserFormE.Show
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
  lrP = Worksheets("E").Range("P" & Rows.Count).End(xlUp).Row
  i = Worksheets("E").Range("Y4:Y" & lrP).Find(What:=Me.txtName, LookIn:=xlValues).Row - 4
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
lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
Set findvalue = Worksheets("E").Range("K4:K" & lrB).Find(What:=Me.txtName, LookIn:=xlValues)
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
findvalue.Offset(0, -9).Value = ""

Call clearList
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
Me.cboSub1.Value = ""
Me.txtASub1.Value = ""
Me.cboSub2.Value = ""
Me.txtASub2.Value = ""
Me.txtName.Value = ""
Me.txtNumber.Value = ""
Application.ScreenUpdating = True
End Sub
Sub SortData()
'1. SortData
   lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
    Range("A3:M" & lrB).Select
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Add Key:=Range("B4:B" & lrB), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Add Key:=Range("K4:K" & lrB), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("E").Sort
        .SetRange Range("A3:M" & lrB)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  ''''2. Sort Outdata
   lrP = Worksheets("E").Range("P" & Rows.Count).End(xlUp).Row
    Worksheets("E").Range("P3:AA" & lrP).Select
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Add Key:=Range("P4:P" & lrP), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Add Key:=Range("Y4:Y" & lrP), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("E").Sort
        .SetRange Range("P3:AA" & lrP)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Private Sub cmdEdit_Click()
Application.ScreenUpdating = False
If Me.cboFamily.Value = "" Or Me.cboCate.Value = "" Then
Exit Sub
''*xet xem da co name bi trung
  Dim findName As Range
  lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
  Set findName = Worksheets("E").Range("K4:K" & lrB).Find(What:=Me.txtName, LookIn:=xlValues)
  ElseIf Me.txtPreview.Value = Me.txtName.Value Then
  MsgBox "Family Name Existed"
  Exit Sub

Else
Dim findvalue As Range
lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
Set findvalue = Sheets("E").Range("M4:M" & lrB).Find(What:=Me.txtID, LookIn:=xlValues)
''findvalue.value = Me.txtID ''''DANG CHO KHONG SUA ID, VI SUA THONG TIN CUA CHINH ID CU
findvalue.Offset(0, -1).Value = Me.txtNumber.Value
findvalue.Offset(0, -3).Value = Me.txtASub2.Value
findvalue.Offset(0, -4).Value = Me.cboSub2.Value
findvalue.Offset(0, -5).Value = Me.txtASub1.Value
findvalue.Offset(0, -6).Value = Me.cboSub1.Value
findvalue.Offset(0, -7).Value = Me.txtAType.Text
findvalue.Offset(0, -8).Value = Me.cboType.Value
findvalue.Offset(0, -9).Value = Me.txtACate.Value
findvalue.Offset(0, -10).Value = Me.cboCate.Value
findvalue.Offset(0, -11).Value = Me.cboFamily.Value
Call cmdSearch_Click

End If
 Call SortData
 '' Click ve listrow moi thao tac
 Dim i As Long
 lrP = Worksheets("E").Range("P" & Rows.Count).End(xlUp).Row
  On Error Resume Next
 i = Worksheets("E").Range("Y4:Y" & lrP).Find(What:=Me.txtPreview, LookIn:=xlValues).Row - 4
 If err.Number <> 0 Then
  MsgBox "You have not chosen any name to Edit"
   Exit Sub
   End If
 ListBox1.Selected(i) = True
 ''*add new sang sheet source
 Worksheets("E_Source").Activate
 Call SaveNew
 Worksheets("E").Activate
 Me.opbCheck.Value = False
 Application.ScreenUpdating = True
 
End Sub


Private Sub cmdFolder_Click()
Application.ScreenUpdating = False
''**KTRA XEM CO SUBFOLDER BEN TRONG FOLDER (TRONG myPath) KHONG
    Dim myPath As String
    Dim folderCount As Long
    Dim FSO As Object, FolderObject As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ''Trong thu muc "myPath" c� bao nhieu thu muc con?
    myPath = Me.txtPath.Text
    Set FolderObject = FSO.GetFolder(myPath)
    folderCount = FolderObject.SubFolders.Count
    ''MsgBox folderCount
    Set FolderObject = Nothing
    Set FSO = Nothing
If folderCount <> 0 Then
    MsgBox "Subfolder are existing in Folder Path." & Chr(13) & Chr(13) & "Please choose correctly the folder containing this file." & Chr(13) & Chr(13) & "Change Path then Get File List Again"
    Exit Sub
Else
    Dim findvalue As Range
  On Error GoTo Existed
    lrAQ = Worksheets("E").Range("AQ" & Rows.Count).End(xlUp).Row
    Set findvalue = Sheets("E").Range("AQ4:AQ" & lrAQ).Find(What:=Me.txtEditFile, LookIn:=xlValues)
    If Me.txtEditFile.Text = "" Then
      Exit Sub
    Else
      Dim OldFile As String
      Dim NewFile As String
      Worksheets("E").Range("AU4").Value = Me.txtEditFile.Value 'dien ten file moi
      Worksheets("E").Range("AU5").Value = Me.txtEditFile.Value 'dien ten file moi
      OldFile = Worksheets("E").Range("AV2")
      NewFile = Worksheets("E").Range("AV3")
Existed:
    If Worksheets("E").Range("AV2").Value = Worksheets("E").Range("AV3").Value Then
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
    lrAQ = ActiveSheet.Range("AQ" & Rows.Count).End(xlUp).Row
    i = ActiveSheet.Range("AQ4:AQ" & lrAQ).Find(What:=Me.txtEditFile, LookIn:=xlValues).Row - 4
     If err.Number <> 0 Then
      MsgBox "Check family name again"
      Exit Sub
     End If
    ListBox2.Selected(i) = True
''''''
Me.cboUndo.Enabled = True
End If
Me.opbCheck.Value = False
Application.ScreenUpdating = True
End Sub
Sub Reget_file()
Application.ScreenUpdating = False
''Xoa Get list file cu
    Worksheets("E").Range("AQ:AQ").ClearContents
''Lay Get list file moi
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
        End If
    Else
    MsgBox "Folder Path Empty !!" & vbNewLine & vbNewLine & "", vbInformation, "T&T Design Group"
    End If
'*Sort va Filter du lieu trc khi dien
    Dim lrAQ As Long
    lrAQ = Worksheets("E").Range("AQ" & Rows.Count).End(xlUp).Row
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Add Key:=Range("AQ4:AQ" & lrAQ), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("E").Sort
        .SetRange Range("AQ4:AQ" & lrAQ)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ''****KHAC NHAU O ADD2 VA ADD TRONG MS 365 VA MS 2016
    'ActiveWorkbook.Worksheets("E").Sort.SortFields.Add2 Key:=Range("AQ4"), _
        'SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'*Dien du lieu vao Listrow
   ListBox2.List = Worksheets("E").Range("AP4:AQ" & lrAQ).Value
    Me.cboCheck.Enabled = True
   Me.opbCheck.Value = False
Application.ScreenUpdating = True
End Sub
Private Sub cmdGetFile_Click()
Application.ScreenUpdating = False
''Xoa Get list file cu
    Worksheets("E").Range("AQ:AQ").ClearContents
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
        End If
    Else
    MsgBox "Folder Path Empty !!" & vbNewLine & vbNewLine & "", vbInformation, "T&T Design Group"
    End If
'*Sort va Filter du lieu trc khi dien
    Dim lrAQ As Long
    lrAQ = Worksheets("E").Range("AQ" & Rows.Count).End(xlUp).Row
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("E").Sort.SortFields.Add Key:=Range("AQ4:AQ" & lrAQ), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("E").Sort
        .SetRange Range("AQ4:AQ" & lrAQ)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ''****KHAC NHAU O ADD2 VA ADD TRONG MS 365 VA MS 2016
    'ActiveWorkbook.Worksheets("E").Sort.SortFields.Add2 Key:=Range("AQ4"), _
        'SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'*Dien du lieu vao Listrow
   ListBox2.List = Worksheets("E").Range("AP4:AQ" & lrAQ).Value
    Me.cboCheck.Enabled = True
    Me.opbCheck.Value = False
Application.ScreenUpdating = True
End Sub
Sub ListFilesInFolder(ByVal xFolderName As String, ByVal xIsSubfolders As Boolean)
Dim xFileSystemObject As Object
Dim xFolder As Object
Dim xSubFolder As Object
Dim xFile As Object
Dim rowIndex As Long
Set xFileSystemObject = CreateObject("Scripting.FileSystemObject")
Set xFolder = xFileSystemObject.GetFolder(xFolderName)
rowIndex = Application.ActiveSheet.Range("AQ1000").End(xlUp).Row + 3
For Each xFile In xFolder.files
  Application.ActiveSheet.Cells(rowIndex, 43).Value = xFile.Name
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
Worksheets("E").Range("AQ:AQ").ClearContents
Dim get_Path As String
With Application.FileDialog(msoFileDialogFolderPicker)
If .Show <> 0 Then
get_Path = .SelectedItems(1)
End If
Worksheets("E").Cells(1, 48).Value = get_Path
txtPath.Text = get_Path
End With
Me.opbCheck.Value = False
End Sub

Private Sub cmdSearch_Click()
Dim Arr As Variant
Application.ScreenUpdating = False
lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
    Arr = Range("B3:M" & lrB).Value
    
Me.opbCheck.Value = False
If cboSelect.Value = "Family Name" Then
  If Len(txtSearch.Value) = 0 Then
       ActiveSheet.AutoFilterMode = False
  Else
    ''If ActiveSheet.AutoFilterMode = True Then
       ''ActiveSheet.AutoFilterMode = False
    ''End If
    ActiveSheet.Range("B3:M" & lrB).AutoFilter Field:=8, Criteria1:="*" & txtSearch.Value & "*", CopyToRange:=Range("E!Extract"), Unique:=False
  End If
''tim dong cuoi cua du lieu ket qua de dien vao listrow
        lrP = Worksheets("E").Range("P" & Rows.Count).End(xlUp).Row
        ListBox1.RowSource = Worksheets("E").Range("P4:AA" & lrP).Address
        Me.txtMax.Value = Worksheets("E").Range("N1")
        Me.txtNumbers.Value = Worksheets("E").Range("N2")
Else
    On Error GoTo errHandler:
        Worksheets("E").Range("O2") = Me.cboSelect.Value
        Worksheets("E").Range("O3") = Me.txtSearch.Value
    Worksheets("E").Range("B3:M" & lrB).AdvancedFilter Action:=xlFilterCopy, _
    CriteriaRange:=Range("E!Criteria"), _
    CopyToRange:=Range("E!Extract"), Unique:=False
 ''tim dong cuoi cua du lieu ket qua de dien vao listrow
        lrP = Worksheets("E").Range("P" & Rows.Count).End(xlUp).Row
        ListBox1.RowSource = Worksheets("E").Range("P4:AA" & lrP).Address
        Me.txtMax.Value = Worksheets("E").Range("N1")
        Me.txtNumbers.Value = Worksheets("E").Range("N2")
    Exit Sub
errHandler:
        MsgBox "No Sort Field OR Sort Field and Search do not match"
End If
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
' BAO LOI KHI NHAP THONG TIN DUPLICATE
Dim findName As Range
lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
Set findName = Worksheets("E").Range("K4:K" & lrB).Find(What:=Me.txtName, LookIn:=xlValues)
 On Error Resume Next
If Me.txtName.Value = Me.txtPreview.Value Or findName.Value = Me.txtPreview.Value Then
MsgBox "Family Name Existed"
Exit Sub
Else
Dim Drng As Range
Set Drng = Worksheets("E").Range("B4")

Drng.End(xlDown).Offset(1, 0).Value = Me.cboFamily.Value
Drng.End(xlDown).Offset(0, 1).Value = Me.cboCate.Value
Drng.End(xlDown).Offset(0, 2).Value = Me.txtACate.Value
Drng.End(xlDown).Offset(0, 3).Value = Me.cboType.Value
Drng.End(xlDown).Offset(0, 4).Value = Me.txtAType.Value
Drng.End(xlDown).Offset(0, 5).Value = Me.cboSub1.Value
Drng.End(xlDown).Offset(0, 6).Value = Me.txtASub1.Value
Drng.End(xlDown).Offset(0, 7).Value = Me.cboSub2.Value
Drng.End(xlDown).Offset(0, 8).Value = Me.txtASub2.Value
Drng.End(xlDown).Offset(0, 10).Value = Me.txtNumber.Value
Drng.End(xlDown).Offset(0, 11).Value = Worksheets("E").Range("N3").Value + 1

' 2.COPY CONG THUC CUA O BEN TREN
Drng.End(xlDown).Offset(-1, 9).Copy
Drng.End(xlDown).Offset(0, 9).PasteSpecial Paste:=xlPasteFormulas
End If

   Call SortData
   Call cmdSearch_Click

''3. Click ve listrow moi thao tac
 Dim i As Long
 lrP = Worksheets("E").Range("P" & Rows.Count).End(xlUp).Row
 i = Worksheets("E").Range("Y4:Y" & lrP).Find(What:=Worksheets("E").Range("O4"), LookIn:=xlValues).Row - 4
 ListBox1.Selected(i) = True
 ''DIEN DU LIEU O DONG VUA MOI ADDNEW VAO
 ''Dien thong tin vao cac box
    Me.cboFamily.Value = Me.ListBox1.Value
    Me.cboCate.Value = Me.ListBox1.Column(1)
    Me.txtACate.Value = Me.ListBox1.Column(2)
    Me.cboType.Value = Me.ListBox1.Column(3)
    Me.txtAType.Value = Me.ListBox1.Column(4)
    Me.cboSub1.Value = Me.ListBox1.Column(5)
    Me.txtASub1.Value = Me.ListBox1.Column(6)
    Me.cboSub2.Value = Me.ListBox1.Column(7)
    Me.txtASub2.Value = Me.ListBox1.Column(8)
    Me.txtName.Value = Me.ListBox1.Column(9)
    Me.txtNumber.Value = Me.ListBox1.Column(10)
    Me.txtID.Value = Me.ListBox1.Column(11)
 End If
 Call SaveNew
 Me.cmdEdit.Enabled = False ''sau khi add khoa edit vi muon edit phai chon name
 Me.opbCheck.Value = False
Application.ScreenUpdating = True
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
On Error Resume Next
''Dien thong tin vao cac box
Me.cboFamily.Value = Me.ListBox1.Value
Me.cboCate.Value = Me.ListBox1.Column(1)
Me.txtACate.Value = Me.ListBox1.Column(2)
Me.cboType.Value = Me.ListBox1.Column(3)
Me.txtAType.Value = Me.ListBox1.Column(4)
Me.cboSub1.Value = Me.ListBox1.Column(5)
Me.txtASub1.Value = Me.ListBox1.Column(6)
Me.cboSub2.Value = Me.ListBox1.Column(7)
Me.txtASub2.Value = Me.ListBox1.Column(8)
Me.txtName.Value = Me.ListBox1.Column(9)
Me.txtNumber.Value = Me.ListBox1.Column(10)
Me.txtID.Value = Me.ListBox1.Column(11)
On Error GoTo 0
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
Me.cboUndo.Enabled = False
''Dien thong tin vao txtbox
Dim i As Long
With Me.ListBox2
    For i = 0 To .ListCount - 1
      If .Selected(i) Then
        Me.txtEditFile.Value = ListBox2.List(i, 1)
        Worksheets("E").Range("AS4") = ListBox2.List(i, 1)
        Worksheets("E").Range("AS5") = ListBox2.List(i, 1)
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
= "" And Me.txtAType.Value = "" And Me.cboSub1.Value = "" And Me.txtASub1.Value = "" And Me.cboSub2.Value = "" And Me.txtASub2.Value = "" Then
  Me.txtPreview.Value = ""
''A.Xet TH Number <10
ElseIf Me.txtNumber.Value < 10 Then
  ''1.TH txtASub2 va txtAsub1 bang - hoac bang ""
   If Me.txtASub2.Value = "-" Or Me.txtASub2.Value = "" Then
      If Me.txtASub1.Value = "-" Or Me.txtASub1.Value = "" Then
      Me.txtPreview.Value = "E_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_0" & Me.txtNumber.Value
      Else
      Me.txtPreview.Value = "E_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtASub1.Value & "_0" & Me.txtNumber.Value
      End If
   Else
   Me.txtPreview.Value = "E_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtASub1.Value & "_" & Me.txtASub2.Value & "_0" & Me.txtNumber.Value
   End If
Else
''B.TH con lai Number>10
   If Me.txtASub2.Value = "-" Or Me.txtASub2.Value = "" Then
      If Me.txtASub1.Value = "-" Or Me.txtASub1.Value = "" Then
      Me.txtPreview.Value = "E_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtNumber.Value
      Else
      Me.txtPreview.Value = "E_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtASub1.Value & "_" & Me.txtNumber.Value
      End If
   Else
   Me.txtPreview.Value = "E_" & Me.txtACate.Value & "_" & Me.txtAType.Value & "_" & Me.txtASub1.Value & "_" & Me.txtASub2.Value & "_" & Me.txtNumber.Value
   End If
End If
End Sub


Private Sub txtACate_Change()
Application.ScreenUpdating = False
Call PreviewName
Application.ScreenUpdating = True
End Sub

Private Sub txtASub1_Change()
Application.ScreenUpdating = False
Call PreviewName
Application.ScreenUpdating = True
End Sub

Private Sub txtASub2_Change()
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
Call PreviewName
End Sub

Private Sub txtNumber_Change()
Call PreviewName
End Sub
Private Sub txtPreview_Change()
Worksheets("E").Range("O4").Value = Me.txtPreview.Value
End Sub

Private Sub UserFormE_Initialize()
Worksheets("E").Activate
lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
Worksheets("E").Range("B4:M" & lrB).Name = "IndataE"
ListBox1.RowSource = "IndataE"

End Sub
Private Sub txtSearch_Change()
Call cmdSearch_Click
End Sub

Private Sub txtSearch_Enter()
Call cmdSearch_Click
End Sub

Private Sub UserForm_Activate()
'KHOI DONG DIEN SAN LINK PATH
Me.txtPath.Value = Worksheets("E").Range("AV1").Value
Me.cboUndo.Enabled = False
Me.cboCheck.Enabled = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 ''1. Copy du lieu sang khu vuc khac
 Application.ScreenUpdating = False
    Worksheets("E").Activate
    Worksheets("E").Columns("B:M").Select
    Selection.Copy
    Worksheets("E").Columns("AC:AN").PasteSpecial xlPasteValues
 ''2. Xoa truong du lieu copy tu Folder doi ten vao
lrAQ = Worksheets("E").Range("AQ" & Rows.Count).End(xlUp).Row
 Worksheets("E").Range("AQ4:AQ" & lrAQ).ClearContents
 Worksheets("E").Range("AS4") = ""
 Worksheets("E").Range("AU4") = ""
     ''Tro ve man hinh chinh
Call FORMAT
Application.ScreenUpdating = True
ThisWorkbook.Sheets("E").Protect Password:="ttdg"
End Sub

Sub FORMAT()
    Worksheets("E").Columns("N:AA").Select
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
     Worksheets("E").Columns("AP:AV").Select
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
    ActiveSheet.Range("M4:M" & lrB).Select
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
    ActiveSheet.Range("M4:M2000").Select
    Selection.NumberFormat = """E""####"
     ActiveWindow.ScrollColumn = 1
End Sub

Sub SaveNew()
lrB = Worksheets("E").Range("B" & Rows.Count).End(xlUp).Row
'' CATEGORY
    Worksheets("E").Activate
    Worksheets("E").Range("C4:D" & lrB).Copy
    Worksheets("E_Source").Activate
    Worksheets("E_Source").Range("B4").PasteSpecial xlPasteValues
    Dim lrBs As Long
    lrBs = Worksheets("E_Source").Range("B" & Rows.Count).End(xlUp).Row - 1
    Worksheets("E_Source").Range("B4:C" & lrBs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
''TYPE
    Worksheets("E").Activate
    Worksheets("E").Range("E4:F" & lrB).Copy
    Worksheets("E_Source").Activate
    Worksheets("E_Source").Range("E4").PasteSpecial xlPasteValues
    Dim lrEs As Long
    lrEs = Worksheets("E_Source").Range("E" & Rows.Count).End(xlUp).Row - 1
    Worksheets("E_Source").Range("E4:F" & lrEs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
''SUBTYPE 1
    Worksheets("E").Activate
    Worksheets("E").Range("G4:H" & lrB).Copy
    Worksheets("E_Source").Activate
    Worksheets("E_Source").Range("H4").PasteSpecial xlPasteValues
    Dim lrHs As Long
    lrHs = Worksheets("E_Source").Range("H" & Rows.Count).End(xlUp).Row - 1
    Worksheets("E_Source").Range("H4:I" & lrHs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes

'''SUBTYPE 2
    Worksheets("E").Activate
    Worksheets("E").Range("I4:J" & lrB).Copy
    Worksheets("E_Source").Activate
    Worksheets("E_Source").Range("K4").PasteSpecial xlPasteValues
    Dim lrKs As Long
    lrKs = Worksheets("E_Source").Range("K" & Rows.Count).End(xlUp).Row - 1
    Worksheets("E_Source").Range("K4:L" & lrHs).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    Worksheets("E").Activate
End Sub



