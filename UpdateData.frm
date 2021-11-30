VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateData 
   Caption         =   "UPDATE DATA"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "UpdateData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UpdateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lrBO As Long
Public lrB As Long
Public lrB1 As Long
Public wb As Workbook
Public wbmain As Workbook
Public i As Long
Public myPath As String

Private Sub cboCancel_Click()
ThisWorkbook.Sheets(Me.lbSheet.Caption).Protect Password:="ttdg"
Unload Me
End Sub

Private Sub cboImport_Click()
Application.ScreenUpdating = False
Dim iRes As Long
iRes = MsgBox("New data will copy and replace old data with this sheet: " & Me.lbSheet.Caption & " ?", vbYesNo + vbQuestion)
Select Case iRes
    Case vbYes
     If Me.lbSheet.Caption = "A" Or Me.lbSheet.Caption = "S" Then
        Dim sh As String
        Set wbmain = ThisWorkbook
        lrB = ThisWorkbook.ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
        ThisWorkbook.ActiveSheet.Range("B4:K" & lrB).ClearContents
        sh = ThisWorkbook.ActiveSheet.Range("BP1").Value
        myPath = Me.txtPath.Text
        Set wb = Workbooks.Open(myPath)
        ''Copy va paste
        lrB1 = wb.Sheets(sh).Range("B" & Rows.Count).End(xlUp).Row
        wb.Sheets(sh).Range("B4:K" & lrB1).Copy ''chu y lrB1 khac lrB
        wbmain.Sheets(Me.lbSheet.Caption).Range("B4").PasteSpecial Paste:=xlPasteFormulas
        wb.Close SaveChanges:=False ''Dong file Import
     Else
        lrB = ThisWorkbook.ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
        ThisWorkbook.ActiveSheet.Range("B4:M" & lrB).ClearContents
        sh = ThisWorkbook.ActiveSheet.Range("BP1").Value
        myPath = Me.txtPath.Text
        Set wb = Workbooks.Open(myPath)
        ''Copy va paste
        lrB1 = wb.Sheets(sh).Range("B" & Rows.Count).End(xlUp).Row
        wb.Sheets(sh).Range("B4:M" & lrB1).Copy ''chu y lrB1 khac lrB
        wbmain.Sheets(Me.lbSheet.Caption).Range("B4").PasteSpecial Paste:=xlPasteFormulas
        wb.Close SaveChanges:=False ''Dong file Import
     End If
    Case vbNo
    Exit Sub
End Select
Application.ScreenUpdating = True
End Sub
Private Sub ListSheet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.cboImport.Enabled = True
Dim i As Long
With Me.ListSheet
    For i = 0 To .ListCount - 1
      If .Selected(i) Then
        ThisWorkbook.ActiveSheet.Range("BP1") = ListSheet.List(i)
        Exit For
      End If
    Next
End With
End Sub
Private Sub txtPath_Change()
Application.ScreenUpdating = False
''**Mo va lay sheet trong file Import
    myPath = Me.txtPath.Text
    Set wbmain = ThisWorkbook
    Set wb = Workbooks.Open(myPath) 'mo file import
    wb.Unprotect Password:="ttdg"
    For i = 1 To Application.Sheets.Count
    wb.Sheets(i).Unprotect Password:="ttdg"
    wbmain.ActiveSheet.Range("BO" & i).Value = wb.Sheets(i).Name
    Next
''*List sheet to Listbox
    lrBO = ThisWorkbook.ActiveSheet.Range("BO" & Rows.Count).End(xlUp).Row
    ListSheet.List = ThisWorkbook.ActiveSheet.Range("BO1:BO" & lrBO + 1).Value
    ThisWorkbook.Sheets(Me.lbSheet.Caption).Activate
    ActiveWindow.SmallScroll Down:=18
Application.ScreenUpdating = True
End Sub
Private Sub UserForm_Activate()
Me.cboImport.Enabled = False
Me.lbSheet.Caption = ActiveWorkbook.ActiveSheet.CodeName
ThisWorkbook.Sheets(Me.lbSheet.Caption).Unprotect Password:="ttdg"
End Sub
Private Sub cboPath_Click()
ThisWorkbook.ActiveSheet.Range("BO:BO").Value = ""
    Dim get_Path As String
    With Application.FileDialog(msoFileDialogFilePicker)
    If .Show <> 0 Then
    get_Path = .SelectedItems(1)
    End If
    Me.txtPath.Text = get_Path
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
ThisWorkbook.Sheets(Me.lbSheet.Caption).Protect Password:="ttdg"
End Sub
