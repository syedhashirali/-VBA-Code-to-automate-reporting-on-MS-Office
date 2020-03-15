VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Set2_UserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6516
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   7810
   OleObjectBlob   =   "VBA_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Set2_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim chk As Integer
Dim chk2 As Integer
Dim chk3 As Integer
Dim chk4 As Integer
Dim Opt As Integer
Dim RowCount As Long
Private Sub CheckBox1_Click()
chk = 1
End Sub

Private Sub CheckBox2_Click()
chk4 = 4
End Sub

Private Sub CheckBox3_Click()
chk3 = 3
End Sub

Private Sub CheckBox4_Click()
chk2 = 2
End Sub

Private Sub CommandButton1_Click()
RowCount = Worksheets("Set2").Range("E16").CurrentRegion.Rows.Count
With Worksheets("Set2").Range("E16")

If Opt = 1 Then .Offset(RowCount, 2) = OptionButton1.Caption
If Opt = 2 Then .Offset(RowCount, 2) = OptionButton3.Caption
If Opt = 3 Then .Offset(RowCount, 2) = OptionButton2.Caption
If chk = 1 Then .Offset(RowCount, 4) = "Yes" Else: .Offset(RowCount, 4) = "No"
If chk2 = 2 Then .Offset(RowCount, 5) = "Yes" Else: .Offset(RowCount, 5) = "No"
If chk3 = 3 Then .Offset(RowCount, 6) = "Yes" Else: .Offset(RowCount, 6) = "No"
If chk4 = 4 Then .Offset(RowCount, 7) = "Yes" Else: .Offset(RowCount, 7) = "No"
.Offset(RowCount, 3) = ComboBox1.Value
.Offset(RowCount, 0) = TextBox1.Value
.Offset(RowCount, 1) = TextBox2.Value
End With
End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.Save

End Sub

Private Sub CommandButton3_Click()
ActiveWorkbook.Close False
End Sub

Private Sub OptionButton1_Click()
Opt = 1
End Sub
Private Sub OptionButton2_Click()
Opt = 3
End Sub
Private Sub OptionButton3_Click()
Opt = 2
End Sub
Private Sub TextBox1_Change()

End Sub
Private Sub TextBox2_Change()
End Sub
Private Sub UserForm_Initialize()
ComboBox1.AddItem "Single"
ComboBox1.AddItem "Married"
ComboBox1.AddItem "Widowed"
ComboBox1.AddItem "Divorced"
End Sub



