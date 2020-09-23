VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   9450
   ClientLeft      =   600
   ClientTop       =   795
   ClientWidth     =   14610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Studlog.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Studlog.frx":6852
   ScaleHeight     =   9450
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private refhis As Integer
'###############################################################'
'#                   AMA Computer College                      #'
'#   F.P Perez Bldg. National Highway Parian, Calamba City     #'
'#                                                             #'
'#                       09204128374                           #'
'#              Ang Mag nakaw ng Code Panget                   #'
'###############################################################'

Private Sub cmdClose_Click()
Unload Me
Form1.Enabled = True
End Sub
Private Sub form_activate()
'txtsearch.SetFocus
End Sub
Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\Database\Entrance.mdb"
'Data1.RecordSource = "Entrance"
'Data1.Connect = ";pwd=G1tqJy2N"
'Data1.Refresh

'Data2.DatabaseName = App.Path & "\Database\Entrance.mdb"
'Data2.RecordSource = "history"
'Data2.Refresh

'Para sa reference number
'Data2.Recordset.MoveLast
'refhis = Data2.Recordset.Fields("Reference") + 1
'txtref.Text = refhis



'
End Sub
Private Sub Command1_Click()
On Error Resume Next
Data1.RecordSource = "Select* from Entrance where FirstName = '" & txtsearch.Text & "' OR LastName = '" & txtsearch.Text & "' OR StudentID = '" & txtsearch.Text & "'"
Data1.Refresh
Text12.Text = Form1.ProperCase(txtsearch.Text)
With Data1.Recordset
If txtsearch.Text Then
txtsearch.Text = Text12.Text
End If
If Not txtsearch.Text = .Fields("FirstName") Xor txtsearch.Text = .Fields("LastName") Xor txtsearch.Text = .Fields("StudentID") Then
MsgBox "No Record Found"
txtsearch.Text = ""
txtsearch.SetFocus
Else
txtPath.Text = .Fields("Picture")
        Label1.Caption = Format(.Fields("StudentID"))
            txtStudentID.Text = .Fields("StudentID")
                Label2.Caption = Format(.Fields("FirstName"))
                    txtFirstName.Text = .Fields("FirstName")
                Label3.Caption = Format(.Fields("LastName"))
            txtLastName.Text = .Fields("LastName")
        Label4.Caption = Format(.Fields("Course"))
txtStatus.Text = .Fields("Status")
'txtTimein.Text = .Fields("Clear")
imgCurrent.Picture = LoadPicture(txtPath.Text)
Timer2.Enabled = True
Frame1.Visible = False
End If

'Function to
If txtStatus.Text = "" Then
    Call IsulatSaLogin
        ElseIf txtStatus.Text = "IN" Then
            Call IsulatSaLogout
        Call KopyahinSaHistory
    Call burahin
End If
End With
End Sub
Private Sub IsulatSaLogin()
With Data1.Recordset
.Edit
.Fields("Time_in") = Label5.Caption
.Fields("Date_in") = Date
.Fields("Status") = "IN"
.Fields("Clear") = "1"
.Update
Text1.Text = "Logged IN"
Label9.Caption = Format(Text1.Text)
End With
End Sub
Private Sub IsulatSaLogout()
With Data1.Recordset
.Edit
.Fields("Time_out") = Label5.Caption
.Fields("Date_out") = Date
.Update
Text1.Text = "Logged OUT"
Label9.Caption = Format(Text1.Text)
End With
End Sub
Private Sub KopyahinSaHistory()
With Data2.Recordset
.AddNew
.Fields("Reference") = txtref.Text
.Fields("StudentID") = txtStudentID.Text
.Fields("FirstName") = txtFirstName.Text
.Fields("LastName") = txtLastName.Text
.Fields("Course") = Data1.Recordset.Fields("Course")
.Fields("Time_in") = Data1.Recordset.Fields("Time_in")
.Fields("Date_in") = Data1.Recordset.Fields("Date_in")
.Fields("Date_out") = Data1.Recordset.Fields("Date_out")
.Fields("Time_out") = Data1.Recordset.Fields("Time_out")
.Update
End With
End Sub
Private Sub burahin()
With Data1.Recordset
.Edit
Data1.Recordset.Fields("Time_in") = txtTimein.Text
Data1.Recordset.Fields("Date_in") = txtTimein.Text
Data1.Recordset.Fields("Time_out") = txtTimein.Text
Data1.Recordset.Fields("Date_out") = txtTimein.Text
Data1.Recordset.Fields("Status") = txtTimein.Text
.Update
End With
End Sub

Private Sub Timer1_Timer()
   Label5.Caption = Time
End Sub

Private Sub Timer2_Timer()
Static iValue As Double
    Dim i As Long
    iValue = iValue + 1
    If iValue > 100 Then
Else
    'ProgressBar1.Value = iValue
    'Label6.Caption = iValue
    Exit Sub
  End If
    'If Timer2.Enabled = False Then
    'Else
    Unload Me
    Form3.Show
    'Form3.Timer2.Enabled = False
    Form3.Left = 300
    Form3.Top = 50
    'End If
End Sub
Private Sub txtsearch_Change()
'Data1.RecordSource = "Select * from brgy where"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'If MsgBox("Are you sure you want to quit the application?", vbQuestion + vbYesNo) = vbNo Then Cancel = True
Main.Enabled = True
End Sub

