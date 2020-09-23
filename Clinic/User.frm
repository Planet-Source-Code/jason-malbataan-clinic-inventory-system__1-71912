VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00004000&
   Caption         =   "Add User"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "User.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   2670
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "User.frx":6852
      Height          =   1575
      Left            =   0
      OleObjectBlob   =   "User.frx":6866
      TabIndex        =   7
      Top             =   4560
      Width           =   4695
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "Change Username/Password"
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtConfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   885
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
On Error Resume Next
Data1.RecordSource = "Select * from login where Password = '" & Text2.Text & "'"
Data1.Refresh
With Data1.Recordset
If Not Text2.Text = Data1.Recordset.Fields("Password") Then
    MsgBox ("Wrong User/Password")
        ElseIf Not Text2.Text = txtConfirm.Text Then
            MsgBox ("Password Mismatch")
            
Data1.Recordset.Delete
Data1.Recordset.MoveNext
MsgBox ("Successfully Changed")
End If
End With
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
With Data1.Recordset
If Len(Text2.Text) < 6 Then
    MsgBox ("Enter Maximum of six Characters")
ElseIf Not txtConfirm.Text = Text2.Text Then
    MsgBox ("Password not Match")
    txtConfirm.Text = ""
    Text2.Text = ""
    Else
    Data1.Recordset.AddNew
    Data1.Recordset.Fields("Username") = Text1.Text
    Data1.Recordset.Fields("Password") = Text2.Text
    Data1.Recordset.Update
    'Login.MemID.Caption = ""
    'Login.Pass.Caption = ""
    Text1.Text = ""
    Text2.Text = ""
    txtConfirm.Text = ""

    End If
    End With
    Data1.Refresh
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Data1.Recordset.Delete
Data1.Recordset.MoveNext
MsgBox ("Record Deleted")

If Data1.Recordset.RecordCount = 0 Then
MsgBox "there are no more records."
Else
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MoveLast
End If
End If


'datPass.Recordset.Delete
'datPass.Recordset.MoveNext
'MsgBox ("Was Delete")
Data1.Refresh



End Sub



Private Sub form_load()
Data1.DatabaseName = App.Path & "\login.mdb"
Data1.RecordSource = "login"
Data1.Refresh
End Sub

Private Sub MSFlexGrid1_Click()
Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1)
Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 2)
End Sub
