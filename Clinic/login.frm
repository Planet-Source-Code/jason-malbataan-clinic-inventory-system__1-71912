VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00004000&
   Caption         =   "Welcome"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Width           =   2340
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1950
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AMA Computer College Calamba"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   6
      Top             =   960
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00B5742D&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Picture         =   "login.frx":6852
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
'set the global var to false
'to denote a failed login
LoginSucceeded = False
Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim user
user = txtUsername.Text
'check for any record, if there is no record there could be error when looping
If Data1.Recordset.RecordCount = 0 Then
MsgBox ("No recordset")
Exit Sub
End If
'check for correct username and password
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
    If txtUsername = Data1.Recordset.Fields("Username") Then
        If txtPassword = Data1.Recordset.Fields("Password") Then
            LoginSucceeded = True
                Unload Me
            Main.Enabled = True
    Main.StatusBar1.Panels(1).Text = "Current User:  " & UCase(user) & ""
    Main.Show

'MEnu
Main.mnuLogIn.Enabled = False
Main.mnuNew.Enabled = True
Main.mnuPatient.Enabled = True
Main.mnuTools.Enabled = True
Main.mnuInventory.Enabled = True


End If
End If
Data1.Recordset.MoveNext
Loop
If Not LoginSucceeded Then
MsgBox "Invalid UserName/Password, try again!", , "Login"
txtPassword.SetFocus
SendKeys "{Home}+{End}"
txtUsername.Text = ""
txtPassword.Text = ""
txtUsername.SetFocus
End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\login.mdb"
Data1.RecordSource = "login"
Data1.Refresh
End Sub
