VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medication Information"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   495
      Left            =   8400
      TabIndex        =   43
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Patient Historical Data"
      Height          =   2655
      Left            =   5040
      TabIndex        =   38
      Top             =   4920
      Width           =   4815
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtDiagnostic 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Prescription"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2400
         TabIndex        =   42
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Diagnostics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Family History"
      Height          =   2655
      Left            =   120
      TabIndex        =   33
      Top             =   4920
      Width           =   4815
      Begin VB.TextBox txtAllergy 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtMedHist 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Allergy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2520
         TabIndex        =   37
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Family History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Information"
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   9735
      Begin VB.TextBox txtLastName 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   2640
         TabIndex        =   21
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtFirstName 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   4920
         TabIndex        =   18
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtCourse 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   2640
         TabIndex        =   17
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtBirthdate 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtSex 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   4920
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtAdd1 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   1035
         Left            =   315
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2100
         Width           =   1935
      End
      Begin VB.TextBox txtDesignation 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   6960
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   6960
         TabIndex        =   11
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   300
         TabIndex        =   32
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2640
         TabIndex        =   31
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Student/Employee ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   4920
         TabIndex        =   28
         Top             =   2760
         Width           =   435
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   27
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Birth Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   0
         Left            =   2670
         TabIndex        =   26
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   2
         Left            =   4905
         TabIndex        =   25
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   24
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   6960
         TabIndex        =   23
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Designation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   6960
         TabIndex        =   22
         Top             =   1200
         Width           =   1080
      End
   End
   Begin VB.TextBox txtsearch 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      Picture         =   "Update.frx":6852
      TabIndex        =   6
      ToolTipText     =   "Name or Lastname"
      Top             =   360
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   53411841
      CurrentDate     =   38965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      ToolTipText     =   "This is the path of the selected file"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Text            =   "Text6"
      Top             =   9600
      Width           =   735
   End
   Begin VB.TextBox txtTime 
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
      Left            =   5760
      TabIndex        =   1
      Text            =   "time"
      Top             =   9600
      Width           =   855
   End
   Begin VB.TextBox txtDate 
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
      Left            =   6720
      TabIndex        =   0
      Text            =   "Date"
      Top             =   9600
      Width           =   975
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
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:"
      ForeColor       =   &H00004000&
      Height          =   210
      Left            =   6480
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      ForeColor       =   &H00004000&
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################'
'#                                  #'
'#    ANG GUMAYA NG CODE MADAYA     #'
'#                                  #'
'####################################'
Private Sub Command1_Click()
On Error Resume Next
On Error GoTo Err
DataEnvironment1.rsCommand1.Source = "Select FirstName, LastName, StudentID, Time_out, Date_out, Time_in, Date_in from history where format(Date_in)  = '" & DTPicker1 & "'"
DataEnvironment1.rsCommand1.Open
DataReport1.Show
Unload Me
Exit Sub
Err:
DataEnvironment1.rsCommand1.Close
DataEnvironment1.rsCommand1.Source = "Select FirstName, LastName, StudentID, Time_out, Date_out, Time_in, Date_in from history where format(Date_in)  = '" & DTPicker1 & "'"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.RecordSource = "Select * from Medication where ID = '" & txtsearch.Text & "'"
Data1.Refresh
With Data1.Recordset
If Not txtsearch.Text = .Fields("ID") Then
MsgBox ("No Record Found")
Else
Text1.Text = .Fields("ID")
txtFirstName.Text = .Fields("FirstName")
txtLastName.Text = .Fields("LastName")
txtAdd1.Text = .Fields("Address1")
txtStatus.Text = .Fields("Status")
txtDesignation.Text = .Fields("Designation")
txtSex.Text = .Fields("Sex")
txtBirthDate.Text = .Fields("BirthDate")
txtDept.Text = .Fields("Department")
txtCourse.Text = .Fields("Course")
txtAge.Text = .Fields("Age")
txtMedHist.Text = .Fields("MedicalHist")
txtAllergy.Text = .Fields("Allergy")
txtDiagnostic.Text = .Fields("Diagnostics")
Text2.Text = .Fields("Prescription")
End If
End With
End Sub

Private Sub Command3_Click()
Dim First, Last
With Form8
First = txtFirstName.Text
Last = txtLastName.Text

.Show
.Label2 = Text1.Text
.Label5.Caption = "" & (First) & "  " & (Last) & ""
.Label14 = txtCourse.Text
.Label16 = txtStatus.Text
.Label17 = txtSex.Text
.Label18 = txtAge.Text
.Label19 = txtBirthDate.Text
.Label20 = txtDept.Text
.Label21 = txtAdd1.Text
.Label23 = txtMedHist.Text
.Label24 = txtAllergy.Text
.Label27 = txtDiagnostic.Text
.Label28 = Text2.Text
'.PrintForm
End With
End Sub

Private Sub form_activate()
On Error Resume Next
txtsearch.SetFocus
End Sub

Private Sub form_load()
Data1.DatabaseName = App.Path & "\Database\Database.mdb"
Data1.RecordSource = "Medication"
Data1.Connect = ";pwd=nujAwlJa"
Data1.Refresh
End Sub
