VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   Caption         =   "Print"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PrintSI.frx":0000
   LinkTopic       =   "Form6"
   Moveable        =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Select Month and Year:"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   74252289
         CurrentDate     =   39763
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Preview"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "This is the path of the selected file"
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Text            =   "Text6"
      Top             =   7440
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
      Left            =   6600
      TabIndex        =   1
      Text            =   "time"
      Top             =   7440
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
      Left            =   7560
      TabIndex        =   0
      Text            =   "Date"
      Top             =   7440
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
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
DataEnvironment1.rsCommand1.Source = "Select StudentID From history where StudentID = '" & Text1.Text & "'"
DataEnvironment1.rsCommand1.Open
DataReport1.Show
Exit Sub
Err:
DataEnvironment1.rsCommand1.Close
DataEnvironment1.rsCommand1.Source = "Select StudentID  From history where StudentID = '" & Text1.Text & "'"
End Sub

Private Sub Form_Load()
'txtYear.Text = Format(Date, "yyyy")
Me.Top = 200
Me.Left = 200
End Sub

Private Sub VScroll1_Change()
txtYear.Text = VScrollBar
End Sub
