VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Medication"
   ClientHeight    =   6855
   ClientLeft      =   105
   ClientTop       =   285
   ClientWidth     =   9870
   ControlBox      =   0   'False
   FillColor       =   &H00004000&
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Width           =   1380
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2295
      Left            =   1080
      TabIndex        =   50
      Top             =   7200
      Width           =   8655
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   7440
         TabIndex        =   66
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   5880
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   6120
         TabIndex        =   63
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   1680
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtdate 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   6120
         TabIndex        =   55
         Top             =   1650
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   1680
         TabIndex        =   54
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   765
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtGen 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   1680
         TabIndex        =   52
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   6120
         TabIndex        =   51
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   960
         TabIndex        =   62
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   60
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generic Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   59
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5040
         TabIndex        =   58
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4920
         TabIndex        =   57
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   56
         Top             =   960
         Width           =   1410
      End
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Width           =   1260
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Width           =   1065
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Width           =   1140
   End
   Begin VB.TextBox txtref 
      Height          =   285
      Left            =   2280
      TabIndex        =   37
      Top             =   10440
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   36
      ToolTipText     =   "This is the path of the selected file"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Width           =   1110
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6360
      Width           =   4575
   End
   Begin VB.CommandButton cmdproceed 
      Caption         =   "Proceed"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search by: ID Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         Picture         =   "Medication.frx":0000
         TabIndex        =   7
         ToolTipText     =   "Name or Lastname"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtsearch 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16384
      ForeColor       =   16384
      TabCaption(0)   =   "Personl Info."
      TabPicture(0)   =   "Medication.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgCurrent"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Image1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAdd3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAdd2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAdd1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtLastName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFirstName"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtID"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtNoofVisit"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtBirthDate"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtStatus"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCourse"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtSex"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtAge"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDesignation"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDepartment"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Patient Historical Data"
      TabPicture(1)   =   "Medication.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Medication"
      TabPicture(2)   =   "Medication.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid2"
      Tab(2).Control(1)=   "txtEx"
      Tab(2).Control(2)=   "txtBra"
      Tab(2).Control(3)=   "txtDes"
      Tab(2).Control(4)=   "txtQuan"
      Tab(2).Control(5)=   "MSFlexGrid1"
      Tab(2).Control(6)=   "cmdLeft"
      Tab(2).Control(7)=   "cmdRight"
      Tab(2).Control(8)=   "List1"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label2"
      Tab(2).Control(11)=   "Label1"
      Tab(2).Control(12)=   "Label9"
      Tab(2).ControlCount=   13
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2295
         Left            =   -72720
         TabIndex        =   65
         Top             =   600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   5
         Cols            =   7
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtEx 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   -71040
         TabIndex        =   49
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtBra 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   -71040
         TabIndex        =   44
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtDes 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   765
         Left            =   -71040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox txtQuan 
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   -66840
         TabIndex        =   42
         Top             =   4560
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "Medication.frx":68A6
         Height          =   2295
         Left            =   -72720
         TabIndex        =   41
         Top             =   600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4048
         _Version        =   393216
         FixedCols       =   0
         ForeColor       =   16384
         Enabled         =   0   'False
         ScrollBars      =   0
         MergeCells      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73680
         TabIndex        =   39
         Top             =   4320
         Width           =   855
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   3630
         Left            =   -74760
         TabIndex        =   38
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtDepartment 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   33
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtDesignation 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         TabIndex        =   23
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtSex 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         TabIndex        =   22
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtCourse 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   21
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtBirthDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtNoofVisit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtFirstName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtLastName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   11
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAdd1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   1920
         Width           =   6255
      End
      Begin VB.TextBox txtAdd2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   2280
         Width           =   6255
      End
      Begin VB.TextBox txtAdd3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   2640
         Width           =   6255
      End
      Begin VB.Frame Frame1 
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   4215
         Begin VB.TextBox txtDiagnostic 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2895
         Left            =   -70080
         TabIndex        =   1
         Top             =   1200
         Width           =   4335
         Begin VB.TextBox txtprescription 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   -72480
         TabIndex        =   48
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   -67920
         TabIndex        =   47
         Top             =   4560
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   -72360
         TabIndex        =   46
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   -72600
         TabIndex        =   45
         Top             =   3600
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   1080
         Picture         =   "Medication.frx":68BA
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Designation:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2880
         TabIndex        =   32
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   6240
         TabIndex        =   31
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Birth Date:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   2880
         TabIndex        =   29
         Top             =   3720
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   7200
         TabIndex        =   28
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   5040
         TabIndex        =   27
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   7200
         TabIndex        =   26
         Top             =   3720
         Width           =   435
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "No. of Visit:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   25
         Top             =   3120
         Width           =   1125
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   5040
         TabIndex        =   24
         Top             =   3720
         Width           =   765
      End
      Begin VB.Image imgCurrent 
         BorderStyle     =   1  'Fixed Single
         Height          =   2415
         Left            =   240
         Stretch         =   -1  'True
         ToolTipText     =   "Picture preview"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Student/Employee ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2820
         TabIndex        =   17
         Top             =   480
         Width           =   2205
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   6120
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   1680
         Width           =   885
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ref As Integer
Function ProperCase(ByVal txt As String) As String
Dim utot As Long
Dim need_cap As Boolean
Dim ipot As Integer
Dim charing As String

    txt = LCase(txt)
    utot = Len(txt)
    need_cap = True
    For ipot = 1 To utot
        charing = Mid$(txt, ipot, 1)
        If charing >= "a" And charing <= "z" Then
            If need_cap Then
                Mid$(txt, ipot, 1) = UCase$(charing)
                need_cap = False
            End If
        Else
            need_cap = True
        End If
    Next ipot
    ProperCase = txt
End Function

Private Sub cmdLeft_Click()
Dim i As Integer
i = List1.ListIndex
If i < List1.ListCount - 1 Then
i = i + 1
List1.ListIndex = i
'Label1 = List1.List(i)
End If
End Sub
Private Sub burahin()
Data5.RecordSource = "Select * from Stockin where Brand_Name = '" & Text5.Text & "'"
Data5.Refresh
With Data5.Recordset
Data5.Recordset.Delete
Data5.Recordset.MoveNext
End With
End Sub
Private Sub quant()
Data5.RecordSource = "Select * from Stockin where Brand_Name = '" & Text5.Text & "'"
Data5.Refresh
With Data5.Recordset
.Edit
.Fields("Quantity") = Text4.Text
.Update
End With
End Sub
Private Sub Medication()
With Data2.Recordset
.AddNew
.Fields("MedRef") = txtref.Text
.Fields("ID") = txtsearch.Text
.Fields("FirstName") = ProperCase(txtFirstName.Text)
.Fields("LastName") = ProperCase(txtLastName.Text)
.Fields("Address1") = ProperCase(txtAdd1.Text)
.Fields("Address2") = ProperCase(txtAdd2.Text)
.Fields("Address3") = ProperCase(txtAdd3.Text)
.Fields("Picture") = txtPath.Text
.Fields("NoofVisit") = ProperCase(txtNoofVisit.Text)
.Fields("BirthDate") = txtBirthDate.Text
.Fields("Age") = txtAge.Text
.Fields("Status") = ProperCase(txtStatus.Text)
.Fields("Sex") = ProperCase(txtSex.Text)
.Fields("Course") = ProperCase(txtCourse.Text)
.Fields("Designation") = ProperCase(txtDesignation.Text)
.Fields("Department") = ProperCase(txtDepartment.Text)
.Fields("MedicalHist") = Data1.Recordset.Fields("MedicalHist")
.Fields("Allergy") = Data1.Recordset.Fields("Allergy")
.Fields("Diagnostics") = ProperCase(txtDiagnostic.Text)
.Fields("Prescription") = ProperCase(txtprescription.Text)
.Update
End With
End Sub
Private Sub Stockout()
With Data6.Recordset
.AddNew
.Fields("Ref_Med") = txtCode.Text
.Fields("Generic_Name") = txtGen.Text
.Fields("Brand_Name") = Text3.Text
.Fields("Description") = Text2.Text
.Fields("Expiration") = txtdate.Text
.Fields("Date_Out") = Date
.Fields("Quantity") = "0"
.Update
End With
End Sub
Private Sub cmdproceed_Click()
Data5.RecordSource = "Select * from Stockin where Brand_Name = '" & Text5.Text & "'"
Data5.Refresh
With Data5.Recordset
If Text4.Text = "" Then
Call Medication
MsgBox "Proceeded Successfully"
ElseIf Text4.Text <= 0 Then
Call Stockout
Call Medication
MsgBox "Proceeded Successfully"
Data5.Recordset.Delete
Data5.Recordset.MoveNext
ElseIf Text4.Text >= 0 Then
Call Medication
Call quant
MsgBox "Proceeded Successfully"
End If
End With
End Sub

Private Sub cmdRight_Click()
Dim i As Integer
i = List1.ListIndex
If i > 0 Then
i = i - 1
List1.ListIndex = i
'Label1 = List1.List(i)
End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Data1.RecordSource = "Select * from Information where ID = '" & txtsearch.Text & "'"
Data1.Refresh
With Data1.Recordset
If Not txtsearch.Text = .Fields("ID") Then
MsgBox ("No Record Found")
Else
cmdproceed.Enabled = True
txtID.Text = .Fields("ID")
txtFirstName.Text = .Fields("FirstName")
txtLastName.Text = .Fields("LastName")
txtAdd1.Text = .Fields("Address1")
txtAdd2.Text = .Fields("Address2")
txtAdd3.Text = .Fields("Address3")
txtPath.Text = .Fields("Picture")
Form9.imgCurrent.Picture = LoadPicture(txtPath.Text)
txtNoofVisit.Text = .Fields("NoofVisit")
txtStatus.Text = .Fields("Status")
txtAge.Text = .Fields("Age")
txtBirthDate.Text = .Fields("BirthDate")
txtCourse.Text = .Fields("Course")
txtSex.Text = .Fields("Sex")
txtDesignation.Text = .Fields("Designation")
txtDepartment.Text = .Fields("Department")
List1.Enabled = True
txtEx.Enabled = True
txtBra.Enabled = True
txtDes.Enabled = True
txtQuan.Enabled = True
MSFlexGrid1.Enabled = True
End If
End With
End Sub

Private Sub Command1_Click()
Call Medication
End Sub

Private Sub Command3_Click()
Unload Me
Main.Show
Main.Enabled = True
End Sub

Private Sub form_load()
Data1.DatabaseName = App.Path & "\Database\Database.mdb"
Data1.RecordSource = "Information"
Data1.Connect = ";pwd=nujAwlJa"
Data1.Refresh

Data2.DatabaseName = App.Path & "\Database\Database.mdb"
Data2.RecordSource = "Medication"
Data2.Connect = ";pwd=nujAwlJa"
Data2.Refresh

Data3.DatabaseName = App.Path & "\Database\Database.mdb"
Data3.RecordSource = "GenericName"
Data3.Connect = ";pwd=nujAwlJa"
Data3.Refresh

Data4.DatabaseName = App.Path & "\Database\Database.mdb"
Data4.RecordSource = "Stockin"
Data4.Connect = ";pwd=nujAwlJa"
Data4.Refresh

Data5.DatabaseName = App.Path & "\Database\Database.mdb"
Data5.RecordSource = "Stockin"
Data5.Connect = ";pwd=nujAwlJa"
Data5.Refresh

Data6.DatabaseName = App.Path & "\Database\Database.mdb"
Data6.RecordSource = "Stockout"
Data6.Connect = ";pwd=nujAwlJa"
Data6.Refresh

Data2.Recordset.MoveLast
ref = Data2.Recordset.Fields("MedRef") + 1
txtref.Text = ref
End Sub

Private Sub form_activate()
Data3.RecordSource = "Select * from stockin"
Do Until Data3.Recordset.EOF
List1.AddItem (Data3.Recordset.Fields("Generic"))
Data3.Recordset.MoveNext
Loop
Data1.Refresh

'Data4.RecordSource = "Select Brand_Name, Description, Quantity, Expiration, Date_Entry From Stockin"
'With Data4.Recordset
'Do Until Data4.Recordset.EOF
'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
MSFlexGrid2.FormatString = " BRAND NAME | DESCRIPTION | QUANTITY | EXPIRATION | DATE ENTRY  "
'Data4.Recordset.MoveNext
'Loop
'End With
End Sub

Private Sub List1_Click()
On Error Resume Next
Data4.RecordSource = "Select  Brand_Name, Description, Quantity, Expiration, Date_Entry, Ref_Med, Generic_Name from Stockin where Generic_Name = '" & List1.Text & "'"
Data4.Refresh
With Data4.Recordset
MSFlexGrid1.FormatString = " BRAND NAME | DESCRIPTION | QUANTITY |  EXPIRATION | DATE ENTRY  "
MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) = Data4.Recordset.Fields("Brand_Name")
MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1) = Data4.Recordset.Fields("Description")
MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 2) = Data4.Recordset.Fields("Quantity")
MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 3) = Data4.Recordset.Fields("Expiration")
MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4) = Data4.Recordset.Fields("Date_Entry")
MSFlexGrid2.Visible = False
End With
End Sub

Private Sub MSFlexGrid1_Click()
On Error Resume Next
MSFlexGrid1.FormatString = " BRAND NAME | DESCRIPTION | QUANTITY |  EXPIRATION | DATE ENTRY  "
txtBra.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0)
txtEx.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 3)
txtDes.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1)
txtQuan.SetFocus
End Sub

Private Sub Text5_Change()
On Error Resume Next
Data5.RecordSource = "Select * from Stockin where Brand_Name = '" & Text5.Text & "'"
Data5.Refresh
With Data5.Recordset
txtCode.Text = .Fields("Ref_Med")
txtGen.Text = .Fields("Generic_Name")
Text3.Text = .Fields("Brand_Name")
Text2.Text = .Fields("Description")
txtdate.Text = .Fields("Expiration")
Text1.Text = .Fields("Quantity")
End With
End Sub

Private Sub txtBra_Change()
Text5.Text = txtBra.Text
End Sub

Private Sub txtCode_Change()
On Error Resume Next
Data5.RecordSource = "Select Brand_Name from Stockin where Ref_Med = '" & txtCode.Text & "'"
Data5.Refresh
With Data5.Recordset
txtGen.Text = .Fields("Generic_Name")
Text3.Text = .Fields("Brand_Name")
Text2.Text = .Fields("Description")
End With
End Sub

Private Sub txtQuan_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   'Case 65 To 90, 97 To 122 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept letters. Number lang.", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub

Private Sub txtQuan_Change()
On Error Resume Next
Dim a, b, c

a = txtQuan.Text
b = Text1.Text

c = b - a
Text4.Text = Format(c)


End Sub

