VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patient Record Entry"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11220
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Patientinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   16384
      TabCaption(0)   =   "Personal Information"
      TabPicture(0)   =   "Patientinfo.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgCurrent"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label7(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Shape2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DTPicker2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "DTPicker1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAdd3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtLastName"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtsearch"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtFirstName"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtAge"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbCourse"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCourse"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtBirthdate"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbSex"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtSex"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmbStatus"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtStatus"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtAdd2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtAdd1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmbDesignation"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDesignation"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmbDept"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDept"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "optStud"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "optEmp"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "Contacts"
      TabPicture(1)   =   "Patientinfo.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtinadd"
      Tab(1).Control(1)=   "txtinname"
      Tab(1).Control(2)=   "txtinphone"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "Line1"
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(7)=   "Label9"
      Tab(1).Control(8)=   "Label10"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Family History"
      TabPicture(2)   =   "Patientinfo.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7(3)"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "Label1"
      Tab(2).Control(3)=   "cmbCheckup"
      Tab(2).Control(4)=   "txtvisit"
      Tab(2).Control(5)=   "txtAllergy"
      Tab(2).Control(6)=   "txtMedHist"
      Tab(2).ControlCount=   7
      Begin VB.OptionButton optEmp 
         Caption         =   "Employee"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7680
         TabIndex        =   59
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optStud 
         Caption         =   "Student"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6000
         TabIndex        =   58
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtinadd 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   -74400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Text            =   "Patientinfo.frx":68A6
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   3000
         TabIndex        =   54
         Top             =   6840
         Width           =   2535
      End
      Begin VB.ComboBox cmbDept 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         ItemData        =   "Patientinfo.frx":68AC
         Left            =   3000
         List            =   "Patientinfo.frx":68C8
         TabIndex        =   53
         Top             =   6840
         Width           =   2295
      End
      Begin VB.TextBox txtDesignation 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   3000
         TabIndex        =   52
         Top             =   6120
         Width           =   2535
      End
      Begin VB.ComboBox cmbDesignation 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         ItemData        =   "Patientinfo.frx":6961
         Left            =   3000
         List            =   "Patientinfo.frx":6986
         TabIndex        =   51
         Top             =   6120
         Width           =   2415
      End
      Begin VB.TextBox txtAdd1 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   50
         Top             =   2520
         Width           =   6255
      End
      Begin VB.TextBox txtAdd2 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   49
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox txtMedHist 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtAllergy 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -70200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox txtvisit 
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbCheckup 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Patientinfo.frx":6A4A
         Left            =   -74715
         List            =   "Patientinfo.frx":6A6C
         TabIndex        =   42
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtinname 
         Enabled         =   0   'False
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
         Left            =   -74400
         TabIndex        =   37
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtinphone 
         Enabled         =   0   'False
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
         Left            =   -74400
         TabIndex        =   36
         Top             =   5640
         Width           =   3495
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   32
         Top             =   4920
         Width           =   1695
      End
      Begin VB.ComboBox cmbStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Patientinfo.frx":6ABA
         Left            =   3000
         List            =   "Patientinfo.frx":6AC7
         TabIndex        =   31
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox txtSex 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5400
         TabIndex        =   30
         Top             =   4980
         Width           =   1695
      End
      Begin VB.ComboBox cmbSex 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Patientinfo.frx":6AE3
         Left            =   5400
         List            =   "Patientinfo.frx":6AED
         TabIndex        =   29
         Top             =   4980
         Width           =   1695
      End
      Begin VB.TextBox txtBirthdate 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   23
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtCourse 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5400
         TabIndex        =   22
         Top             =   4080
         Width           =   1695
      End
      Begin VB.ComboBox cmbCourse 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Patientinfo.frx":6AFF
         Left            =   5400
         List            =   "Patientinfo.frx":6B12
         TabIndex        =   21
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   7800
         TabIndex        =   20
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtFirstName 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtsearch 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   1020
         Width           =   2175
      End
      Begin VB.TextBox txtLastName 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   6300
         TabIndex        =   13
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtAdd3 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3000
         TabIndex        =   12
         Top             =   3240
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3000
         TabIndex        =   24
         Top             =   4080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   15990785
         CurrentDate     =   40155
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   8520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   15990785
         CurrentDate     =   39867
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "If the symptoms persist consult your Doctor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69600
         TabIndex        =   60
         Top             =   3600
         Width           =   3780
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   3855
      End
      Begin VB.Line Line1 
         X1              =   -70080
         X2              =   -70080
         Y1              =   1440
         Y2              =   6360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   56
         Top             =   3000
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "In Case of emergency, please notify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74400
         TabIndex        =   55
         Top             =   960
         Width           =   3180
      End
      Begin VB.Label Label1 
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
         Left            =   -74760
         TabIndex        =   48
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label Label2 
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
         Left            =   -70200
         TabIndex        =   47
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "No. of Visit:"
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
         Index           =   3
         Left            =   -74685
         TabIndex        =   46
         Top             =   720
         Width           =   1035
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
         Left            =   3000
         TabIndex        =   41
         Top             =   5880
         Width           =   1080
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
         Left            =   3000
         TabIndex        =   40
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   39
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tel. No.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   38
         Top             =   5280
         Width           =   885
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
         Left            =   2925
         TabIndex        =   35
         Top             =   4680
         Width           =   645
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
         Left            =   5430
         TabIndex        =   34
         Top             =   4740
         Width           =   405
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "{M/DD/YYYY}"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   150
         Index           =   5
         Left            =   3990
         TabIndex        =   28
         Top             =   3780
         Width           =   885
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
         Left            =   2955
         TabIndex        =   27
         Top             =   3780
         Width           =   945
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
         Left            =   5340
         TabIndex        =   26
         Top             =   3780
         Width           =   675
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
         Left            =   7815
         TabIndex        =   25
         Top             =   3780
         Width           =   435
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
         Left            =   2955
         TabIndex        =   18
         Top             =   780
         Width           =   1905
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
         Left            =   3000
         TabIndex        =   17
         Top             =   1560
         Width           =   1005
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
         Left            =   6240
         TabIndex        =   16
         Top             =   1560
         Width           =   1005
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
         Left            =   2985
         TabIndex        =   15
         Top             =   2280
         Width           =   765
      End
      Begin VB.Image imgCurrent 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   2415
         Left            =   360
         Stretch         =   -1  'True
         ToolTipText     =   "Picture preview"
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Click Here"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   2100
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   9615
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         Enabled         =   0   'False
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
         Left            =   6360
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   7800
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   3360
         Picture         =   "Patientinfo.frx":6B34
         TabIndex        =   7
         ToolTipText     =   "Name or Lastname"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
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
         Left            =   4920
         Picture         =   "Patientinfo.frx":D386
         TabIndex        =   6
         ToolTipText     =   "Name or Lastname"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      MaskColor       =   &H00004000&
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00004000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Width           =   1980
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "This is the path of the selected file"
      Top             =   9120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RefNum As Integer
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
Private Sub cmdAdd_Click()
Call enable
txtStudentID.SetFocus
End Sub
Private Sub enable()
txtStudentID.Enabled = True
Combo1.Enabled = True
txtLastName.Enabled = True
txtFirstName.Enabled = True
txtMiddle.Enabled = True
DTPicker1.Enabled = True
txtAge.Enabled = True
imgCurrent.Enabled = True
cmdSave.Enabled = True
cmdAdd.Enabled = False
cmdCancel.Enabled = True
cmdClose.Enabled = False
End Sub
Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to abort the current Registration?" & vbCrLf & "(NOTE: You will loose all information you have entered.)", vbCritical + vbYesNo) = vbYes Then
    Unload Me
    Main.Enabled = True
End If
End Sub
Private Sub cmdClose_Click()
Unload Me
Main.Show
Main.Enabled = True
End Sub

Private Sub cmdEdit_Click()
'enable textbox (Name\Address)
imgCurrent.Enabled = True
txtFirstName.Enabled = True
txtLastName.Enabled = True
txtAdd1.Enabled = True
txtAdd2.Enabled = True
txtAdd3.Enabled = True
'enable textbox (Contacts)
txtinname.Enabled = True
txtinadd.Enabled = True
txtinphone.Enabled = True
'enable combobox (Personal Info.)
txtvisit.Enabled = True
txtBirthDate.Enabled = True
txtStatus.Enabled = True
txtAge.Enabled = True
txtStatus.Enabled = True
txtSex.Enabled = True
txtCourse.Enabled = True
'enable textbox (Employment)
txtDesignation.Enabled = True
txtDept.Enabled = True
'enable textbox (Medical Info.)
txtMedHist.Enabled = True
txtAllergy.Enabled = True
Command3.Enabled = True
cmdEdit.Enabled = False


End Sub

Private Sub cmdSave_Click()
On Error Resume Next
With Data1.Recordset
If cmdSave.Caption Then
.AddNew
.Fields("StudentID") = ProperCase(txtStudentID)
.Fields("Course") = Combo1.Text
.Fields("FirstName") = ProperCase(txtFirstName.Text)
.Fields("LastName") = ProperCase(txtLastName.Text)
.Fields("Middle") = ProperCase(txtMiddle.Text)
.Fields("BirthDate") = DTPicker1
.Fields("Age") = txtAge.Text
.Fields("Picture") = txtPath.Text
.Update
Unload Me
Form1.Show

'Call ahahah
End If
End With
End Sub

Private Sub cmdUpdate_Click()
With Data1.Recordset
.Edit
.Fields("LastName") = ProperCase(txtLastName.Text)
.Fields("FirstName") = ProperCase(txtFirstName.Text)
.Fields("Course") = Combo1.Text
.Fields("BirthDate") = DTPicker1
.Fields("Picture") = txtPath.Text
.Fields("Age") = txtAge.Text
End With
End Sub
Private Sub Combo3_Change()
Combo3.List = "First"
End Sub

Private Sub cmdNew_Click()
If MsgBox("Are you Sure You want to Cancel?", vbYesNo + vbQuestion) = vbYes Then Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Data1.RecordSource = "Select * from Information where ID = '" & Text1.Text & "'"
Data1.Refresh
With Data1.Recordset
If Not Text1.Text = .Fields("ID") Then
MsgBox ("No Record Found")
Else
cmdEdit.Enabled = True
cmdSearch.Enabled = False
txtsearch.Text = .Fields("ID")
txtFirstName.Text = .Fields("FirstName")
txtLastName.Text = .Fields("LastName")
txtAdd1.Text = .Fields("Address1")
txtAdd2.Text = .Fields("Address2")
txtAdd3.Text = .Fields("Address3")
txtinname.Text = .Fields("Incase_Name")
txtinadd.Text = .Fields("Incase_Address")
txtinphone.Text = .Fields("Incase_phone")
txtPath.Text = .Fields("Picture")
Form1.imgCurrent.Picture = LoadPicture(txtPath.Text)
txtvisit.Text = .Fields("NoofVisit")
txtStatus.Text = .Fields("Status")
txtAge.Text = .Fields("Age")
txtBirthDate.Text = .Fields("BirthDate")
txtCourse.Text = .Fields("Course")
txtSex.Text = .Fields("Sex")
txtDesignation.Text = .Fields("Designation")
txtDept.Text = .Fields("Department")
txtMedHist.Text = .Fields("MedicalHist")
txtAllergy.Text = .Fields("Allergy")
End If
End With
End Sub

Private Sub Command1_Click()
With Data1.Recordset
On Error Resume Next
If txtsearch.Text = "" Then
    MsgBox ("You may have forgotten to specify your Student ID")
        ElseIf cmbCourse.Text = "" Then
            MsgBox ("You may have forgotten to specify your Course")
                ElseIf txtLastName.Text = "" Then
                    MsgBox ("You may have forgotten to specify your Last Name")
                ElseIf txtFirstName.Text = "" Then
            MsgBox ("You may have forgotten to specify your First Name")
        ElseIf txtAge.Text = "" Then
    MsgBox ("You may have forgotten to specify your Age")
Else
.AddNew
.Fields("ID") = txtsearch.Text
.Fields("FirstName") = ProperCase(txtFirstName.Text)
.Fields("LastName") = ProperCase(txtLastName.Text)
.Fields("Address1") = ProperCase(txtAdd1.Text)
.Fields("Address2") = ProperCase(txtAdd2.Text)
.Fields("Address3") = ProperCase(txtAdd3.Text)
.Fields("Picture") = txtPath.Text
.Fields("Incase_Name") = ProperCase(txtinname.Text)
.Fields("Incase_Address") = ProperCase(txtinadd.Text)
.Fields("Incase_phone") = txtinphone.Text
.Fields("NoofVisit") = ProperCase(cmbCheckup.Text)
.Fields("BirthDate") = DTPicker1
.Fields("Age") = txtAge.Text
.Fields("Status") = ProperCase(cmbStatus.Text)
.Fields("Sex") = ProperCase(cmbSex.Text)
.Fields("Course") = ProperCase(cmbCourse.Text)
.Fields("Designation") = ProperCase(txtDesignation.Text)
.Fields("Department") = ProperCase(txtDept.Text)
.Fields("MedicalHist") = ProperCase(txtMedHist.Text)
.Fields("Allergy") = ProperCase(txtAllergy.Text)
.Update
MsgBox ("Successfully Saved")
Unload Me
Me.Show
End If
End With
End Sub

Private Sub Command2_Click()
'If MsgBox("Are you Sure You want to Exit?", vbYesNo + vbQuestion) = vbYes Then
'Unload Me
'Else
Unload Me
Main.Show
Main.Enabled = True
'End If
End Sub

Private Sub Command3_Click()
With Data1.Recordset
.Edit
.Fields("ID") = txtsearch.Text
.Fields("FirstName") = ProperCase(txtFirstName.Text)
.Fields("LastName") = ProperCase(txtLastName.Text)
.Fields("Address1") = ProperCase(txtAdd1.Text)
.Fields("Address2") = ProperCase(txtAdd2.Text)
.Fields("Address3") = ProperCase(txtAdd3.Text)
.Fields("Picture") = txtPath.Text
.Fields("Incase_Name") = ProperCase(txtinname.Text)
.Fields("Incase_Address") = ProperCase(txtinadd.Text)
.Fields("Incase_phone") = txtinphone.Text
.Fields("NoofVisit") = ProperCase(cmbCheckup.Text)
.Fields("BirthDate") = DTPicker1
.Fields("Age") = txtAge.Text
.Fields("Status") = ProperCase(cmbStatus.Text)
.Fields("Sex") = ProperCase(cmbSex.Text)
.Fields("Course") = ProperCase(cmbCourse.Text)
.Fields("Designation") = ProperCase(txtDesignation.Text)
.Fields("Department") = ProperCase(txtDept.Text)
.Fields("MedicalHist") = ProperCase(txtMedHist.Text)
.Fields("Allergy") = ProperCase(txtAllergy.Text)
.Update
MsgBox ("Information Has Been Updated")
cmdSearch.Enabled = True
Command2.Enabled = True
Unload Me
Form1.Show
Form1.Height = 9360
Form1.Width = 9930
End With
End Sub

Private Sub DTPicker1_Change()
Dim age
Dim txt
Dim ewan As Double
Dim basura As Double

age = Date
ewan = DTPicker2.Year - DTPicker1.Year
txtAge.Text = Format(ewan)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If MsgBox("Are you sure you want to quit the application?", vbQuestion + vbYesNo) = vbNo Then Cancel = True
Main.Enabled = True
End Sub
Private Sub form_load()
Data1.DatabaseName = App.Path & "\Database\Database.mdb"
Data1.RecordSource = "Information"
Data1.Connect = ";pwd=nujAwlJa"
Data1.Refresh

'Data1.Recordset.MoveLast
'RefNum = Data1.Recordset.Fields("Reference") + 1
'Text1.Text = RefNum
End Sub
Private Sub Image1_Click()
Call enable
txtStudentID.SetFocus
End Sub

Private Sub imgCurrent_Click()
Form2.Show
End Sub

Private Sub Timer1_Timer()
Form1.Height = Form1.Height - 70
If Form1.Height = "510" Then
Timer1.Enabled = False
Unload Me
End If
End Sub

Private Sub Timer2_Timer()
Form1.Width = Form1.Width - 70
If Form1.Width = "1170" Then
Timer2.Enabled = False
End If
End Sub

Private Sub optEmp_Click()
txtinname.Enabled = True
txtinadd.Enabled = True
txtinphone.Enabled = True
'enable textbox (Name\Address)
imgCurrent.Enabled = True
txtsearch.Enabled = True
txtFirstName.Enabled = True
txtLastName.Enabled = True
txtAdd1.Enabled = True
txtAdd2.Enabled = True
txtAdd3.Enabled = True
'enable textbox (Contacts)
'txtPhone1.Enabled = True
'txtPhone2.Enabled = True
'enable combobox (Personal Info.)
cmbCheckup.Enabled = True
cmbCourse.Enabled = True
DTPicker1.Enabled = True
txtAge.Enabled = True
cmbStatus.Enabled = True
cmbSex.Enabled = True
'enable textbox (Employment)
'cmbDesignation.Enabled = True
'cmbDept.Enabled = True
'enable textbox (Medical Info.)
txtMedHist.Enabled = True
txtAllergy.Enabled = True
txtsearch.SetFocus
Command1.Enabled = True



cmbDesignation.Enabled = True
cmbDept.Enabled = True
End Sub

Private Sub optStud_Click()
txtinname.Enabled = True
txtinadd.Enabled = True
txtinphone.Enabled = True
'enable textbox (Name\Address)
imgCurrent.Enabled = True
txtsearch.Enabled = True
txtFirstName.Enabled = True
txtLastName.Enabled = True
txtAdd1.Enabled = True
txtAdd2.Enabled = True
txtAdd3.Enabled = True
'enable textbox (Contacts)
'txtPhone1.Enabled = True
'txtPhone2.Enabled = True
'enable combobox (Personal Info.)
cmbCheckup.Enabled = True
cmbCourse.Enabled = True
DTPicker1.Enabled = True
txtAge.Enabled = True
cmbStatus.Enabled = True
cmbSex.Enabled = True
'enable textbox (Employment)
'cmbDesignation.Enabled = True
'cmbDept.Enabled = True
'enable textbox (Medical Info.)
txtMedHist.Enabled = True
txtAllergy.Enabled = True
txtsearch.SetFocus
Command1.Enabled = True


cmbDesignation.Enabled = False
cmbDept.Enabled = False
End Sub

Private Sub txtAllergy_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   'Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   Case 65 To 90, 97 To 122, 32, 8, 127 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept a Number", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   'Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   Case 65 To 90, 97 To 122, 32, 8, 127 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept a Number", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub

Private Sub txtinname_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   'Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   Case 65 To 90, 97 To 122, 32, 8, 127 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept a Number", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   'Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   Case 65 To 90, 97 To 122, 32, 8, 127 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept a Number", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub

Private Sub txtPhone1_KeyPress(KeyAscii As Integer)
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

Private Sub txtPhone2_KeyPress(KeyAscii As Integer)
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

Private Sub txtMedHist_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   'Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   Case 65 To 90, 97 To 122, 32, 8, 127 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept a Number", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub
