VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm Main 
   BackColor       =   &H00004000&
   Caption         =   "Clini Management System"
   ClientHeight    =   6120
   ClientLeft      =   3315
   ClientTop       =   3300
   ClientWidth     =   9015
   Icon            =   "MDIEntrance.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   5160
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   688
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Current User:"
            TextSave        =   "Current User:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "3/5/2009"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "11:53 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogIn 
         Caption         =   "&Log-In"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuPatient 
      Caption         =   "&Patient"
      Enabled         =   0   'False
      Begin VB.Menu mnuEditinfo 
         Caption         =   "&Edit Patient Information"
      End
      Begin VB.Menu mnuMedication 
         Caption         =   "&Patient Medication"
      End
      Begin VB.Menu mnuMI 
         Caption         =   "&Medication Information"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "&Inventory"
      Enabled         =   0   'False
      Begin VB.Menu mnuStock 
         Caption         =   "New Stock"
      End
      Begin VB.Menu mnuInward 
         Caption         =   "Inward Report"
      End
      Begin VB.Menu mnuOutward 
         Caption         =   "&Outward Report"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Enabled         =   0   'False
      Begin VB.Menu mnuOPT 
         Caption         =   "&Option"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to quit the application?", vbQuestion + vbYesNo) = vbNo Then Cancel = True
End Sub
Private Sub MDIForm_Load()
With Form3
.Show
.Left = 300
.Top = 250
End With
End Sub

Private Sub mnuEditinfo_Click()
Form1.Show
Form1.Height = 9360
Form1.Width = 9930
End Sub

Private Sub mnuFileExit_Click()
If MsgBox("Are you Sure You want to Exit?", vbYesNo + vbQuestion) = vbYes Then Unload Me
'MsgBox ("Are you Sure You Want To Exit", MsgBoxStyle.OkCancel)
'Unload Me
End Sub

Private Sub mnuInward_Click()
'On Error Resume Next
'On Error GoTo Err
'DataEnvironment1.rsCommand2.Source = "Select Ref_Med From Stockin"
'DataEnvironment1.rsCommand2.Open
DataReport1.Show
'Exit Sub
'Err:
'DataEnvironment1.rsCommand2.Close
'DataEnvironment1.rsCommand2.Source = "Select Ref_Med From Stockin"
End Sub

Private Sub mnuLogIn_Click()
Form4.Show
Main.Enabled = False

End Sub

Private Sub mnuMedication_Click()
With Form9
.Show
End With
Main.Enabled = False
End Sub

Private Sub mnuMI_Click()
With Form6
.Show
End With
Main.Enabled = False
End Sub

Private Sub mnuNew_Click()
With Form1
.optStud.Enabled = True
.optEmp.Enabled = True
.Show
.txtvisit.Visible = False
.txtCourse.Visible = False
.txtBirthDate.Visible = False
.txtStatus.Visible = False
.txtSex.Visible = False
.txtDesignation.Visible = False
.txtDept.Visible = False
End With
Main.Enabled = False
End Sub

Private Sub mnuOPT_Click()
Form5.Show
End Sub

Private Sub mnuSearch_Click()
Form6.txtsearch.Visible = False
With Form6
.Label1.Visible = True
.Label2.Visible = False
.Show
.Left = 5500
.Top = 5000
.DTPicker1.Visible = True
.Command1.Visible = True
End With
End Sub

Private Sub mnuUpdate_Click()
Form6.Show
End Sub

Private Sub mnuOutward_Click()
DataReport2.Show
End Sub

Private Sub mnuStock_Click()
With Form7
.Show
Main.Enabled = False
End With
End Sub

Private Sub mnuStudentID_Click()
Form6.DTPicker1.Visible = False
With Form6
.Label1.Visible = False
.Label2.Visible = True
.Show
.Left = 5500
.Top = 5000
.txtsearch.Visible = True
.Command2.Visible = True
End With
End Sub

Private Sub Timer1_Timer()

Static iValue As Double
   Dim i As Long
    iValue = iValue + 1
    If iValue > 100 Then
Else
    'ProgressBar1.Value = iValue
    'Label6.Caption = iValue
    Exit Sub
  End If
    If Timer1.Enabled = False Then
    Else
   ' Unload Me
   ' Form3.Show
    
    End If

StatusBar1.Panels(2).Text = Date
StatusBar1.Panels(3).Text = Time
End Sub

