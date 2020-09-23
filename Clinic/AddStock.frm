VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form Form7 
   Caption         =   "Stocks"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form7"
   ScaleHeight     =   6165
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   7080
      Width           =   1455
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
      Top             =   7800
      Width           =   1935
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   615
      Left            =   3000
      TabIndex        =   21
      Top             =   7200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "AddStock.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColor       =   16384
      Enabled         =   0   'False
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
   Begin VB.Data Data1 
      Caption         =   "datint"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adding Entry"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
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
         Left            =   7920
         TabIndex        =   22
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtQuan 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
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
         Left            =   7920
         TabIndex        =   19
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         Left            =   7920
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtGen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53477377
         CurrentDate     =   39834
      End
      Begin VB.TextBox txtDes 
         Enabled         =   0   'False
         Height          =   765
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtBra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtCode 
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
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   75
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
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label5 
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
         Left            =   5040
         TabIndex        =   9
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label4 
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
         Left            =   5160
         TabIndex        =   8
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generic Name:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   1680
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   1560
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private id As Integer

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
'txtCode.Enabled = True
cmdEdit.Enabled = False
Command1.Enabled = True
txtGen.Enabled = True
txtBra.Enabled = True
txtDes.Enabled = True
txtQuan.Enabled = True
DTPicker1.Enabled = True
txtGen.SetFocus
End Sub

Private Sub cmdLeft_Click()
Dim i As Integer
i = MSFlexGrid1.Index
If i < MSFlexGrid1.Index - 1 Then
i = i + 1
MSFlexGrid1.TabIndex = i
End If
End Sub

Private Sub cmdRight_Click()
Dim i As Integer
i = List1.ListIndex
If i > 0 Then
i = i - 1
List1.ListIndex = i
End If
End Sub

Private Sub cmdEdit_Click()
cmdEdit.Enabled = False
cmdAdd.Enabled = False
cmdFind.Enabled = True
txtGen.Enabled = True
txtBra.Enabled = True
txtDes.Enabled = True
txtQuan.Enabled = True
DTPicker1.Enabled = True
txtGen.SetFocus
MSFlexGrid1.Enabled = True
End Sub

Private Sub cmdFind_Click()
With Data1.Recordset
.Edit
.Fields("Ref_Med") = txtCode.Text
.Fields("Generic_Name") = ProperCase(txtGen.Text)
.Fields("Brand_Name") = ProperCase(txtBra.Text)
.Fields("Description") = ProperCase(txtDes.Text)
.Fields("Quantity") = ProperCase(txtQuan.Text)
.Fields("Expiration") = DTPicker1
.Fields("Date_Entry") = Date
.Update
MsgBox "Has Been Updated"
Unload Me
Me.Show
End With
End Sub

Private Sub Command1_Click()
On Error Resume Next
Data2.RecordSource = "Select * from GenericName where Generic = '" & txtGen.Text & "'"
Data2.Refresh
With Data2.Recordset
If txtGen.Text = "" Or txtBra.Text = "" Then
MsgBox "Pls Fill All Information"
Else
Text1.Text = .Fields("Generic")
.AddNew
.Fields("Generic") = ProperCase(txtGen.Text)
.Update
End If
End With

With Data1.Recordset
If txtGen.Text = "" Or txtBra.Text = "" Then
MsgBox "Pls Fill All Information"
Else
.AddNew
.Fields("Ref_Med") = txtCode.Text
.Fields("Generic_Name") = ProperCase(txtGen.Text)
.Fields("Brand_Name") = ProperCase(txtBra.Text)
.Fields("Description") = ProperCase(txtDes.Text)
.Fields("Quantity") = ProperCase(txtQuan.Text)
.Fields("Expiration") = DTPicker1
.Fields("Date_Entry") = Date
.Update
MsgBox ("Has Been Save")
Unload Me
Me.Show
End If
End With
End Sub

Private Sub Command2_Click()
Unload Me
Main.Show
Main.Enabled = True
End Sub

Private Sub form_activate()
Data1.RecordSource = "Select * From stock"
With Data1.Recordset
Do Until Data1.Recordset.EOF
'MSFlexGrid1.AddItem "    " & Data1.Recordset.Fields("Ref_Med") & ""
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
MSFlexGrid1.FormatString = "ITEM CODE    |    GENERIC NAME   |   BRAND NAME   |   DESCRIPTION   |   QUANTITY  |   EXPIRATION   |   DATE ENTRY   "
Data1.Recordset.MoveNext
Loop
Label8 = Data1.Recordset.RecordCount
Label10 = txtCode.Text
End With
End Sub
Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\Database\Database.mdb"
Data1.RecordSource = "stockin"
Data1.Connect = ";pwd=nujAwlJa"
Data1.Refresh

Data2.DatabaseName = App.Path & "\Database\Database.mdb"
Data2.RecordSource = "GenericName"
Data2.Connect = ";pwd=nujAwlJa"
Data2.Refresh


Data1.Recordset.MoveLast
id = Data1.Recordset.Fields("Ref_Med") + 1
txtCode.Text = id

With ListView1
.ColumnHeaders.Add(, , "Item Code").Tag = "STRING"
.ColumnHeaders.Add(, , "Generic NAme").Tag = "STRING"
.ColumnHeaders.Add(, , "Item Code").Tag = "STRING"
End With
End Sub

Private Sub MSFlexGrid1_Click()
Data1.RecordSource = "Select * from Stockin"
Data1.Refresh
With Data1.Recordset
txtCode.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0)
txtGen.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1)
txtBra.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 2)
txtDes.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 3)
txtQuan.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4)
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If txtQuan.Text = "l" Then
Unload Form7
End If
End Sub

End Sub

Private Sub txtBra_KeyPress(KeyAscii As Integer)
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

Private Sub txtDes_KeyPress(KeyAscii As Integer)
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

Private Sub txtGen_KeyPress(KeyAscii As Integer)
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

Private Sub txtQuan_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    
   Case 48 To 57, 8  ' A-Z, 0-9 and backspace
   'Let these key codes pass through
   'Case 65 To 90, 97 To 122 'a-z and backspace
   'Let these key codes pass through
   Case Else 'All others get trapped
   MsgBox "Can not accept a letter", vbOKOnly, "warning"
   KeyAscii = 0 ' set ascii 0 to trap others input
   End Select
End Sub
