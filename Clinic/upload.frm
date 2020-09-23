VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Find Personal Picture"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   Icon            =   "upload.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8295
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Upload Picture"
      ForeColor       =   &H00008000&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.DriveListBox Drive1 
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Select a drive to search in"
         Top             =   360
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         ForeColor       =   &H00004000&
         Height          =   3240
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Select a directory to search in"
         Top             =   840
         Width           =   2535
      End
      Begin VB.FileListBox filBox 
         ForeColor       =   &H00004000&
         Height          =   3795
         Left            =   2880
         Pattern         =   "*.bmp*; *.jpg*; *.gif*;*.BMP*;*.JPG*;*.GIF*"
         TabIndex        =   3
         ToolTipText     =   "Click a picture to work with"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "This is the path of the selected file"
         Top             =   4800
         Width           =   2535
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Enabled         =   0   'False
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
         Left            =   5160
         TabIndex        =   1
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Image imgCurrent 
         BorderStyle     =   1  'Fixed Single
         Height          =   3255
         Left            =   5160
         Stretch         =   -1  'True
         ToolTipText     =   "Picture preview"
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdUpload_Click()
Form1.imgCurrent.Picture = LoadPicture(txtPath.Text)
Unload Me
End Sub

Private Sub Dir1_Change()
' Change drive
filBox.Path = Dir1.Path
cmdUpload.Enabled = False
End Sub

Private Sub Drive1_Change()
' Change path

    On Error GoTo error                            ' if drive is not ready or other errors
    Dir1.Path = Drive1.Drive
    Exit Sub
    
error: MsgBox Err.Description + vbLf
Drive1.Refresh
End Sub

Private Sub filBox_Click()
' Lets you see a preview of the picture.  It sends the path
' to the picture box

    On Error GoTo error
       
    If Right(filBox.Path, 1) <> "\" Then                           ' if pic is not in root then "\" needed before filename
        txtPath.Text = filBox.Path & "\" & filBox.FileName
        Form1.txtPath.Text = filBox.Path & "\" & filBox.FileName
        Form6.txtPath.Text = filBox.Path & "\" & filBox.FileName
    Else
        txtPath.Text = filBox.Path & filBox.FileName          ' no "\" needed if pic is in root
    End If
Form1.txtPath = txtPath.Text
imgCurrent.Picture = LoadPicture(txtPath.Text)
cmdUpload.Enabled = True
Exit Sub

error: MsgBox Err.Description + vbLf
filBox.Refresh


End Sub

