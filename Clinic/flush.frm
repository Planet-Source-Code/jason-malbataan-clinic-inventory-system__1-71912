VERSION 5.00
Begin VB.Form frmflush 
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8220
   Icon            =   "flush.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "flush.frx":6852
   ScaleHeight     =   6000
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   4680
   End
End
Attribute VB_Name = "frmflush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    Unload Me
    Main.Show
    End If
    

'Unload Me
'Main.Show
End Sub
