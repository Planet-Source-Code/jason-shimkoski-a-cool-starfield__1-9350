VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Starfield"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2295
      Left            =   480
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a really simple and easy way to create a starfield.
'I wouldn't use it in a game as it would be to slow, but by
'itself it is pretty cool to look at.
'
'Thanks,
'Jason Shimkoski (basspler@aol.com)

Private Sub Form_Load()
    Form_Resize
    'this initializes the starfield
    InitStars
End Sub

Private Sub Form_Resize()
    
    'this will get rid of any errors when the form is maximized or minimized
    On Error Resume Next
    
    'this resizes the form to fit to the screen
    With frmMain
    .Top = 0
    .Left = 0
    .Width = Screen.Width
    .Height = Screen.Height
    End With
    
    'this resizes the picture box to fit to the form
    With picField
    .Top = frmMain.ScaleTop
    .Left = frmMain.ScaleLeft
    .Height = frmMain.ScaleHeight
    .Width = frmMain.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub picField_Click()
    'this is the main program loop
    Do
    DoEvents
    DrawStars
    Loop
End Sub
