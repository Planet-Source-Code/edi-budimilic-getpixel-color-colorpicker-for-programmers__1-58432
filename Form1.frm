VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GetPixel Color v1.4"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FEEDEE&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   3855
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the textbox to save the value in clipboard"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   120
      Picture         =   "Form1.frx":0E42
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   120
      Width           =   3870
      Begin VB.Shape Shape1 
         Height          =   135
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
   Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = Me.hWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    OFName.lpstrFilter = "Bitmap files (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "Jpg files (*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = App.Path
    'Set the title
    OFName.lpstrTitle = "Open Picture - GetPixel Color"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
    ''    On Error Resume Next
        Picture1.Picture = LoadPicture(Trim$(OFName.lpstrFile))
    End If
    
        Dim tmpWidth As Integer, tmpHeight As Integer
        If Picture1.Height > 600 Then Picture1.Height = 600
        Form1.Height = (Picture1.Top * 15) + (Picture1.Height * 15) + (Frame1.Height * 15) + 600
        
        If Picture1.Width > 600 Then Picture1.Width = 600
        tmpWidth = (Picture1.Left * 15) + (Picture1.Width * 15) + 210
        If tmpWidth > (Frame1.Left * 15) + (Frame1.Width * 15) + 210 Then Form1.Width = tmpWidth

    Frame1.Top = Picture1.Top + Picture1.Height + 2
End Sub

Private Sub Form_Load()
    Form1.Left = 0
    Form1.Top = (Screen.Height) - Form1.Height - 600
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Do all only if left mouse button is down
    If Button = 1 Then
        CheckColor X, Y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Do all only if left mouse button is down
    If Button = 1 Then
        CheckColor X, Y
    End If
End Sub
Public Sub CheckColor(X, Y)
    'Clip mouse in the picture field
    If X < 1 Then X = 0
    If Y < 1 Then Y = 0
    If X > Picture1.Width - 3 Then X = Picture1.Width - 3
    If Y > Picture1.Height - 3 Then Y = Picture1.Height - 3

    'I'm creating variable
    Dim DetectColor As Long
    
    'Writing the color on which I clicked on into variable
    DetectColor = GetPixel(Picture1.hdc, X, Y)
    
    'Move Shape to mouse position
    Shape1.Left = X - (Shape1.Width \ 2)
    Shape1.Top = Y - (Shape1.Height \ 2)
    
    'Make Background color of Labels to be the color from the variable
    If DetectColor > 0 Then Label1.BackColor = DetectColor
    Dim ColorR As String, ColorG As String, ColorB As String
    
    'Show color value in text boxes
    ' RGB
    If DetectColor > 0 Then
        ColorR = Label1.BackColor And 255
        ColorG = (Label1.BackColor And 65280) / 256
        ColorB = (Label1.BackColor And 16711680) / 65535
        Text3.Text = Int(ColorR) & ", " & Int(ColorG) & ", " & Int(ColorB)
    End If
    ' HEX (VB)
    If DetectColor > 0 Then
        On Error Resume Next
        Dim Hr As String, Hg As String, Hb As String
        Hr = Hex$(ColorR)
        If Len(Hr) = 1 Then Hr = "0" & Hr
        Hg = Hex$(ColorG)
        If Len(Hg) = 1 Then Hg = "0" & Hg
        Hb = Hex$(ColorB)
        If Len(Hb) = 1 Then Hb = "0" & Hb
        Text2.Text = "&H" & Hr & Hg & Hb
    End If

End Sub

Private Sub Text2_Click()
    'Save to clipboard
    Clipboard.Clear
    Clipboard.SetText Text2.Text
End Sub

Private Sub Text3_Click()
    'Save to clipboard
    Clipboard.Clear
    Clipboard.SetText Text3.Text
End Sub
