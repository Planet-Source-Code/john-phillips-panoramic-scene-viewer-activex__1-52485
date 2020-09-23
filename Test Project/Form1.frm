VERSION 5.00
Object = "*\A..\PanoramicViewer.vbp"
Begin VB.Form Form1 
   Caption         =   "Test Project For Eye Poppers Panoramic Viewer ActiveX Control"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5880
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Scroll Right"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Scroll Left"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change Speed"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Text            =   "Scroll Direction"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Direction"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Text            =   "Scroll Speed"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto Scroll On"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide Title"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin PanoramicViewer.PanoViewer PanoViewer1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   9975
      _ExtentY        =   7011
      Appearance      =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483644
      BorderStyle     =   1
      MousePointer    =   5
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Show Title" Then
    Command1.Caption = "Hide Title"
    PanoViewer1.TitleVisible = True
    Exit Sub
ElseIf Command1.Caption = "Hide Title" Then
    Command1.Caption = "Show Title"
    PanoViewer1.TitleVisible = False
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Auto Scroll On" Then
    Command2.Caption = "Auto Scroll Off"
    Command4_Click
    PanoViewer1.AutoScroll True, pvLeft, pvNormal
    Command3.Enabled = True
    Command4.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = True
    Exit Sub
ElseIf Command2.Caption = "Auto Scroll Off" Then
    Command2.Caption = "Auto Scroll On"
    PanoViewer1.AutoScroll False
    Command3.Enabled = False
    Command4.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    Exit Sub
End If
End Sub

Private Sub Command3_Click()
Select Case Combo2.ListIndex
    Case 0
        PanoViewer1.ScrollDirection = pvLeft
    Case 1
        PanoViewer1.ScrollDirection = pvRight
    Case Else
        PanoViewer1.ScrollDirection = pvLeft
End Select
End Sub

Private Sub Command4_Click()
Select Case Combo1.ListIndex
    Case 0
        PanoViewer1.ScrollSpeed = pvVerySlow
    Case 1
        PanoViewer1.ScrollSpeed = pvSlow
    Case 2
        PanoViewer1.ScrollSpeed = pvNormal
    Case 3
        PanoViewer1.ScrollSpeed = pvFast
    Case 4
        PanoViewer1.ScrollSpeed = pvVeryFast
    Case Else
        PanoViewer1.ScrollSpeed = pvNormal
End Select
End Sub

Private Sub Command5_Click()
PanoViewer1.ScrollLeft
End Sub

Private Sub Command6_Click()
PanoViewer1.ScrollRight
End Sub

Private Sub Form_Load()

PanoViewer1.PicturePath = App.Path & "\St James Park.jpg"
PanoViewer1.InitializePano

PanoViewer1.Title = "This ActiveX Control Was Created By John Phillips"
PanoViewer1.TitleBackstyle = pvOpaque

With Combo1
    .AddItem "Very Slow"
    .AddItem "Slow"
    .AddItem "Normal"
    .AddItem "Fast"
    .AddItem "Very Fast"
    .ListIndex = 2
End With

With Combo2
    .AddItem "Left"
    .AddItem "Right"
    .ListIndex = 0
End With

End Sub

Private Sub PanoViewer1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then ' left arrow
    PanoViewer1.ScrollLeft
ElseIf KeyCode = 39 Then
    PanoViewer1.ScrollRight
End If
End Sub

Private Sub PanoViewer1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
End Sub
