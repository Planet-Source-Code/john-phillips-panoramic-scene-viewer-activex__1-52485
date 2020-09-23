VERSION 5.00
Begin VB.UserControl PanoViewer 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "PanoViewer.ctx":0000
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   2880
   End
   Begin VB.Label lblProcessing 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Processing. . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "PanoViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Default Property Values:
Const m_def_AutoSize = True
Const m_def_TitlePosition = 0
Const m_def_ScrollSpeed = 5
Const m_def_ScrollDirection = 0
Const m_def_PicturePath = ""
Const m_def_TitleVisible = True
' End Default Property Values

' Private Declerations
Dim lX As Single, lY As Single ' Location Holders
Dim lCurX As Long ' Current X position on Screen
Dim lCurY As Long ' Current Y position on Screen
Dim lMaxX As Long ' Maxium Amount Allowed of Value X
Dim bLeft As Boolean ' Boolean True = Scroll Left, False = Scroll Right
Dim lStartX As Long ' The Starting X value when the user presses the mouse down
Dim lStartY As Long ' The Starting Y value when the user presses the mouse down
Dim lNewX As Long, lNewY As Long ' The New X,Y values when the user moves the mouse with the button down
Dim bScroll As Boolean ' Boolean True = Scroll, False = Not Scrolling
' Dim bZoom As Boolean ' User for zoom option
' Dim lZoomLvl As Long ' The Current Zoom Level
' Dim objZoom() As PicZoom ' the different Zoom Levels
Dim ScrollInterval As PVScrollInter ' The Scroll interval
Dim bPointerSelect As Boolean ' Auto Select Mouse Pointer
' End Private Declerations

' Public User Defined Types
Public Type CenterPic ' Holds the coordinates for the Pano Image Centered
    sngLeft As Single
    sngTop As Single
End Type
' End User Defined Types

' Public Enumerations:
Public Enum PVScrollDir ' Scroll Direction Enumeration
    pvLeft = 0
    pvRight = 1
    pvUp = 2
    pvDown = 3
End Enum

Public Enum PVPointer ' Mouse Pointer Enumeration
    pvDefault = 0
    pvArrow = 1
    pvSizeAll = 5
    pvSizeWE = 9
    pvAutoSelect = 100
End Enum

Public Enum PVScrollSpeed 'Scroll Speed Enumeration
    pvVerySlow = 100
    pvSlow = 50
    pvNormal = 25
    pvFast = 10
    pvVeryFast = 1
End Enum

Public Enum PVScrollInter ' Scroll interval Enumeration
    pvVS = 50
    pvS = 100
    pvN = 200
    pvF = 300
    pvVF = 500
End Enum

Public Enum PVApperance ' User Control Appearance Enumeration
    pv3D = 1
    pvFlat = 0
End Enum

Public Enum PVBorderstyle ' User Control BorderStyle Enumeration
    pvFlat = 1
    pvNone = 0
End Enum

Public Enum PVTitleBackStyle ' Title Background Enumeration
    pvTransparent = 0
    pvOpaque = 1
End Enum

Public Enum PVTitlePosition ' Title Position Enumeration
    pvTop = 0
    pvBottom = 1
End Enum
' End Public Enumerations

' Property Variables:
Dim m_AutoSize As Boolean
Dim m_PicturePath As String
Dim m_TitlePosition As PVTitlePosition
Dim m_ScrollSpeed As PVBorderstyle
Dim m_ScrollDirection As PVBorderstyle
Dim m_CentPic As CenterPic
Dim m_TitleVisible As Boolean
' End Property Variables

' Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
' End Event Declarations

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As PVApperance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As PVApperance)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    ' The auto size will only set the size
    ' for the height of the graphic
    AutoSizeControl
End Property

Private Sub AutoSizeControl()
On Error GoTo errNotSet

' Make sure there is a picture loaded into the object
If objPic(0) > 0 Then
    Dim yHeight As Long
    
    ' Convert the height from himetric (8) to twips (1)
    yHeight = UserControl.ScaleY(objPic(0).Height, 8, 1)
    
    ' Set the height of the user control
    UserControl.Height = yHeight

End If

Exit Sub
errNotSet:

    If Err.Number = 91 Then Exit Sub

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As PVBorderstyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As PVBorderstyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub lblTitle_Click()
    UserControl_Click
End Sub

Private Sub tmrScroll_Timer()

' Change the current Position of the Pano Image
If bLeft = True Then
    lCurX = lCurX + ScrollInterval
Else
    lCurX = lCurX - ScrollInterval
End If

DrawPano ' Draw the Pano Image to the Control

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
    UserControl.Picture = LoadPicture() ' Clear The Picture
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_ExitFocus()
    UserControl.Refresh
End Sub

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
    
If Button = 1 Then
    ' Turn the Scroll Timer Off
    tmrScroll.Enabled = False
    ' Set the scroll boolean to true
    bScroll = True
    ' Turn the zoom boolean off
    ' bZoom = False
    
    ' Get the starting X, Y Position
    lStartX = x
    lStartY = Y

ElseIf Button = 2 Then
    
    ' The Zoom features were taken out
    ' set the scroll boolean off
    'bScroll = False
    ' set the zoom boolean on
    'bZoom = True
    
    ' Get the starting X, Y Position
    lStartX = x
    lStartY = Y
    
End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)

' AH, I forgot what I did in here ???????
' ok joking aside here it goes

If Button = 1 Then ' left mouse button is pressed
    lNewX = x ' get the new X poisitio
    lNewY = Y ' Get the new Y position
    
    ' Make sure that the values are set, otherwise set them to defaults
    If ScrollSpeed <= 0 Then ScrollSpeed = pvNormal
    If ScrollInterval <= 0 Then ScrollInterval = pvN
    
    ' If the new X value is greater then the Starting X value Scroll Right
    If lNewX > lStartX Then
    ' scroll right
        bLeft = False
    ElseIf lNewX < lStartX Then ' the opposite of the above comment :)
        bLeft = True
    End If
    
    ' Enable the timer
    tmrScroll.Enabled = True
ElseIf Button = 2 Then ' Right Mouse Button 'Zooming'
    
    ' Get the current Y value
    lCurY = Y
    
    ' The Zoom features were taken out
    ' Must i really explain all of this???
    ' If the user moves the mouse up, It zooms In
    ' else if the user mouse the mouse down, it zooms out
    'If lStartY < lCurY Then
    '    bZoomIn = False
    'Else
    '    bZoomIn = True
    'End If
    
    lStartY = lCurY
    
    'If bZoom = True Then
    '    If bZoomIn = True Then
    '        ZoomIn ' Zoom in on the pano graphic
    '    Else
    '        ZoomOut ' Zoom out on the pano graphic
    '    End If
    'End If
    
End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As PVPointer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As PVPointer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    If New_MousePointer = pvAutoSelect Then
        bPointerSelect = True
    Else
        bPointerSelect = False
    End If
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)

If Button = 1 Then
    ' turn off the scrolling
    tmrScroll.Enabled = False
    bScroll = False
ElseIf Button = 2 Then
    ' The Zoom features were taken out
    ' lCurX = UserControl.Width / 2
    ' lCurY = UserControl.Height / 2
    ' bZoom = False
End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    
    ' Basic resize events in here
    
    If TitlePosition = pvTop Then
        lblTitle.Top = 0
    Else
        lblTitle.Top = UserControl.Height - lblTitle.Height
    End If
    
        lblTitle.Left = 0
        lblTitle.Width = UserControl.Width
        
        lblProcessing.Top = (UserControl.Height / 2) - (lblProcessing.Height / 2)
        lblProcessing.Left = 0
        lblProcessing.Width = UserControl.Width
        
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ScrollLeft()
Attribute ScrollLeft.VB_Description = "Scrolls the Pano Left"
lCurX = lCurX + ScrollInterval
' Self explanetroy
DrawPano
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ScrollRight()
Attribute ScrollRight.VB_Description = "Scrolls the Pano Right"
lCurX = lCurX - ScrollInterval
' Self explanetroy
DrawPano
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ScrollUp()
Attribute ScrollUp.VB_Description = "Scrolls the Pano Upward"
' Taken out of this version
' due to the problems with the zoom
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ScrollDown()
Attribute ScrollDown.VB_Description = "Scrolls the Pano Downward"
' Taken out of this version
' due to the problems with the zoom
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Caption
Public Property Get Title() As String
Attribute Title.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Title = lblTitle.Caption
End Property

Public Property Let Title(ByVal New_Title As String)
    lblTitle.Caption() = New_Title
    PropertyChanged "Title"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TitlePosition() As PVTitlePosition
Attribute TitlePosition.VB_Description = "Set the position of the Title Text"
    TitlePosition = m_TitlePosition
End Property

Public Property Let TitlePosition(ByVal New_TitlePosition As PVTitlePosition)
    m_TitlePosition = New_TitlePosition
    PropertyChanged "TitlePosition"
    
    ' Change the title position
    If New_TitlePosition = pvTop Then
        lblTitle.Top = 0
    Else
        lblTitle.Top = UserControl.Height - lblTitle.Height
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,BackStyle
Public Property Get TitleBackstyle() As PVTitleBackStyle
Attribute TitleBackstyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    TitleBackstyle = lblTitle.BackStyle
End Property

Public Property Let TitleBackstyle(ByVal New_TitleBackstyle As PVTitleBackStyle)
    lblTitle.BackStyle() = New_TitleBackstyle
    PropertyChanged "TitleBackstyle"
    lblTitle.BackStyle = New_TitleBackstyle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Font
Public Property Get TitleFont() As Font
Attribute TitleFont.VB_Description = "Returns a Font object."
    Set TitleFont = lblTitle.Font
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
    Set lblTitle.Font = New_TitleFont
    PropertyChanged "TitleFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,ForeColor
Public Property Get TitleForecolor() As OLE_COLOR
Attribute TitleForecolor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TitleForecolor = lblTitle.ForeColor
End Property

Public Property Let TitleForecolor(ByVal New_TitleForecolor As OLE_COLOR)
    lblTitle.ForeColor() = New_TitleForecolor
    PropertyChanged "TitleForecolor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AutoSize = m_def_AutoSize
    m_TitlePosition = m_def_TitlePosition
    m_ScrollSpeed = m_def_ScrollSpeed
    m_ScrollDirection = m_def_ScrollDirection
    m_picturepathPath = m_def_PicturePath
    m_TitleVisible = m_def_TitleVisible
    ScrollInterval = pvN
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    lblTitle.Caption = PropBag.ReadProperty("Title", "Title")
    m_TitlePosition = PropBag.ReadProperty("TitlePosition", m_def_TitlePosition)
    lblTitle.BackStyle = PropBag.ReadProperty("TitleBackstyle", 0)
    Set lblTitle.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
    lblTitle.ForeColor = PropBag.ReadProperty("TitleForecolor", &H80000012)
    lblProcessing.Caption = PropBag.ReadProperty("ProcessingMessage", "Processing. . .")
    m_ScrollSpeed = PropBag.ReadProperty("ScrollSpeed", m_def_ScrollSpeed)
    m_ScrollDirection = PropBag.ReadProperty("ScrollDirection", m_def_ScrollDirection)
    m_PicturePath = PropBag.ReadProperty("PicturePath", m_def_PicturePath)
    m_TitleVisible = PropBag.ReadProperty("TitleVisible", m_def_TitleVisible)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 100)
    Call PropBag.WriteProperty("Title", lblTitle.Caption, "Title")
    Call PropBag.WriteProperty("TitlePosition", m_TitlePosition, m_def_TitlePosition)
    Call PropBag.WriteProperty("TitleBackstyle", lblTitle.BackStyle, 0)
    Call PropBag.WriteProperty("TitleFont", lblTitle.Font, Ambient.Font)
    Call PropBag.WriteProperty("TitleForecolor", lblTitle.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ProcessingMessage", lblProcessing.Caption, "Processing. . .")
    Call PropBag.WriteProperty("ScrollSpeed", m_ScrollSpeed, m_def_ScrollSpeed)
    Call PropBag.WriteProperty("ScrollDirection", m_ScrollDirection, m_def_ScrollDirection)
    Call PropBag.WriteProperty("PicturePath", m_PicturePath, m_def_PicturePath)
    Call PropBag.WriteProperty("TitleVisible", m_TitleVisible, m_def_TitleVisible)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblProcessing,lblProcessing,-1,Caption
Public Property Get ProcessingMessage() As String
Attribute ProcessingMessage.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    ProcessingMessage = lblProcessing.Caption
End Property

Public Property Let ProcessingMessage(ByVal New_ProcessingMessage As String)
    lblProcessing.Caption() = New_ProcessingMessage
    PropertyChanged "ProcessingMessage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,0
Public Property Get ScrollSpeed() As PVScrollSpeed
Attribute ScrollSpeed.VB_Description = "Sets the Speed for autoscroll"
    ScrollSpeed = m_ScrollSpeed
End Property

Public Property Let ScrollSpeed(ByVal New_ScrollSpeed As PVScrollSpeed)
    m_ScrollSpeed = New_ScrollSpeed
    PropertyChanged "ScrollSpeed"
    ' chane the scroll speed
    tmrScroll.Interval = New_ScrollSpeed
    Select Case m_ScrollSpeed
        Case 1
            ScrollInterval = pvVF
        Case 10
            ScrollInterval = pvF
        Case 25
            ScrollInterval = pvN
        Case 50
            ScrollInterval = pvS
        Case 100
            ScrollInterval = pvVS
        Case Else
            ScrollInterval = pvN
    End Select
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,0
Public Property Get ScrollDirection() As PVScrollDir
Attribute ScrollDirection.VB_Description = "Sets the direction for autoscroll"
    ScrollDirection = m_ScrollDirection
End Property

Public Property Let ScrollDirection(ByVal New_ScrollDirection As PVScrollDir)
    m_ScrollDirection = New_ScrollDirection
    PropertyChanged "ScrollDirection"
    ' change the scroll direction
    Select Case New_ScrollDirection
        Case 0
            bLeft = True
        Case 1
            bLeft = False
        Case Else
            bLeft = True
    End Select
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AutoScroll(bOn As Boolean, Optional pvDir As PVScrollDir = pvLeft, Optional pvSpd As PVScrollSpeed = 100) As Boolean
' Auto scroll timer on or off
If bOn = True Then
    tmrScroll.Interval = pvSpd
    tmrScroll.Enabled = True
Else
    tmrScroll.Enabled = False
End If
End Function

Private Sub ConvertPicWH()
' get the center coordinates for the graphic
lX = UserControl.ScaleX(objPic(0).Width, vbHimetric, vbTwips)
lY = UserControl.ScaleY(objPic(0).Height, vbHimetric, vbTwips)

m_CentPic.sngTop = (UserControl.Height / 2) - (lY / 2)
m_CentPic.sngLeft = (UserControl.Width / 2) - (lX / 2)

End Sub

Public Sub CenterPano()

' center the pano graphic on the control
lblProcessing.Visible = False

UserControl.Cls ' Clear the picture

If m_PicturePath <> "" Then
    UserControl.PaintPicture objPic(1), m_CentPic.sngLeft, m_CentPic.sngTop
End If

End Sub

Public Sub DrawPano()
Dim xc As Long

' Clear the control
UserControl.Cls

If lCurX <= lMaxX Then
' draw the second picture at the end of the first

xc = lX + lCurX

If xc <= 0 Then
    lCurX = xc
End If

    UserControl.PaintPicture objPic(0), lCurX, -((lY / 2) - (UserControl.Height / 2))
    UserControl.PaintPicture objPic(0), xc, -((lY / 2) - (UserControl.Height / 2))

ElseIf lCurX > 0 And lCurX < UserControl.Width Then
' draw the second picture (end part) at the begining of the first

xc = lCurX - lX

If xc >= UserControl.Width Then
    lCurX = xc
End If

        UserControl.PaintPicture objPic(0), lCurX, -((lY / 2) - (UserControl.Height / 2))
        UserControl.PaintPicture objPic(0), xc, -((lY / 2) - (UserControl.Height / 2))

Else
' just draw the picture

If lCurX >= UserControl.Width Then
    lCurX = -(lX - UserControl.Width) + (lCurX - UserControl.Width)
End If

UserControl.PaintPicture objPic(0), lCurX, -((lY / 2) - (UserControl.Height / 2))

End If

UserControl.Refresh

End Sub

Public Sub InitializePano()
If m_PicturePath <> "" Then
    Dim lTempAmtW As Long
    Dim lTempAmtH As Long
    
    lblProcessing.Visible = True
            
    Set objPic(0) = LoadPicture(m_PicturePath)
    Set objPic(1) = LoadPicture(m_PicturePath)
    
    ConvertPicWH
    
    lTempAmtW = objPic(0).Width
    lTempAmtH = objPic(0).Height
        
    If bPointerSelect = True Then
        If lTempAmtW > UserControl.Width And lTempAmtH <= UserControl.Height Then
            MousePointer = pvSizeWE
        ElseIf lTempAmtW <= UserControl.Width And lTempAmtH <= UserControl.Height Then
            MousePointer = pvArrow
        ElseIf lTempAmtW > UserControl.Width And lTempAmtH > UserControl.Height Then
            MousePointer = pvSizeAll
        End If
    End If
    
    If AutoSize = True Then
        AutoSizeControl
    End If
    
    UserControl.PaintPicture objPic(0), m_CentPic.sngLeft, m_CentPic.sngTop
    'UserControl.Refresh
    
    lCurX = -((lX / 2) - (UserControl.Width / 2))
    
    lMaxX = -(lX - UserControl.Width)
    
    lblProcessing.Visible = False
    
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get PicturePath() As String
Attribute PicturePath.VB_Description = "Returns/Sets the file path to a panoramic picture."
    PicturePath = m_PicturePath
    sPicPath = m_PicturePath
End Property

Public Property Let PicturePath(ByVal New_PicturePath As String)
    m_PicturePath = New_PicturePath
    PropertyChanged "PicturePath"
    sPicPath = m_PicturePath
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get TitleVisible() As Boolean
    TitleVisible = m_TitleVisible
End Property

Public Property Let TitleVisible(ByVal New_TitleVisible As Boolean)
    m_TitleVisible = New_TitleVisible
    PropertyChanged "TitleVisible"
    lblTitle.Visible = New_TitleVisible
End Property

