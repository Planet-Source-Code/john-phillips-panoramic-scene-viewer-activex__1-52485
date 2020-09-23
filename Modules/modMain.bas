Attribute VB_Name = "modMain"
Option Explicit

Public objPic(2) As StdPicture
Public bInvalidPano As Boolean
Public sPicPath As String

Public Type POINTAPI ' Used for GetCursor - gets mouse location
    x As Integer ' in screen coordinates.
    Y As Integer ' in screen coordinates.
End Type

'Public Type PicZoom
'    lngWidth As Long
'    lngHeight As Long
'End Type

Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)

Public Const GW_HWNDPREV = 3

Public Function Pix2TwipX(sPxls As Long) As Single

    Pix2TwipX = sPxls * Screen.TwipsPerPixelX
    
End Function

Public Function Pix2TwipY(sPxls As Long) As Single

    Pix2TwipY = sPxls * Screen.TwipsPerPixelY
    
End Function

Public Function Twip2PixX(sTwips As Single) As Long
    
    Twip2PixX = sTwips / Screen.TwipsPerPixelX
    
End Function

Public Function Twip2PixY(sTwips As Single) As Long

    Twip2PixY = sTwips / Screen.TwipsPerPixelY
    
End Function
