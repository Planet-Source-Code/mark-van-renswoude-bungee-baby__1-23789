VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBaby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' API declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' Properties
Private oDest As Object

' Declarations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private pBaby As PictureBox
Private tMiddle As POINTAPI
Private tMouse As POINTAPI
Private tBaby As POINTAPI
Private tSpeed As POINTAPI
Private lFrameWidth As Long
Private lFrame As Long

Public Sub Blit()
    Dim pSrc As POINTAPI
    Dim pDest As POINTAPI
    
    ' Clear control
    oDest.Cls
    
    ' Draw baby frame
    Call BitBlt(oDest.hdc, tBaby.x - tMiddle.x, tBaby.y - tMiddle.y, lFrameWidth, pBaby.ScaleHeight, pBaby.hdc, lFrameWidth * lFrame, 0, vbSrcCopy)
    
    ' Copy values to temporary variables
    ' Somehow, the MoveToEx or LineTo API changes the values,
    ' so we use temporary variables to keep the original.
    pSrc = tMouse
    pDest = tBaby
    
    ' Draw line
    Call MoveToEx(oDest.hdc, pSrc.x, pSrc.y, pSrc)
    Call LineTo(oDest.hdc, pDest.x, pDest.y)
    
    ' Apply changes
    oDest.Refresh
End Sub

Private Function CreatePoint(newX As Long, newY As Long) As POINTAPI
    CreatePoint.x = newX
    CreatePoint.y = newY
End Function

Public Property Get Destination() As Object
    Set Destination = oDest
End Property

Public Property Let Destination(Value As Object)
    Set oDest = Value
    
    ' Adjust the object
    oDest.AutoRedraw = True
End Property


Public Property Get Frame() As Long
    Frame = lFrame
End Property

Public Property Let Frame(Value As Long)
    lFrame = Value
End Property

Public Property Get FrameWidth() As Long
    FrameWidth = lFrameWidth
End Property

Public Property Let FrameWidth(Value As Long)
    lFrameWidth = Value
End Property
Public Sub MoveBaby()
    ' Move baby
    tSpeed.x = tSpeed.x + CLng((tMouse.x - tBaby.x) / 50) - tSpeed.x / 25
    tSpeed.y = tSpeed.y + CLng((tMouse.y - tBaby.y) / 50) - tSpeed.y / 25
    tBaby.x = tBaby.x + tSpeed.x
    tBaby.y = tBaby.y + tSpeed.y
End Sub

Public Sub MoveCursor(x As Long, y As Long)
    ' Move cursor
    tMouse = CreatePoint(x, y)
End Sub


Public Property Get Source() As PictureBox
    Source = pBaby
End Property

Public Property Let Source(Value As PictureBox)
    Set pBaby = Value
    
    pBaby.ScaleMode = vbPixels
    tMiddle.x = lFrameWidth / 2
    tMiddle.y = lFrameWidth / 2
End Property
Private Sub Class_Initialize()
    ' Initialize variables
    tMouse = CreatePoint(0, 0)
    tBaby = CreatePoint(-50, 0)
    tSpeed = CreatePoint(0, 0)
End Sub


Private Sub Class_Terminate()
    ' Clean up
    Set pBaby = Nothing
End Sub


