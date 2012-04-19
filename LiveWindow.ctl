VERSION 5.00
Begin VB.UserControl LiveWindow 
   ClientHeight    =   3588
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   ScaleHeight     =   3588
   ScaleWidth      =   4980
End
Attribute VB_Name = "LiveWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CapCreateCaptureWindow Lib "avicap32" Alias "capCreateCaptureWindowA" (ByVal lpszWndName As String, ByVal dwStyle As Long, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal hwnd As Long, ByVal nID As Integer) As Long

Const WS_CHILD = &H40000000
Const WS_VISIBLE = &H10000000
Const WM_USER = &H400
Public hWndC As Long
Private FrameDelay As Integer



Private Sub UserControl_Click()
    MsgBox "Digital Cam Acquisition Module...." + vbCrLf + "Developed By :" + vbCrLf + " sivamgr" + vbCrLf + " Free to Use"
End Sub

Private Sub UserControl_Initialize()
    hWndC = CapCreateCaptureWindow("LiveAcquire", WS_CHILD + WS_VISIBLE, 0, 0, 352, 288, UserControl.hwnd, 0)
    'SendMessage hWndC, WM_CAP_DRIVER_CONNECT, 0, 0
    SendMessage hWndC, WM_USER + 10, 0, 0
    'SendMessage hWndC, WM_CAP_SEQUENCE, 0, 0
    SendMessage hWndC, WM_USER + 62, 0, 0
    'SendMessage hWndC, WM_CAP_SET_PREVIEWRATE, 0, 0
    SendMessage hWndC, WM_USER + 52, 30, 0
    'SendMessage hWndC, WM_CAP_SET_PREVIEW, 0, 0
    SendMessage hWndC, WM_USER + 50, True, 0
End Sub


Public Property Get FrameRate() As Integer
    FrameRate = 1000 / FrameDelay
End Property

Public Property Let FrameRate(ByVal vNewValue As Integer)
    FrameDelay = 1000 / vNewValue
    'SendMessage hWndC, WM_CAP_SET_PREVIEWRATE, 0, 0
    SendMessage hWndC, WM_USER + 52, FrameDelay, 0
    'SendMessage hWndC, WM_CAP_SET_PREVIEW, 0, 0
    SendMessage hWndC, WM_USER + 50, True, 0
End Property

Private Sub UserControl_InitProperties()
    Let FrameRate = 30
End Sub

Private Sub UserControl_Resize()
    hWndC = CapCreateCaptureWindow("LiveAcquire", WS_CHILD + WS_VISIBLE, 0, 0, UserControl.Width, UserControl.Height, UserControl.hwnd, 0)
End Sub

Private Sub UserControl_Terminate()
    MsgBox "Digital Cam Acquisition Module...." + vbCrLf + "Developed By :" + vbCrLf + " sivamgr" + vbCrLf + " Free to Use"
End Sub

Public Sub CopyToClipboard()
    'SendMessage hWndC, WM_CAP_GRAB_FRAME_NOSTOP , 0, 0
    SendMessage hWndC, WM_USER + 61, 0, 0
    'SendMessage hWndC, WM_CAP_EDIT_COPY , 0, 0
    If (SendMessage(hWndC, WM_USER + 30, 0, 0)) Then MsgBox "success" Else MsgBox "failure"
    
End Sub

Public Sub CopyToFileBMP(Fname As String)
    'SendMessage hWndC, WM_CAP_GRAB_FRAME_NOSTOP , 0, 0
    SendMessage hWndC, WM_USER + 61, 0, 0
    'SendMessage hWndC, WM_CAP_FILE_SAVEDIB , 0, 0
    SendMessage hWndC, WM_USER + 25, 0, Fname
End Sub
