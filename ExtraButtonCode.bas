Attribute VB_Name = "Module2"
Option Explicit
Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
    Public Declare PtrSafe Function GetWindowLong _
                           Lib "user64" Alias "GetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong _
                           Lib "user64" Alias "SetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar _
                           Lib "user64" ( _
                           ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA _
                           Lib "user64" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As Long


