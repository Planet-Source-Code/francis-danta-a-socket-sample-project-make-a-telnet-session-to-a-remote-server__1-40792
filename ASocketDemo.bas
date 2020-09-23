Attribute VB_Name = "Module1"

Option Explicit

Declare Function SetTimer Lib "user32" _
       (ByVal hwnd As Long, _
       ByVal nIDEvent As Long, _
       ByVal uElapse As Long, _
       ByVal lpTimerFunc As Long) As Long

Declare Function KillTimer Lib "user32" _
       (ByVal hwnd As Long, _
       ByVal nIDEvent As Long) As Long

Global strRecvData As String
Global asObject As ASOCKETLib.Socket

 Sub TimerProc(ByVal hwnd As Long, _
                ByVal uMsg As Long, _
                ByVal idEvent As Long, _
                ByVal dwTime As Long)

     strRecvData = asObject.ReceiveString
     If strRecvData <> "" Then
        Form1.TXT_Received.Text = Form1.TXT_Received.Text & strRecvData
     End If
 End Sub


