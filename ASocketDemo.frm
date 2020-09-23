VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ASocketDemo"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT_Command 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Text            =   "guest"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton BTN_Submit 
      Caption         =   "Send command"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox TXT_Host 
      Height          =   405
      Left            =   1920
      TabIndex        =   6
      Text            =   "library.uah.edu"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton BTN_Disconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton BTN_Connect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   4935
      Begin VB.Label TXT_Result 
         Caption         =   "..."
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Result:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Receive 
      Caption         =   "Received"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4935
      Begin VB.TextBox TXT_Received 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Text            =   "ASocketDemo.frx":0000
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Please visit our site at http://www.activxperts.com for more information, faq's and downloads."
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "(telnet server)"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' You need the FREEWARE ASocket.dll to run the sample.
' Download it from http://www.vahland.com/pub/asocket.dll
' and register it on your machine.
' Read http://www.vahland.com/pub/asocket.htm for more info
' DO NOT FORGET: ADD A REFERENCE TO THE ACTIVSOCKET LIBRARY FROM THE PROJECT->REFERENCE MENU

Public strRecv As String
Public lTimerID As Long

Private Sub BTN_Connect_Click()
    asObject.Connect TXT_Host, 23
    TXT_Result = "CONNECT: " & asObject.LastError
    
    If asObject.LastError = 0 Then
        lTimerID = SetTimer(0, 0, 200, AddressOf TimerProc)
        If lTimerID = 0 Then
          MsgBox "Could not create Timer"
        End If
    End If
End Sub

Private Sub BTN_Disconnect_Click()
    asObject.Disconnect
    TXT_Result = "DISCONNECT: " & asObject.LastError
    
    If lTimerID <> 0 Then
        lTimerID = KillTimer(0, lTimerID)
        lTimerID = 0
    End If
End Sub


Private Sub BTN_Submit_Click()
    asObject.SendString TXT_Command, True
    TXT_Result = "SEND: " & asObject.LastError
End Sub

Private Sub Form_Load()
    Set asObject = CreateObject("ActivXperts.Socket")
    asObject.Protocol = 2   ' Telnet protocol
    
    TXT_Result = "none"
    TXT_Received = ""
    
    lTimerID = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lTimerID <> 0 Then
        lTimerID = KillTimer(0, lTimerID)
        lTimerID = 0
    End If
End Sub
