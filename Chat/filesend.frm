VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form filesend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send A File"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "filesend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SendFile"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Listen"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1680
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock Clients 
      Left            =   1200
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   369
   End
   Begin MSWinsockLib.Winsock Servers 
      Left            =   720
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   369
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: Waiting"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "filesend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clients_DataArrival(ByVal bytesTotal As Long)
ReceiveData Clients
End Sub

Private Sub Command1_Click()
On Error GoTo NoFile
CD.ShowOpen
If CD.Filename = "" Then Exit Sub
' Get a free file number
FileNum = FreeFile
Open CD.Filename For Binary As #FileNum
SendFile Clients, CD.Filename, FileNum
NoFile:
End Sub
Private Sub clients_connect()
Label1.Caption = "Status: Connected"
End Sub
Private Sub clients_close()
Label1.Caption = "Status: Closed"
End Sub

Private Sub Command2_Click()
' Connect to the computer
Clients.Close
Clients.Connect Text1, 369
' disable the connect button, and enable the disconnect button
Command2.Enabled = False
Command3.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
Label1.Caption = "Status: Closed"
'disconnect
Clients.Close
' enable the connect button and disable the disconnect button
Command3.Enabled = False
Command2.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command4_Click()
Servers.Close
Servers.Listen
End Sub

Private Sub form_load()
Servers.Close
Servers.Listen
End Sub

Private Sub servers_Close()
Servers.Close
Servers.Listen
End Sub

Private Sub servers_ConnectionRequest(ByVal requestID As Long)
Servers.Close
Servers.Accept requestID
End Sub

Private Sub servers_DataArrival(ByVal bytesTotal As Long)
ReceiveData Servers
End Sub
