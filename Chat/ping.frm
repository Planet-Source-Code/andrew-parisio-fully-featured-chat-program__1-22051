VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Pinger"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3180
   Icon            =   "ping.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1440
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ping"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "32"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IP Host"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text > 65500 Then Text2.Text = 65500
Shell "Ping.exe -t -l " & Text2 & " " & Text1, vbNormalFocus
End Sub

