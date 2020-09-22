VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Update 
   Caption         =   "Form4"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   Icon            =   "update.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3045
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame update 
      Caption         =   "UPDATE/DOWNLOAD"
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "READ DIRECTIONS FIRST!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Download"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Directions"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame directions 
      Caption         =   "Directions"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      Begin VB.Label Label2 
         Caption         =   "READ THIS FIRST!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   $"update.frx":27A2
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1680
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public path As String
Private Sub Command1_Click()
Dim data() As Byte
Dim version1 As String, version2 As String
    path = "C:\temp.htm"
    Label4.Caption = "Finding newest Version"
    Open path For Binary As #1
                data() = Inet1.OpenURL("www.geocities.com/wospter/files/version.htm", icByteArray)
                Put #1, , data()
            Close #1
           
            
    Label4.Caption = "Checking Current Version"
    Open path For Input As #1
        Line Input #1, version2
        Close #1
       
     Open App.path + "\version.htm" For Input As #2
        Line Input #2, version1
        Close #2
       Label4.Caption = "Comparing Versions"
     
        If version1 = version2 Then
          Kill path
          Msg = MsgBox("There are no new updates", vbOKOnly, "Done")
            Label4.Caption = "NO UPDATES"
        Else
        path = "C:\server.zip"
        Label4.Caption = "Downloading Current Version"
            Open path For Binary As #1
                data() = Inet1.OpenURL("www.geocities.com/wospter/files/server.zip", icByteArray)
                Put #1, , data()
            Close #1
            msgboxans = MsgBox("To install extract C:/server.zip and copy the file in place of the original", vbOKCancel, "done")
        Unload Me
        End If
End Sub

Private Sub form_load()
directions.Visible = True
End Sub

Private Sub Option1_Click()
If Option1 = True Then
directions.Visible = True
update.Visible = False

End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
directions.Visible = False
update.Visible = True
End If
End Sub


