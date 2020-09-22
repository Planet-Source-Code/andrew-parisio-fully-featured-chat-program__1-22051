VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   Caption         =   "Chat Server"
   ClientHeight    =   4350
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4455
   ForeColor       =   &H00000000&
   Icon            =   "server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Chat"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox outgoing 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox incoming 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   2415
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "127.0.0.1"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Listen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Text            =   "1210"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Connect"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton Addicon 
         Caption         =   "Minimize"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   3600
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   1680
         Top             =   1440
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3240
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Disconnected"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label label 
         Caption         =   "Port #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "MSG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Host:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   495
      End
      Begin MediaPlayerCtl.MediaPlayer media 
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   0   'False
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   0   'False
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin VB.Menu files 
      Caption         =   "&File"
      Begin VB.Menu mnuuser 
         Caption         =   "&User Name"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnunothin 
         Caption         =   "-----------"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuextras 
      Caption         =   "&Extras"
      Begin VB.Menu mnuupdate 
         Caption         =   "&Auto Update"
      End
      Begin VB.Menu mnucolor 
         Caption         =   "&Color"
         Begin VB.Menu mnublack 
            Caption         =   "&Black"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnublue 
            Caption         =   "&Blue"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnugreen 
            Caption         =   "&Green"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnured 
            Caption         =   "&Red"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnufile 
         Caption         =   "&File Transfer"
      End
      Begin VB.Menu mnuping 
         Caption         =   "&Ping"
      End
      Begin VB.Menu mnusound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Start with Windows"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnumode 
      Caption         =   "&Mode"
      Begin VB.Menu mnuserver 
         Caption         =   "&Server"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuclient 
         Caption         =   "&Client"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inputboxs As String, Msg As String
Dim inifile As String
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim heigh As Integer, widt As Integer
Private Sub Form_Resize()
Form1.SetFocus
If Form1.ActiveControl = True Then
Width = widt
Height = heigh
End If
End Sub
Private Sub mnuclient_click()
On Error Resume Next
mnuclient.Checked = True
mnuserver.Checked = False
Command3.Enabled = False
Command2.Enabled = True
Winsock1.Close
Winsock1.LocalPort = Text2
Winsock1.Listen
Label2.Caption = "Listening"
End Sub
Private Sub cmdBrowse_Click()
cdopen.ShowOpen
txtFileName.Text = cdopen.Filename
End Sub
Private Sub Command1_Click()
Winsock1.Close
End Sub
Private Sub Command5_Click()
filesend.Visible = True
End Sub
Private Sub nosuc()
msgboxs = MsgBox("File not saved try again", vbOKOnly, "Error")
End Sub
Private Sub Command7_Click()
Form3.Visible = True
End Sub
Private Sub print_click()
a = MsgBox("Is the printer ready?", vbYesNo, "?")
If a = vbNo Then
a = MsgBox("Maybe later then", vbOKOnly, "C YA")
Else
Printer.Print incoming
Printer.EndDoc
End If
End Sub
Private Sub incoming_dblclick()
incoming.Text = ""
End Sub
Private Sub mnublack_Click()
mnublack.Checked = True
mnublue.Checked = False
mnured.Checked = False
mnugreen.Checked = False
End Sub
Private Sub mnublue_Click()
mnublack.Checked = False
mnublue.Checked = True
mnured.Checked = False
mnugreen.Checked = False
End Sub
Private Sub mnucopy_Click()
If mnucopy.Checked = False Then
Call SetStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "Chat Program", App.path + "\" + App.EXEName + ".exe")
mnucopy.Checked = True
Else
Call DeleteStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "Chat Program")
mnucopy.Checked = False
End If
End Sub
Private Sub mnuexit_Click()
exits
End Sub

Private Sub mnufile_Click()
filesend.Visible = True
End Sub

Private Sub mnugreen_Click()
mnublack.Checked = False
mnublue.Checked = False
mnured.Checked = False
mnugreen.Checked = True
End Sub
Private Sub mnuping_Click()
Form3.Visible = True
End Sub
Private Sub mnuprint_Click()
print_click
End Sub
Private Sub mnured_Click()
mnublack.Checked = False
mnublue.Checked = False
mnured.Checked = True
mnugreen.Checked = False
End Sub
Private Sub mnusave_Click()
a = InputBox("Where would you like to save it (Default = C:\chat.txt)", "Save", "C:\chat.txt")
On Error GoTo nosuc
Open a For Output As #1
Write #1, incoming
Close #1
m = MsgBox("Chat session saved to: " + a, vbOKOnly, "Saved")
Exit Sub
nosuc:
msgboxs = MsgBox("File not saved try again", vbOKOnly, "Error")
Exit Sub
End Sub
Private Sub mnusound_Click()
If mnusound.Checked = False Then
mnusound.Checked = True
Else
 mnusound.Checked = False
 End If
End Sub
Private Sub mnuupdate_Click()
update.Visible = True
End Sub
Private Sub mnuuser_Click()
Dim a As String
a = InputBox("What is your nick name?", "User Name", inputboxs)
inputboxs = a
End Sub
Private Sub outgoing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Msg = "<" + inputboxs + ">" + outgoing.Text
outgoing.Text = ""
On Error Resume Next
Winsock1.SendData Msg
If mnublack.Checked = True Then incoming.ForeColor = &H80000017
If mnugreen.Checked = True Then incoming.ForeColor = &H4000&
If mnured.Checked = True Then incoming.ForeColor = &HFF&
If mnublue.Checked = True Then incoming.ForeColor = &HFF0000
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   incoming = incoming + Msg + vbCrLf
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   End If
End Sub
Private Sub Command2_Click()
Winsock1.Close
Winsock1.LocalPort = Text2.Text
Winsock1.Listen
Label2.Caption = "Listening"
End Sub
Private Sub Command3_Click()
On Error Resume Next
Winsock1.Close
Winsock1.Connect Text1, Text2
End Sub
Private Sub form_load()
widt = Width
heigh = Height
inifile = App.path + "\serverdata.txt"
On Error Resume Next
Open inifile For Input As #1
Input #1, names
Input #1, adres
Input #1, Port
Input #1, clientell
Input #1, Colors
Input #1, sound
Input #1, copyval
Close #1
If copyval = "True" Then mnucopy.Checked = True
If copyval = "False" Then mnucopy.Checked = False
If mnucopy.Checked = True Then Call SetStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "Chat Program", App.path + "\" + App.EXEName + ".exe")
inputboxs = names
mnublack.Checked = False
mnured.Checked = False
mnublue.Checked = False
mnugreen.Checked = False
If Colors = "Blue" Then mnublue.Checked = True
If Colors = "Red" Then mnured.Checked = True
If Colors = "Green" Then mnugreen.Checked = True
If Colors = "Black" Then mnublack.Checked = True
mnusound.Checked = sound
If clientell = "Client" Then
    mnuclient.Checked = True
    mnuserver.Checked = False
    Command2.Enabled = True
    Winsock1.LocalPort = Text2
    Winsock1.Listen
    Label2.Caption = "Listening"
    Else
    Winsock1.Connect Text1, Text2
    mnuserver.Checked = True
    mnuclient.Checked = False
End If
Text1.Text = adres
Text3.Text = names
Text2.Text = Port
Winsock1.LocalPort = Port
addicon_Click
Form1.Visible = True
End Sub
Private Sub mnuserver_Click()
mnuserver.Checked = True
mnuclient.Checked = False
On Error Resume Next
Winsock1.Close
Winsock1.Connect Text1, Text2
Command2.Enabled = False
Command3.Enabled = True
Label2.Caption = "Disconnected"
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Label2.Caption = "Connected"
Winsock1.Close
Winsock1.Accept requestID
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
ForeColor = &HC000&
If Form1.Visible = False Then
Form1.Visible = True
Form1.WindowState = 0
Form1.SetFocus
End If
Winsock1.GetData Msg
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   incoming = incoming + Msg + vbCrLf
   incoming.SelStart = Len(incoming.Text)
   incoming.SelLength = 0
   If mnusound.Checked = True Then
   media.Filename = App.path + "\phonaut.wav"
   media.Play
   End If
End Sub
Private Sub winsock1_connect()
Label2 = "Connected"
End Sub
Private Sub Winsock1_Close()
Winsock1.Close
Label2 = "Disconnected"
If mnuclient.Checked = True Then Winsock1.Listen
End Sub
'SYS TRAY ICON
Private Sub addicon_Click()
Form1.Visible = False
Dim NID As NOTIFYICONDATA
NID.hwnd = Me.hwnd
NID.cbSize = Len(NID)
NID.uID = vbNull
NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
NID.hIcon = Me.Icon
NID.uCallbackMessage = WM_MOUSEMOVE
NID.szTip = "Right-Click to display Popupmenu" & vbCrLf
Shell_NotifyIcon NIM_ADD, NID
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim Msg As Long
Msg = x / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONDOWN
        Me.WindowState = 0
        Form1.Visible = True
        AppActivate Me.Caption
        Case WM_RBUTTONUP
        Dim pAPI As POINTAPI
        Dim PMParams As TPMPARAMS

        GetCursorPos pAPI
        tmpPop% = CreatePopupMenu
        InsertMenu tmpPop%, 0, MF_BYPOSITION, 69, "Pinger"
        InsertMenu tmpPop%, 2, MF_BYPOSITION, 71, "Restore"
        InsertMenu tmpPop%, 3, MF_SEPARATOR, 72, vbNullString
        InsertMenu tmpPop%, 4, MF_BYPOSITION, 73, "Exit"
        
        PMParams.cbSize = 20
        tmpReply% = TrackPopupMenuEx(tmpPop%, TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RETURNCMD, pAPI.x, pAPI.y, Me.hwnd, PMParams)
        Select Case tmpReply%
            Case 69
                 Form3.Visible = True
            Case 71
                Me.WindowState = 0
                Form1.Visible = True
                AppActivate Me.Caption
            Case 73
                Call exits
                End
        End Select
    End Select
End Sub
Private Sub form_unload(cancel As Integer)
exits
End Sub
Public Sub exits()
On Error Resume Next
Open inifile For Output As #1
names = inputboxs
adres = Text1.Text
Port = Text2.Text
Write #1, names
Write #1, adres
Write #1, Port
If mnuclient.Checked = True Then Write #1, "Client"
If mnuserver.Checked = True Then Write #1, "Server"
If mnublue.Checked = True Then Colors = "Blue"
If mnured.Checked = True Then Colors = "Red"
If mnugreen.Checked = True Then Colors = "Green"
If mnublack.Checked = True Then Colors = "Black"
Write #1, Colors
If mnusound.Checked = True Then
    sound = "True"
    Else
    sound = "False"
    End If
Write #1, sound
copyval = "False"
If mnucopy.Checked = True Then copyval = "True"
Write #1, copyval
Close #1
Dim NID As NOTIFYICONDATA
NID.hwnd = Me.hwnd
NID.cbSize = Len(NID)
NID.uID = vbNull
NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
NID.hIcon = Me.Icon
NID.uCallbackMessage = WM_MOUSEMOVE
NID.szTip = "Right-Click to display Popupmenu" & vbCrLf
Shell_NotifyIcon NIM_DELETE, NID
Unload Form1
Unload Form3
Unload update
Unload filesend
End Sub

