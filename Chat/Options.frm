VERSION 5.00
Begin VB.Form Options 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Options"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "File"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Options.frx":0000
         Left            =   3360
         List            =   "Options.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Pinger"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sound"
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Color"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
