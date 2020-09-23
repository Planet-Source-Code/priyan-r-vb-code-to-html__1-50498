VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblvote 
      AutoSize        =   -1  'True
      Caption         =   "Vote This Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1440
      Width           =   1620
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Mail Me At"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   750
   End
   Begin VB.Label Label5 
      Caption         =   "Visit Me at"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblmail 
      AutoSize        =   -1  'True
      Caption         =   "vb@priyan.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Priyan R"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblsite 
      Caption         =   "http://www.priyan.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      MouseIcon       =   "frmabout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "I think this  will be useful to you . If you like this program please vote for me (Click The Link Below)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   120
      Picture         =   "frmabout.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







Private Sub Form_Load()
lblmail.MouseIcon = lblsite.MouseIcon
lblvote.MouseIcon = lblsite.MouseIcon
End Sub

Private Sub lblmail_Click()
ShellExecute 0, "open", "mailto:" & lblmail.Caption & "?subject=vb:" & App.title, "", "", 1
End Sub

Private Sub lblsite_Click()
ShellExecute 0, "open", lblsite.Caption, "", "", 1
End Sub

Private Sub lblvote_Click()
modcommon.vote

End Sub
