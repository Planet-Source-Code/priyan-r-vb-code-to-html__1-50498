VERSION 5.00
Begin VB.Form frmstatus 
   Caption         =   "Creating HTML Files"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1830
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmstatus.frx":0000
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Text1.Text = ""
Me.Text1.BackColor = Me.BackColor
End Sub

Private Sub Label1_Click()

End Sub
