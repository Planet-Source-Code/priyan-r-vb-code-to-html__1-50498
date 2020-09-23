VERSION 5.00
Begin VB.Form frmdirselect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Folder"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdnewfolder 
      Caption         =   "New Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmdirselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()
modcommon.tempvar = Dir1.Path
Unload Me
End Sub

Private Sub Command1_Click()
modcommon.tempvar = ""
Unload Me

End Sub

Private Sub cmdnewfolder_Click()
Dim fld$
fld = InputBox("Enter Folder Name", "Create Folder", "VB HTML")
On Error Resume Next
If Dir(addstrap(Dir1.Path, fld), vbDirectory) = "" Then
    Err.Clear
    MkDir addstrap(Dir1.Path, fld)
    If Err Then
        MsgBox "Can't Create Folder", vbCr, App.title
        Exit Sub
    End If
End If
Dir1.Path = addstrap(Dir1.Path, fld)
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Drive1_Change()
On Error GoTo ext:
Dim dr$
Dir1.Path = Drive1.Drive
Dir1.Refresh
Exit Sub
ext:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Load()
'Me.Icon = MDIForm1.Icon
Drive1.Drive = "C:\"
Dir1.Path = "c:\"
End Sub
Public Function showfolderdia(Optional defaultdir$, Optional title$, Optional allowcreatefolder As Boolean = True) As String
If title <> "" Then Me.Caption = title
If allowcreatefolder = False Then Me.cmdnewfolder.Enabled = False
On Error Resume Next
If defaultdir <> "" Then
    If Dir(defaultdir, vbDirectory) <> "" Then
        Dir1.Path = defaultdir
    End If
End If
Me.Show vbModal
showfolderdia = modcommon.tempvar
End Function
