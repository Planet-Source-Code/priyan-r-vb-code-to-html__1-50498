VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan's Vb Code TO HTML(With Highlighting)"
   ClientHeight    =   4125
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdbuild 
      Caption         =   "Create Html Files"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "About"
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
      Height          =   240
      Left            =   5520
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Files In the project"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Vb Project"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modules_count%, forms_count%, class_count%, usercontrol_count%
Option Explicit

Private Sub cmdbrowse_Click()
With Me.CommonDialog1
.Filter = "Vb Project(*.vbp)|*.vbp"
.filename = ""
.Flags = cdlOFNFileMustExist
.ShowOpen
If .filename <> "" Then
  Text1.Text = .filename
  'Search for forms,class,modules,user controls..
  searchfiles .filename
  cmdbuild.Enabled = True
End If
End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdbuild_Click()
If Me.ListView1.ListItems.Count = 0 Then
    MsgBox "No files to create html", vbCritical, App.title
    Exit Sub
End If
Dim fld$
fld = GetSetting("priyan", App.title, "lastdir", "")
fld = frmdirselect.showfolderdia(fld, "Select a folder to save HTML files")
If fld = "" Then
    MsgBox "User Canceld", vbCritical, App.title
    Exit Sub
End If
cmdbuild.Enabled = False
'!!Important!! change the cur dir to the path of the project
ChDir getdirname(Text1.Text)
SaveSetting "priyan", App.title, "lastdir", fld
Dim itm As ListItem, found As Boolean
Dim fno%, fno1%, str$, last$
frmstatus.Show , Me
Dim i%
For i = 1 To Me.ListView1.ListItems.Count
Set itm = Me.ListView1.ListItems(i)
DoEvents
    fno = FreeFile
    'Open the VB file
    Open itm.Text For Input As #fno
    frmstatus.Text1.Text = "Creating HTML file for " & itm.Text & vbCrLf & vbCrLf & "File " & i & " Of " & Me.ListView1.ListItems.Count
     fno1 = FreeFile
     'open the output HTML file
    Open addstrap(fld, itm.SubItems(1) & ".htm") For Output As #fno1
'    ==========================================
'    adds html code for the header of every html file
    Print #fno1, "<html><body><A href=""Index.htm"">Home</a>&nbsp;&nbsp;&nbsp;<font color=red>File : " & itm.Text & "&nbsp;&nbsp;" & "Type:" & itm.SubItems(2) & " Name: " & itm.SubItems(1) & "</font><br><hr>"
'    ==========================================
'Loop until the end of the vb file
    Do Until EOF(fno)
    DoEvents
    'read a line of code
         Line Input #1, str
         'Every Vb files form,module,calss,user control..
         'Have a format
         'The first lines of the files are attributes for VB IDE
         'Found become true when we reach the code part of the file
         'That is done by the following lines
         If found = True Then
            'Write HTMl code with the highlightcode() in module modhighlight
               Print #fno1, modhighlight.highlightcode(str) & "<br>"
               'A function or sub is finished we need to goto next line
                If InStr(1, str, "End Sub") <> 0 Or InStr(1, str, "End Function") <> 0 Then
                'Put a break tag
                    Print #fno1, "<br>"
                End If
            'End If
        End If
        If itm.SubItems(2) <> "Module" Then
        'If the file is not a module then last attribute in the file is Attribute VB_Exposed
        'The next line start with the code
            If InStr(1, str, "Attribute VB_Exposed") <> 0 Then
               found = True
            End If
        Else
        'If the file is nott a module then last attribute in the file is Attribute VB_Name
        'The next line start with the code
            If InStr(1, str, "Attribute VB_Name") <> 0 Then
               found = True
            End If
        End If
      
    Loop
    'Put a </html> tag
    Print #fno1, "</body></html>"
    'close the output file
    Close #fno1
    'close the input file
    Close #fno
    found = False
Next
frmstatus.Text1.Text = "Creating Index File( Index.htm )"
DoEvents
'===============================================
'Crate a index file arranged with links to all forms,class,user control..
'===============================================
createindexfile fld
frmstatus.Text1.Text = "Finished!"
'===============================================
'open the index file
ShellExecute 0, "open", addstrap(fld, "index.htm"), "", "", 1
MsgBox "HTML files created", vbInformation, App.title
cmdbuild.Enabled = True
Unload frmstatus
End Sub

Private Sub Form_Load()
Me.ListView1.ColumnHeaders.Add , , "File", Me.ListView1.Width / 2
Me.ListView1.ColumnHeaders.Add , , "VB Name", Me.ListView1.Width / 4
Me.ListView1.ColumnHeaders.Add , , "Type", Me.ListView1.Width / 5
modhighlight.initkeywords
End Sub
Public Sub searchfiles(ByVal projectfile$)
Me.ListView1.ListItems.Clear
forms_count = 0
modules_count = 0
class_count = 0
usercontrol_count = 0
Dim itm As ListItem, str As String, pos%
Dim file_type$, filename$, check As Boolean
Open projectfile For Input As #1 'Open the project file
Do Until EOF(1)
Input #1, str
'split each line wo parts
'speperacted by '=')

'Get the name property name
file_type = Trim(modcommon.extractstring(str, "=", 0))
'Get the name property value
filename = Trim(modcommon.extractstring(str, "=", 1))
'check the property
Select Case LCase(file_type)
    'ITS A FORM
    Case "form"
        check = True
    'ITS A MODULE
    Case "module"
        pos = InStr(1, filename, ";")
        filename = Trim(Mid(filename, pos + 1, Len(filename) - pos))
        check = True
    'ITS A CLASS
    Case "class"
       pos = InStr(1, filename, ";")
        filename = Trim(Mid(filename, pos + 1, Len(filename) - pos))
        check = True
    Case "usercontrol"
        check = True
    Case Else
        check = False
End Select
If check = True Then
    '!!Important!! change the cur dir to the path of the project
    ChDir getdirname(projectfile)
   If filexists(filename) Then
        Set itm = Me.ListView1.ListItems.Add(, , filename)
        'Get the name property of the  vb fiLE file
        itm.ListSubItems.Add , , Trim(getvbname(filename))
        itm.ListSubItems.Add , , file_type
'===========================================
'Increment the count variables
        Select Case LCase(file_type)
            Case "form"
                forms_count = forms_count + 1
            Case "module"
                modules_count = modules_count + 1
            Case "class"
                class_count = class_count + 1
            Case "usercontrol"
                usercontrol_count = usercontrol_count + 1
        End Select
'===========================================
    Else
        frmnotfound.List1.AddItem filename
   End If
End If
DoEvents

Loop
If frmnotfound.List1.ListCount <> 0 Then
    frmnotfound.Show , Me
Else
    Unload frmnotfound
End If
Close #1

End Sub
Public Function getvbname(ByVal filename$) As String
'===============================================
'Gets the name of the VB FIle
'==============================================='
'Gets the name of the vbfile(form,class,module...."
Dim fno%, str$, name_, atribute$
fno = FreeFile
Open filename For Input As fno
    Do Until EOF(fno)
        Input #fno, str
        'replace all double spaces with single
        str = Replace(str, Space(2), "")
        atribute = extractstring(str, "=", 0)
        name_ = extractstring(str, "=", 1)
        If atribute = "Attribute VB_Name " Then
            getvbname = Replace(name_, """", "")
            Exit Do
        End If
    Loop
Close #fno
End Function

Public Sub createindexfile(ByVal FOLDER$)
'======================================
'CREATES A FILE INDEX.HTM with links to each HTML files
'======================================
Dim fno1%
fno1 = FreeFile
Dim itm As ListItem
Open addstrap(FOLDER, "Index.htm") For Output As #fno1

    Print #fno1, "<html><font color=""blue"">Created by " & "Priyan's VbCode TO HTML(With Highlightig)" & "</font><br><br>"
    Print #fno1, "<font color=""red"">" & Text1.Text & "<font><br><hr>"
    Print #fno1, "<font color=""blue"">Forms : " & forms_count & "</font><br><hr><table width=100%><tr>"
    Dim iname$, itms%
    itms = 0
    'We put a five links in a row
    
    'Puts links for forms
    '=============================================
      For Each itm In Me.ListView1.ListItems
      DoEvents
        If itm.SubItems(2) = "Form" Then
            itms = itms + 1
            Print #fno1, "<td><a href=""" & itm.SubItems(1) & ".htm"">" & itm.SubItems(1) & "</a>&nbsp;&nbsp;&nbsp;</td>"
            If itms = 5 Then
                Print #fno1, "</tr><tr>"
                itms = 0
            End If
        End If
    Next
    Print #fno1, "</tr></table><Br><hr>"
    '=============================================
    Print #fno1, "<font color=""blue"">Modules: " & modules_count & "</font><br><hr><table width=100%><tr>"
    itms = 0
    For Each itm In Me.ListView1.ListItems
    DoEvents
        If itm.SubItems(2) = "Module" Then
            itms = itms + 1
            Print #fno1, "<td><a href=""" & itm.SubItems(1) & ".htm"">" & itm.SubItems(1) & "</a>&nbsp;&nbsp;&nbsp;</td>"
            If itms = 5 Then
                Print #fno1, "</tr><tr>"
                itms = 0
            End If
        End If
    Next
    Print #fno1, "</tr></table><hr>"
    '=============================================
      '=============================================
    Print #fno1, "<font color=""blue"">Classes : " & class_count & "</font><br><hr><table width=100%><tr>"
    itms = 0
    For Each itm In Me.ListView1.ListItems
    DoEvents
        If itm.SubItems(2) = "Class" Then
                  itms = itms + 1
            Print #fno1, "<td><a href=""" & itm.SubItems(1) & ".htm"">" & itm.SubItems(1) & "</a>&nbsp;&nbsp;&nbsp;</td>"
            If itms = 5 Then
                Print #fno1, "</tr><tr>"
                itms = 0
            End If
      
        End If
    Next
    Print #fno1, "</tr></table><hr>"
    '=============================================
      '=============================================
    Print #fno1, "<font color=""blue"">User Controls: " & class_count & "</font><br><hr><table width=100%><tr>"
    itms = 0
    For Each itm In Me.ListView1.ListItems
    DoEvents
        If itm.SubItems(2) = "UserControl" Then
            itms = itms + 1
            Print #fno1, "<td><a href=""" & itm.SubItems(1) & ".htm"">" & itm.SubItems(1) & "</a>&nbsp;&nbsp;&nbsp;</td>"
            If itms = 5 Then
                Print #fno1, "</tr><tr>"
                itms = 0
            End If
        End If
    Next
    Print #fno1, "</tr></table>"
    '=============================================
      Print #fno1, "</html>"
    
    '
Close #fno1
End Sub

Private Sub Label3_Click()
frmabout.Show
End Sub
Public Function filexists(ByVal file$) As Boolean
On Error GoTo ext:
If Dir(file, vbNormal) <> "" Then
    filexists = True
End If
Exit Function
ext:
End Function
