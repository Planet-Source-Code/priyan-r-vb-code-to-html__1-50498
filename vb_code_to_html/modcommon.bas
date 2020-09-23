Attribute VB_Name = "modcommon"
Public tempvar As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Public Function extractstring(ByVal str$, ByVal cmp$, ByVal no%) As String
Dim arr() As String
arr = Split(str, cmp)
If no <= UBound(arr) Then
    extractstring = arr(no)
Else
    extractstring = ""
End If

End Function
Public Function getdirname(file$) As String
Dim pos%
getdirname = StrReverse(file)
pos = InStr(1, getdirname, "\")
getdirname = Mid(getdirname, pos + 1, Len(file) - pos)
getdirname = StrReverse(getdirname)
End Function

Public Function addstrap(ByVal path1 As String, ByVal path2 As String) As String
If Right$(path1, 1) = "\" Then
     addstrap = path1 & path2
Else
         addstrap = path1 & "\" & path2
End If
End Function

Public Sub vote()
Const url = "http://www.websamba.com/priyanr/pscredirect/redir.asp?appname="
'Const url = "http://priyan/home/pscredirect/redir.asp?appname="
Const appid = "Vb Code TO HTML(With Highlighting)"
ShellExecute 0, "open", url & appid, "", "", 1
End Sub
Public Function getfilename(file$) As String
Dim pos%
getfilename = StrReverse(file)
pos = InStr(1, getfilename, "\")
getfilename = Left(getfilename, pos - 1)
getfilename = StrReverse(getfilename)
End Function

