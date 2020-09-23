Attribute VB_Name = "modhighlight"
Dim keywords As New Collection
Option Explicit


'=============================================
'Priyan's Vb Code to HTMl
'---------------------------
'add keywords i missed to the initkeywords() in this module
'
'=============================================
Public Function highlightcode(ByVal str$) As String
Dim words() As String, obj, comment$, pos%
Dim ret$
If Left(LTrim(str), 1) = "'" Then
        highlightcode = "<font color=green>" & str & "</font>"
Else
 'split a code line into parts
     pos = InStr(1, str, "'") 'Find if there is any comment in the line
    If pos <> 0 Then
        'Put the comment into comment var
        comment = "'" & Mid(str, pos + 1, Len(str) - pos)
        str = Left(str, pos - 1)
    End If
    words = Split(str, Space(1))
    For Each obj In words
    'check weather that string is in the keyword list
        If iskeyword(obj) Then '
        'adds <font> tags in html to highlight the text
            ret = ret & "<font color=blue>" & obj & "</font>&nbsp;"
        Else
            ret = ret & obj & "&nbsp;"
    
        End If
    Next
    If comment <> "" Then 'A comment in the line
        ret = ret & "<font color=""green"">" & comment & "</font>"
    End If
    highlightcode = ret
    
End If
End Function
Public Sub initkeywords()
'====================================================
'add keywords i missed to here
'eg: if you want to heighlight 'option explicit'
'you have to right to lines
'keywords.Add "option"
'keywords.Add "explicit"
'====================================================
keywords.Add "open"
keywords.Add "for"
keywords.Add "next"
keywords.Add "each"
keywords.Add "in"
keywords.Add "open"
keywords.Add "to"
keywords.Add "if"
keywords.Add "then"
keywords.Add "elseif"
keywords.Add "else"
keywords.Add "end if"
keywords.Add "public"
keywords.Add "private"
keywords.Add "function"
keywords.Add "sub"
keywords.Add "byval"
keywords.Add "as"
keywords.Add "string"
keywords.Add "integer"
keywords.Add "boolean"
keywords.Add "long"
keywords.Add "byte"
keywords.Add "double"
keywords.Add "Variant"
keywords.Add "end"
keywords.Add "Option"
keywords.Add "explicit"
keywords.Add "with"
keywords.Add "dim"
keywords.Add "exit"
keywords.Add "do"
keywords.Add "loop"
keywords.Add "until"
keywords.Add "while"
keywords.Add "Declare"
keywords.Add "property"
keywords.Add "output"
keywords.Add "input"
keywords.Add "write"
keywords.Add "get"
keywords.Add "put"
keywords.Add "true"
keywords.Add "false"
keywords.Add "select"
keywords.Add "case"
keywords.Add "optional"
keywords.Add "compare"
keywords.Add "text"
keywords.Add "on"
keywords.Add "error"
keywords.Add "resume"
keywords.Add "goto"
keywords.Add "const"
keywords.Add "enum"
keywords.Add "type"
keywords.Add "global"
End Sub
Public Function iskeyword(ByVal str$) As Boolean
Dim obj
For Each obj In keywords
    If LCase(str) = LCase(obj) Then
        iskeyword = True
        Exit Function
    End If
Next
End Function
