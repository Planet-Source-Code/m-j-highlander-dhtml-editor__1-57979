Attribute VB_Name = "RegExp_Functions"
Option Explicit

Private Const Quote = """"
'Private Const ALL_SPECIAL_CHARS = "[\s" & Quote & "> !#$%&'\(\)\*\+,\-\./:;=\?@\[\]\^_`{\|}~\\]"

Public Enum HTML_REMOVAL_OPTIONS
    REMOVE_HEAD = 1
    REMOVE_TAIL = 2
    REMOVE_BOTH = 3
End Enum

Public Function RemovePath(ByVal Attr As String) As String
' INPUT:  Attribute with contents
' Output: Attribute with content after removing file path
' Example - HREF attribute: (Quotes are part of the strings)
' Input   href="file:///cool folder/help.html"
' Output  href="help.html"

Dim sTemp As String
Dim OpenQuote As Long, CloseQuote As Long, LastSlash As Long

CloseQuote = 0

sTemp = Attr
OpenQuote = InStr(1, sTemp, Quote)
CloseQuote = InStr(OpenQuote + 1, sTemp, Quote)

If CloseQuote <> 0 Then
    '''''Quotes Found, Handle Path
    sTemp = Mid$(sTemp, OpenQuote, CloseQuote - OpenQuote)
    sTemp = Replace$(Attr, "\", "/") ' just in case
    LastSlash = InStrRev(sTemp, "/")
    If LastSlash = 0 Then
        'No path info, do nothing.
        RemovePath = sTemp
    Else
        'Remove Path:
        sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash - 1)
        RemovePath = Left$(Attr, OpenQuote - 1) & Quote & sTemp & Quote
    End If

Else
    '''''Quotes NOT Found, do nothing
    RemovePath = sTemp
End If

'MsgBox Attr
'MsgBox RemovePath

End Function


Public Function TrimWhiteSpace(ByVal Text As String) As String
' Trims leading and trailing whitespace characters
' including Space, Tab, Cr , Lf , ...

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.MultiLine = False
objRegExp.IgnoreCase = False
objRegExp.Global = True
objRegExp.Pattern = "^[\s]*(.*?)[\s]*$"

TrimWhiteSpace = objRegExp.Replace(Text, "$1")

Set objRegExp = Nothing

End Function
Public Function RX_ReplaceSubMatch(ByVal Text As String, ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = True) As String
Dim SC As CStrCat
Dim m As Match
Dim objRegExp As RegExp

Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.Pattern = Pattern


For Each m In objRegExp.Execute(Text)
    
    Text = Replace(Text, m.Value, "x", 1, -1, vbTextCompare)
    
Next


Set objRegExp = Nothing


End Function
Public Function RX_GenericExtractToArray(ByVal Text As String, ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = True) As String()
Dim idx As Long
Dim sTempArray() As String
Dim m As Match
Dim mc As MatchCollection

Dim objRegExp As RegExp

Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.Pattern = Pattern


Set mc = objRegExp.Execute(Text)
If mc.Count <> 0 Then
    
    ReDim sTempArray(0 To mc.Count - 1)
    For idx = 0 To mc.Count - 1
        sTempArray(idx) = mc.Item(idx).Value
    Next
    
    RX_GenericExtractToArray = sTempArray
Else
    'no matches found

End If

Set objRegExp = Nothing


End Function
Public Function RX_ReplaceForInArray(ByVal Text As String) As String
'convert "FOR index IN array" syntax to "FOR index=0 TO LBOUND(array)"

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.MultiLine = True

objRegExp.Pattern = "[ \t]*for[ \t]+(\w+)[ \t]+in[ \t]+(\w+)[ \t]*\r\n"
'what pattern means:

RX_ReplaceForInArray = objRegExp.Replace(Text, "For $1=0 To ubound($2)" & vbCrLf)

Set objRegExp = Nothing

End Function
Public Function RX_ReplaceWriteFunction(ByVal Text As String) As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.MultiLine = True

objRegExp.Pattern = "^[ \t]*?write[ \t]+?(.*?)\r\n"


RX_ReplaceWriteFunction = objRegExp.Replace(Text, "writes $1" & vbCrLf)

Set objRegExp = Nothing

End Function

Public Function RX_GenericExtractSubMatch(ByVal Text As String, ByVal Pattern As String, Optional ByVal SubMatchIndex As Integer = 0, Optional ByVal IgnoreCase As Boolean = True) As String
Dim SC As CStrCat
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.Pattern = Pattern

SC.MaxLength = Len(Text)

For Each m In objRegExp.Execute(Text)
    SC.AddStr m.SubMatches(SubMatchIndex) & vbCrLf
Next

RX_GenericExtractSubMatch = SC.StrVal

Set objRegExp = Nothing
Set SC = Nothing

End Function
Public Function EscapeNonPrintableChars(ByVal sText As String) As String

sText = Replace$(sText, vbCrLf, "\r\n")
sText = Replace$(sText, vbTab, "\t")
EscapeNonPrintableChars = sText

End Function
Public Function RX_ReplaceTagKeepContent(ByVal Html As String, ByVal FindTag As String, ByVal ReplaceTagOpen As String, ByVal ReplaceTagClose As String) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

FindTag = Replace$(FindTag, "<", "")
FindTag = Trim$(Replace$(FindTag, ">", ""))
sOpenTag = "<" & FindTag
sCloseTag = "</" & FindTag & ">"

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True

'the pattern reads as follows:
'Find ">" e.g:   <B>
'OR: Find "at least one non-alpha char,
'    followed by any char and finally >"
objRegExp.Pattern = sOpenTag & "(>|[^a-z][^\v]*?>)"
Html = objRegExp.Replace(Html, ReplaceTagOpen)

objRegExp.Pattern = sCloseTag
Html = objRegExp.Replace(Html, ReplaceTagClose)

RX_ReplaceTagKeepContent = Html

End Function
Public Function RX_ReplaceTagAndContents(ByVal Html As String, ByVal Tag As String, ByVal ReplaceWith As String, Optional ByVal TagIsSingle As Boolean = True) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
If Not (TagIsSingle) Then
    sCloseTag = "</" & Tag & ">"
Else
    sCloseTag = ">"
End If

Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = sOpenTag & "(>|[^a-z][^\v]*?)" & sCloseTag
Html = objRegExp.Replace(Html, ReplaceWith)

RX_ReplaceTagAndContents = Html

End Function

Public Function RX_Test(ByVal Text As String, ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean

'Dim m As Match
Dim objRegExp As RegExp

Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = False  ' once is enough
objRegExp.MultiLine = True
objRegExp.Pattern = Pattern


RX_Test = objRegExp.Test(Text)

Set objRegExp = Nothing

End Function
Public Function RX_KeepHTMLBody(ByVal Html As String, ByVal OptRemove As HTML_REMOVAL_OPTIONS) As String
Dim rx As RegExp

Set rx = New RegExp
rx.Global = False
rx.IgnoreCase = True

Select Case OptRemove
    Case REMOVE_HEAD
        rx.Pattern = "<[^\v]*?<BODY.*?>"
        Html = rx.Replace(Html, "")
    Case REMOVE_TAIL
        'intentionaly greedy:
        rx.Pattern = "</body[^\v]*>"
        Html = rx.Replace(Html, "")
    Case REMOVE_BOTH
        rx.Pattern = "<[^\v]*?<BODY.*?>"
        Html = rx.Replace(Html, "")
        'intentionaly greedy:
        rx.Pattern = "</body[^\v]*>"
        Html = rx.Replace(Html, "")
End Select

RX_KeepHTMLBody = Html

Set rx = Nothing

End Function
Public Function RX_Check_Invalid_NewLine(ByVal Html As String) As Boolean
Dim bFlag1 As Boolean, bFlag2 As Boolean, bFlag3 As Boolean
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.MultiLine = True

' CR not followed by LF , or LF not preceded with CR
' this means either a single CR or a single LF or a LFCR couple

objRegExp.Pattern = "(\r[^\n])|([^\r]\n)"

RX_Check_Invalid_NewLine = objRegExp.Test(Html)

Set objRegExp = Nothing

End Function
Public Function RX_NormalizePre(ByVal Html As String) As String
Dim sTemp As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "<pre[^\v]*?>([^\v]*?)</pre>"

Dim m
For Each m In objRegExp.Execute(Html)
   'Add <BR> and convert >,<,",& to Entities
   'sTemp = AddBR(m.SubMatches(0), False)
   sTemp = m.SubMatches(0)
   sTemp = Replace$(sTemp, " ", Chr$(2))
   sTemp = Replace$(sTemp, vbTab, Chr$(3))
   Html = Replace$(Html, m.Value, sTemp)
Next

RX_NormalizePre = Html

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function EscapeRegExpChars(ByVal PlainText As String) As String
'Input:  plain text, NO RegExp special chars allowed
'Output: RegExp compatible text
Dim sTemp As String
' Chars that need Escaping:
' \^$*+{}?.:=!|[]-(),
sTemp = PlainText
sTemp = Replace$(sTemp, "\", "\\") 'this has to be first
sTemp = Replace$(sTemp, "^", "\^")
sTemp = Replace$(sTemp, "$", "\$")
sTemp = Replace$(sTemp, "*", "\*")
sTemp = Replace$(sTemp, "+", "\+")
sTemp = Replace$(sTemp, "{", "\{")
sTemp = Replace$(sTemp, "}", "\}")
sTemp = Replace$(sTemp, "?", "\?")
sTemp = Replace$(sTemp, ".", "\.")
sTemp = Replace$(sTemp, ":", "\:")
sTemp = Replace$(sTemp, "=", "\=")
sTemp = Replace$(sTemp, "!", "\!")
sTemp = Replace$(sTemp, "|", "\|")
sTemp = Replace$(sTemp, "[", "\[")
sTemp = Replace$(sTemp, "]", "\]")
sTemp = Replace$(sTemp, "-", "\-")
sTemp = Replace$(sTemp, "(", "\(")
sTemp = Replace$(sTemp, ")", "\)")
sTemp = Replace$(sTemp, ",", "\,")
EscapeRegExpChars = sTemp

End Function
Public Function RX_GenericReplace(ByVal Text As String, ByVal Pattern As String, ByVal ReplaceWith As String, Optional ByVal IgnoreCase As Boolean = True) As String

Dim objRegExp As RegExp

Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.MultiLine = True
objRegExp.Pattern = Pattern


RX_GenericReplace = objRegExp.Replace(Text, ReplaceWith)


Set objRegExp = Nothing

End Function
Public Function RX_ExtractHREFs(ByVal Html As String) As String
Dim SC As CStrCat
Dim sImgFile As String
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "< ?A[^\v]*?HREF=""([^\v]*?)""[^\v]*?>"

SC.MaxLength = Len(Html)
For Each m In objRegExp.Execute(Html)
    'MsgBox m.Value
    'MsgBox m.SubMatches(0)
    SC.AddStr m.SubMatches(0) & vbCrLf
Next

RX_ExtractHREFs = SC 'default value

'Overkill, it will go out of scope anyway.
Set SC = Nothing
Set objRegExp = Nothing

End Function
Public Function RX_RemoveMultipleSpaces(ByVal Text As String) As String
Dim RegEx As RegExp

Set RegEx = New RegExp
RegEx.Pattern = " {2,}"
RegEx.MultiLine = True
RegEx.Global = True

RX_RemoveMultipleSpaces = RegEx.Replace(Text, " ")

End Function
Public Function RX_GenericExtract(ByVal Text As String, ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = True) As String
Dim SC As CStrCat
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.Pattern = Pattern

SC.MaxLength = Len(Text)

For Each m In objRegExp.Execute(Text)
    SC.AddStr m.Value & vbCrLf
Next

RX_GenericExtract = SC 'default value


Set objRegExp = Nothing
Set SC = Nothing

End Function

Public Function RX_ExtractURLs(ByVal Html As String) As String
Dim sTemp As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "((ht|f)tp://w?w?w?\.?.*?\..*?)[""\s<>]"

Dim m
For Each m In objRegExp.Execute(Html)
    sTemp = sTemp & m.SubMatches(0) & vbCrLf
Next

RX_ExtractURLs = sTemp

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_ChangeFontSize(ByVal Html As String, ByVal NewFontSize As Byte) As String
Dim SIZE  As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

SIZE = Quote & CStr(NewFontSize) & Quote

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "<FONT[^\v]*?SIZE=(""?[1-7]""?)"

Dim m
For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
'    Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
Next


RX_ChangeFontSize = Html

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function


Public Function RX_ChangeFont(ByVal Html As String, ByVal NewFont As String) As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "FACE=""?[^ ]*""?" 'anything but spaces,quotes optional


NewFont = "FACE=" & Quote & NewFont & Quote
RX_ChangeFont = objRegExp.Replace(Html, NewFont)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function

Public Function RX_(ByVal Html As String) As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

'objRegExp.MultiLine = True
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = ""

'Dim m As Match
'For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
'   sTemp = sTemp & vbCrLf & m.Value
'Next

RX_ = objRegExp.Replace(Html, "")

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_CompactBlankLines(ByVal Text As String) As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.MultiLine = False
objRegExp.IgnoreCase = True
objRegExp.Global = True

'Remove Leading and Trailing blank lines
objRegExp.Pattern = "(^(\r\n){1,})|((\r\n){1,}$)"
Text = objRegExp.Replace(Text, "")

'Compact blank lines
objRegExp.Pattern = "(\r\n){3,}" ' match 3 or more CRLF
RX_CompactBlankLines = objRegExp.Replace(Text, vbCrLf & vbCrLf)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_RemoveBlankLines(ByVal Text As String) As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.MultiLine = False
objRegExp.Global = True

'Remove Leading and Trailing blank lines
objRegExp.Pattern = "(^(\r\n){1,})|((\r\n){1,}$)"
Text = objRegExp.Replace(Text, "")

'Remove blank lines
objRegExp.Pattern = "(\r\n){2,}" ' match 2 or more CRLF
RX_RemoveBlankLines = objRegExp.Replace(Text, vbCrLf)


'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_AddBR(ByVal Html As String) As String

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "\r\n" ' = CRLF

RX_AddBR = objRegExp.Replace(Html, "<br>" & vbCrLf)

'//OR, without RegExp (slightly slower):
'RX_AddBR = Replace(HTML, vbCrLf, "<BR>" & vbCrLf)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_ExtractTagWithContents(ByVal Html As String, ByVal Tag As String) As String
Dim sTemp As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
sCloseTag = "</" & Tag & ">"

Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = sOpenTag & "(>|[^a-z][^\v]*?)" & sCloseTag
Dim m
For Each m In objRegExp.Execute(Html)
    sTemp = sTemp & m.Value & vbCrLf
Next


RX_ExtractTagWithContents = sTemp

End Function

Public Function RX_RemoveTagWithContents(ByVal Html As String, ByVal Tag As String, Optional ByVal TagIsSingle As Boolean = True) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
If Not (TagIsSingle) Then
    sCloseTag = "</" & Tag & ">"
Else
    sCloseTag = ">"
End If

Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = sOpenTag & "(>|([^a-z][^\v]*?))" & sCloseTag
Html = objRegExp.Replace(Html, "")

RX_RemoveTagWithContents = Html

End Function
Public Function RX_RemoveCommentTagAndContent(ByVal Html As String) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Set objRegExp = New RegExp

sOpenTag = "<!"
sCloseTag = "->"

objRegExp.IgnoreCase = True
objRegExp.Global = True
'BOTH WORK FINE!
'objRegExp.Pattern = sOpenTag & "((>[^<\n\r])|\w|[""\n\r\t\.\(\)\[\]\+:\|&;/,@ =%{<}\?#'!/\-\*])*" & sCloseTag
objRegExp.Pattern = sOpenTag & "(([^\-]>)|\w|[""\n\r\t\.\(\)\[\]\+:\|&;/,@ =%{<}\?#'!/\-\*])*" & sCloseTag
'Dim m
'For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
'Next

Html = objRegExp.Replace(Html, "")

RX_RemoveCommentTagAndContent = Html

End Function
Public Function RX_RemoveAllTags(ByVal Html As String)

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "<([^\v]*?)>"
RX_RemoveAllTags = objRegExp.Replace(Html, "")
 
End Function
Public Function RX_RemoveTagKeepContent(ByVal Html As String, ByVal Tag As String) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
sCloseTag = "</" & Tag & ">"

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = sOpenTag & "(>|[^a-z][^\v]*?>)"
Html = objRegExp.Replace(Html, "")

objRegExp.Pattern = sCloseTag
Html = objRegExp.Replace(Html, "")

RX_RemoveTagKeepContent = Html

End Function

