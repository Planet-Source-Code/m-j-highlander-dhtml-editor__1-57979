Attribute VB_Name = "Generic_Functions"
Option Explicit

'required for isDirty work-around!
Public g_IsDirty As Boolean

'Init CpmCtl32.dll ver 6 for WinXP styles
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Sub ColorizeCode(rtfHTML As RichTextBox)


Dim needle1, needle2, colorvar As Variant
Dim holder, difference1, difference2, start, startpos, endpos, between, check, goo As Long


holder = rtfHTML.SelStart   'Remember Where Cursor Is
rtfHTML.SelStart = 0        ' Make everything black
rtfHTML.SelLength = 90000   '
rtfHTML.SelColor = vbBlack  '
redo:
        Select Case goo
        
        Case Is = 0
            needle1 = "<"     ''1st Search String                               \
            needle2 = ">"     ''2nd Search String                               |
            colorvar = vbBlue ''Color To HighLight With                         |Color Tags
            difference1 = 1   ''# of spaces to add to the end of selection set  |
            difference2 = 0   ''Skip this many spaces from initial find         /
            
        Case Is = 1
            needle1 = Chr(61) & Chr(34)  ''\<--61 is a = and 34 is a "
            needle2 = Chr(34)            ''|
            colorvar = RGB(100, 100, 100)           ''|Color Variables
            difference1 = 0              ''|
            difference2 = 1              ''/
        
        Case Is = 2
            needle1 = "<!--"             ''\
            needle2 = "-->"              ''|
            colorvar = RGB(60, 114, 0)   ''|Color Comments
            difference1 = 3              ''|
            difference2 = 0              ''/
        
        Case Is = 3
            needle1 = "<SCRIPT"          ''\
            needle2 = "</SCRIPT>"        ''|
            colorvar = RGB(255, 140, 10) ''|Color Scripting
            difference1 = 9              ''|
            difference2 = 0              ''/
            
        Case Is = 4
            GoTo final                  ''Skip Coloring Process
    End Select
        

        Do Until 1 = 2
            If check = 0 Then        ''\
                start = 0            ''|
                check = 1            ''|Ensure the search starts from the beginning
                Else                 ''|
                start = startpos + 1 ''|
            End If                   ''/
                
            startpos = rtfHTML.Find(needle1, start) ''Find Begin Tag
                
            If startpos = -1 Then  ''\
            GoTo ender             ''|Check to see if it wasn't found
            End If                 ''/
            
            endpos = rtfHTML.Find(needle2, (startpos + 2)) '' Find End Tag
            
            between = endpos - startpos          ''Find space between needles
            
            rtfHTML.SelStart = (startpos + difference2)     ''\
            rtfHTML.SelLength = (between + difference1)     ''|Select and color the code
            rtfHTML.SelColor = colorvar                     ''/
            
        Loop
        
ender:
        goo = goo + 1 ''Advance to next coloring step
        GoTo redo:    ''Restart Coloring Process

final:
rtfHTML.SelStart = holder ''Return to where the cursor was before color code


End Sub
Public Sub LogErrors(ByVal ErrNumber As Long, ErrDescription As String, ErrSource As String)
Dim ff As Integer, sErrorInfo As String

ff = FreeFile


sErrorInfo = "TIME = " & Format(Now, "dd/mm/yyyy - hh:mm") & vbCrLf
sErrorInfo = sErrorInfo & "ERROR NUMBER = " & ErrNumber & vbCrLf
sErrorInfo = sErrorInfo & "ERROR DESCRIPTION = " & ErrDescription & vbCrLf
sErrorInfo = sErrorInfo & "ERROR SOURCE = " & ErrSource & vbCrLf

Open App.Path & "\error.log" For Output As #ff
    Print #ff, sErrorInfo
    Print #ff, "---------------------------------------------------------------------------------"
Close #ff


End Sub
Function RemoveSlash(ByVal sPath As String) As String

sPath = Trim$(sPath)

If Right$(sPath, 1) = "\" Then
    RemoveSlash = Left(sPath, Len(sPath) - 1)
Else
    RemoveSlash = sPath
End If


End Function


Function ChangeFileExtension(ByVal FileName As String, ByVal NewExtension As String) As String
Dim sOldExt As String, sDot As String

sOldExt = ExtractFileExtension(FileName)
If sOldExt = "" Then
    sDot = "."
Else
    sDot = ""
End If

ChangeFileExtension = Left$(FileName, Len(FileName) - Len(sOldExt)) & sDot & NewExtension

End Function
Function ExtractFileExtension(ByVal FileName As String) As String

Dim ThePos As Integer

'In case the path contains a dot
FileName = ExtractFileName(FileName)

ThePos = InStrRev(FileName, ".")
If ThePos = 0 Then
    ExtractFileExtension = ""
Else
    ExtractFileExtension = Right$(FileName, Len(FileName) - ThePos)
End If


End Function
Public Function ExtractFileName(ByVal FilePath As String) As String

' Extract the File name from a full file path

Dim iLastSlash As Integer

iLastSlash = InStrRev(FilePath, "\")

If iLastSlash = 0 Then
        ExtractFileName = FilePath
Else
    ExtractFileName = Right(FilePath, Len(FilePath) - iLastSlash)
End If


End Function
Public Function GetLinkHref(ByVal LinkHtml As String) As String
Dim iPos  As Integer
Dim sTemp As String



iPos = InStr(1, LinkHtml, "href", vbTextCompare)

If iPos > 0 Then
    
    sTemp = Right(LinkHtml, Len(LinkHtml) - iPos - 4)
    iPos = InStr(2, sTemp, """")
    If iPos <> 0 Then sTemp = Left(sTemp, iPos)
    GetLinkHref = Replace(sTemp, """", "")

Else
    'no href, could be an Anchor
    GetLinkHref = ""

End If




End Function
Public Function UnFormatRGBString(ByVal Color As String) As Long
Dim lColor As Long
Dim r As String
Dim g As String
Dim b As String
    
    
 Color = Right(Color, 6)

 r = Left(Color, 2)
 g = Mid(Color, 3, 2)
 b = Right(Color, 2)

 Color = "&h" & b & g & r
 
 UnFormatRGBString = CLng(Color)

End Function
Public Function FormatRGBString(val As Long) As String
    Dim Color As String
    Dim pad As Long
    Dim r As String
    Dim g As String
    Dim b As String
    
    ' This function formats a long consisting of rgb values
    ' taken from the CommonDialog color dialog
    ' to a string in the form of "#RRGGBB" where RRGGBB are
    ' hex values
    
    ' convert to hex
    Color = Hex(val)
    'determine how many zeros to pad in front of converted value
    pad = 6 - Len(Color)
    
    If pad Then
        Color = String(pad, "0") & Color
    End If
        
    'Extract the rgb components
    r = Right(Color, 2)
    g = Mid(Color, 3, 2)
    b = Left(Color, 2)
    
    ' Swab r and b position, color dialog returns
    ' bgr instead of rgb
    Color = "#" & r & g & b
    
    FormatRGBString = Color

End Function
