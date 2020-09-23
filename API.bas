Attribute VB_Name = "API_Functions"
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long




Public Type ColorVals
    ColorName As String
    ColorLong As Long
    ColorHex As String
End Type

Public ga_ColorVals(0 To 15) As ColorVals

Private Const MF_BITMAP = &H4&
Public Sub InitColors()

'ga_ColorVals().ColorName=
'ga_ColorVals().ColorHex=
'ga_ColorVals().ColorLong=

ga_ColorVals(1).ColorName = "Black"
ga_ColorVals(1).ColorHex = "#000000"
ga_ColorVals(1).ColorLong = &H0&

ga_ColorVals(12).ColorName = "Silver"
ga_ColorVals(12).ColorHex = "#C0C0C0"
ga_ColorVals(12).ColorLong = &HC0C0C0

ga_ColorVals(4).ColorName = "Gray"
ga_ColorVals(4).ColorHex = "#808080"
ga_ColorVals(4).ColorLong = &H808080

ga_ColorVals(14).ColorName = "White"
ga_ColorVals(14).ColorHex = "#FFFFFF"
ga_ColorVals(14).ColorLong = &HFFFFFF

ga_ColorVals(7).ColorName = "Maroon"
ga_ColorVals(7).ColorHex = "#800000"
ga_ColorVals(7).ColorLong = &H80&

ga_ColorVals(11).ColorName = "Red"
ga_ColorVals(11).ColorHex = "#FF0000"
ga_ColorVals(11).ColorLong = &HFF&

ga_ColorVals(10).ColorName = "Purple"
ga_ColorVals(10).ColorHex = "#800080"
ga_ColorVals(10).ColorLong = &H800080

ga_ColorVals(3).ColorName = "Fuchsia"
ga_ColorVals(3).ColorHex = "#FF00FF "
ga_ColorVals(3).ColorLong = &HFF00FF

ga_ColorVals(5).ColorName = "Green"
ga_ColorVals(5).ColorHex = "#008000"
ga_ColorVals(5).ColorLong = &H8000&

ga_ColorVals(6).ColorName = "Lime"
ga_ColorVals(6).ColorHex = "#00FF00"
ga_ColorVals(6).ColorLong = &HFF00&

ga_ColorVals(9).ColorName = "Olive"
ga_ColorVals(9).ColorHex = "#808000"
ga_ColorVals(9).ColorLong = &H8080&

ga_ColorVals(15).ColorName = "Yellow"
ga_ColorVals(15).ColorHex = "#FFFF00"
ga_ColorVals(15).ColorLong = &HFFFF&

ga_ColorVals(8).ColorName = "Navy"
ga_ColorVals(8).ColorHex = "#000080"
ga_ColorVals(8).ColorLong = &H800000

ga_ColorVals(2).ColorName = "Blue"
ga_ColorVals(2).ColorHex = "#0000FF"
ga_ColorVals(2).ColorLong = &HFF0000

ga_ColorVals(13).ColorName = "Teal"
ga_ColorVals(13).ColorHex = "#008080"
ga_ColorVals(13).ColorLong = &H808000

ga_ColorVals(0).ColorName = "Aqua"
ga_ColorVals(0).ColorHex = "#00FFFF"
ga_ColorVals(0).ColorLong = &HFFFF00


End Sub


Function RevRGB(ByVal VBHexRGB As String) As String
' VB generated Hex RGB must be reversed to be used in HTML

Dim var1 As String
Dim var2 As String
Dim Var3 As String

var1 = Left$(VBHexRGB, 2)
var2 = Mid$(VBHexRGB, 3, 2)
Var3 = Right$(VBHexRGB, 2)

RevRGB = Var3 & var2 & var1

End Function


Public Function ColorToHex(ByVal lColor As Long) As String
Dim sTemp As String

sTemp = Hex$(lColor)

If Len(sTemp) < 6 Then sTemp = String$(6 - Len(sTemp), "0") + sTemp
sTemp = "#" & RevRGB(sTemp)

ColorToHex = sTemp

End Function

Public Sub SetMenuIcon(hwnd As Long, MenuIndex As Long, SubIndex As Long, pic As Picture)
Dim hMenu As Long, hSubMenu As Long, hID As Long

'Get the menuhandle of the form
hMenu = GetMenu(hwnd)

'Get the handle of the first submenu
hSubMenu = GetSubMenu(hMenu, MenuIndex)

'Get the menuId of the first entry
hID = GetMenuItemID(hSubMenu, SubIndex)

'Add the bitmap
SetMenuItemBitmaps hMenu, hID, MF_BITMAP, pic, pic

End Sub

