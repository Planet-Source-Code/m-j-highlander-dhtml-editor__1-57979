VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "cmax20.ocx"
Begin VB.Form frmEditSource 
   Caption         =   "Edit HTML Source"
   ClientHeight    =   9795
   ClientLeft      =   465
   ClientTop       =   765
   ClientWidth     =   14220
   Icon            =   "ViewSource.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   14220
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   1349
      ButtonWidth     =   1032
      ButtonHeight    =   1191
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "OK"
            Key             =   "ok"
            Object.Tag             =   ""
            ImageKey        =   "ok"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cancel"
            Key             =   "cancel"
            Object.Tag             =   ""
            ImageKey        =   "cancel"
         EndProperty
      EndProperty
   End
   Begin CodeMaxCtl.CodeMax CodeMax1 
      Height          =   5595
      Left            =   960
      OleObjectBlob   =   "ViewSource.frx":000C
      TabIndex        =   0
      Top             =   930
      Width           =   7005
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":0176
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":0398
            Key             =   "info"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":0AA6
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":11B4
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":18C2
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":1AE4
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewSource.frx":1D06
            Key             =   "clear"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ms_HtmlSource As String

Private ms_BaseURL As String

Private Function CodeMax1_KeyUp(ByVal Control As CodeMaxCtl.ICodeMax, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean

If KeyCode = vbKeyEscape Then
        ' Cancel
    If CodeMax1.Modified Then
        MainForm.DHTMLEdit1.DocumentHTML = ms_HtmlSource
    End If

    'Restore BaseURL:
    MainForm.DHTMLEdit1.BaseURL = ms_BaseURL
    Unload Me
End If


End Function
Private Sub Form_Load()

frmEditSource.CodeMax1.Font.Name = "Courier New"
frmEditSource.CodeMax1.Font.Size = 13

frmEditSource.CodeMax1.Text = MainForm.DHTMLEdit1.DocumentHTML
ms_HtmlSource = frmEditSource.CodeMax1.Text

Me.WindowState = GetSetting("DHTML-Edit", "EditHTMLSource", "Window-Size", 0)

'Setting the DocumentHTML property clears the BaseURL property, so we save it:
ms_BaseURL = MainForm.DHTMLEdit1.BaseURL

Me.Icon = MainForm.Icon

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


If UnloadMode = vbFormControlMenu Then 'clicked the X button, just hide
    
     MainForm.DHTMLEdit1.DocumentHTML = ms_HtmlSource

Else
     SaveSetting "DHTML-Edit", "EditHTMLSource", "Window-Size", CStr(Me.WindowState)

End If

End Sub
Private Sub Form_Resize()
On Error Resume Next

CodeMax1.Left = 0
CodeMax1.Top = Me.Toolbar1.Height
CodeMax1.Width = Me.ScaleWidth
CodeMax1.Height = Me.ScaleHeight - Me.Toolbar1.Height

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim vTemp As Variant

Select Case Button.Key

    Case "cancel"
        If CodeMax1.Modified Then
            MainForm.DHTMLEdit1.DocumentHTML = ms_HtmlSource
        End If

    Case "ok"
        MainForm.DHTMLEdit1.DocumentHTML = frmEditSource.CodeMax1.Text
        'Give the lady enought time to finish...
        Do While MainForm.DHTMLEdit1.Busy
            DoEvents
        Loop

        g_IsDirty = True

    Case Else
    
End Select


'Restore BaseURL:
MainForm.DHTMLEdit1.BaseURL = ms_BaseURL

Unload Me

'not-needed
'MainForm.DHTMLEdit1.DOM.ExecCommand "refresh"

End Sub
