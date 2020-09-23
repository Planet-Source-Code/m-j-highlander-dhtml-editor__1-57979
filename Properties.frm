VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Properties"
   ClientHeight    =   4050
   ClientLeft      =   2505
   ClientTop       =   3450
   ClientWidth     =   6660
   ControlBox      =   0   'False
   Icon            =   "Properties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6660
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTopMargin 
      Height          =   285
      Left            =   1710
      TabIndex        =   15
      Top             =   3450
      Width           =   525
   End
   Begin VB.TextBox txtLeftMargin 
      Height          =   285
      Left            =   1710
      TabIndex        =   13
      Top             =   3060
      Width           =   525
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1710
      TabIndex        =   11
      Top             =   2550
      Width           =   4665
   End
   Begin VB.Frame Frame1 
      Caption         =   " Preview "
      Height          =   1665
      Left            =   3000
      TabIndex        =   16
      Top             =   390
      Width           =   3555
      Begin VB.PictureBox picBackground 
         Height          =   1275
         Left            =   150
         ScaleHeight     =   1215
         ScaleWidth      =   3195
         TabIndex        =   21
         Top             =   270
         Width           =   3255
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A visited link: "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   750
            Width           =   1230
         End
         Begin VB.Label lblVisited 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "www.coolsite.com"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   18
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label lblHyperlink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "www.somesite.net"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   540
            TabIndex        =   17
            Top             =   450
            Width           =   1560
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "This is some document text, with default color..., now check out this link:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   90
            TabIndex        =   22
            Top             =   60
            Width           =   2955
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5280
      TabIndex        =   20
      Top             =   3570
      Width           =   1250
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5280
      TabIndex        =   19
      Top             =   3060
      Width           =   1250
   End
   Begin VB.CheckBox chkColors 
      Height          =   350
      Index           =   4
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1980
      Width           =   1000
   End
   Begin VB.CheckBox chkColors 
      Height          =   350
      Index           =   3
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1000
   End
   Begin VB.CheckBox chkColors 
      Height          =   350
      Index           =   2
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1140
      Width           =   1000
   End
   Begin VB.CheckBox chkColors 
      Height          =   350
      Index           =   1
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1000
   End
   Begin VB.CheckBox chkColors 
      Height          =   350
      Index           =   0
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   300
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog cdlgColor 
      Left            =   60
      Top             =   3510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTopMargin 
      AutoSize        =   -1  'True
      Caption         =   "&Top Margin"
      Height          =   195
      Left            =   810
      TabIndex        =   14
      Top             =   3510
      Width           =   810
   End
   Begin VB.Label lblLeftMargin 
      AutoSize        =   -1  'True
      Caption         =   "&Left Margin"
      Height          =   195
      Left            =   825
      TabIndex        =   12
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "&Document Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   330
      TabIndex        =   10
      Top             =   2580
      Width           =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "&Active hyperlink"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Visited hyperlink"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   1635
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Hyperlink"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   810
      TabIndex        =   4
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Text"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1245
      TabIndex        =   2
      Top             =   765
      Width           =   375
   End
   Begin VB.Label lblBackground 
      AutoSize        =   -1  'True
      Caption         =   "&Background"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdatePreview()
    
    'update preview
    picBackground.BackColor = chkColors(0).BackColor
    lblText.ForeColor = chkColors(1).BackColor
    lblHyperlink.ForeColor = chkColors(2).BackColor
    lblVisited.ForeColor = chkColors(3).BackColor

End Sub
Private Sub chkColors_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeySpace Then
    chkColors(Index).Value = vbUnchecked
Else
    chkColors(Index).Value = vbChecked
    Call chkColors_MouseUp(Index, vbLeftButton, 0, 0, 0)
End If

End Sub

Private Sub chkColors_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button <> vbLeftButton Then Exit Sub

cdlgColor.Color = chkColors(Index).BackColor
cdlgColor.ShowColor

If Err Then
    Err.Clear
Else
    chkColors(Index).BackColor = cdlgColor.Color
    UpdatePreview

End If

chkColors(Index).Value = vbUnchecked

End Sub
Private Sub CmdCancel_Click()

Unload Me

End Sub
Private Sub cmdOK_Click()
Dim m As MSHTML.HTMLBody

MainForm.DHTMLEdit1.DOM.bgColor = FormatRGBString(chkColors(0).BackColor)
MainForm.DHTMLEdit1.DOM.fgColor = FormatRGBString(chkColors(1).BackColor)
MainForm.DHTMLEdit1.DOM.linkColor = FormatRGBString(chkColors(2).BackColor)
MainForm.DHTMLEdit1.DOM.vlinkColor = FormatRGBString(chkColors(3).BackColor)
MainForm.DHTMLEdit1.DOM.alinkColor = FormatRGBString(chkColors(4).BackColor)
MainForm.DHTMLEdit1.DOM.Title = txtTitle.Text

Set m = MainForm.DHTMLEdit1.DOM.body

m.leftMargin = txtLeftMargin.Text
m.topMargin = txtTopMargin.Text

Set m = Nothing

Unload Me

End Sub
Private Sub Form_Activate()
Dim m As MSHTML.HTMLBody
Set m = MainForm.DHTMLEdit1.DOM.body

chkColors(0).BackColor = UnFormatRGBString(MainForm.DHTMLEdit1.DOM.bgColor)
chkColors(1).BackColor = UnFormatRGBString(MainForm.DHTMLEdit1.DOM.fgColor)
chkColors(2).BackColor = UnFormatRGBString(MainForm.DHTMLEdit1.DOM.linkColor)
chkColors(3).BackColor = UnFormatRGBString(MainForm.DHTMLEdit1.DOM.vlinkColor)
chkColors(4).BackColor = UnFormatRGBString(MainForm.DHTMLEdit1.DOM.alinkColor)

txtTitle.Text = MainForm.DHTMLEdit1.DocumentTitle


txtLeftMargin.Text = m.leftMargin
txtTopMargin.Text = m.topMargin

UpdatePreview

Set m = Nothing

End Sub
Private Sub Form_Load()

cdlgColor.Color = 0
cdlgColor.CancelError = True
cdlgColor.Flags = cdlCCFullOpen Or cdlCCRGBInit

End Sub
