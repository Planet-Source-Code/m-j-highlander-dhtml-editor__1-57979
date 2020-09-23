VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   ClientHeight    =   5715
   ClientLeft      =   2295
   ClientTop       =   3465
   ClientWidth     =   9090
   Icon            =   "vbemform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9090
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "props"
            Object.ToolTipText     =   "Properties"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "html"
            Object.ToolTipText     =   "Edit HTML Source"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "revert"
            Object.ToolTipText     =   "Revert"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "redo"
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "font"
            Description     =   "Font"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1800
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "FontSize"
            Description     =   "FontSize"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1000
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   "Bold"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            Object.Tag             =   "Italic"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   "Underline"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Color"
            Description     =   "Color"
            Object.ToolTipText     =   "Color"
            Object.Tag             =   "Color"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BackColor"
            Object.ToolTipText     =   "Back Color"
            Object.Tag             =   "Back Color"
            ImageKey        =   "cpallette"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Numbers"
            Description     =   "Numbers"
            Object.ToolTipText     =   "Numbers"
            Object.Tag             =   "Numbers"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bullets"
            Description     =   "Bullets"
            Object.ToolTipText     =   "Bullets"
            Object.Tag             =   "Bullets"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Outdent"
            Description     =   "Outdent"
            Object.ToolTipText     =   "Outdent"
            Object.Tag             =   "Outdent"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Indent"
            Description     =   "Indent"
            Object.ToolTipText     =   "Indent"
            Object.Tag             =   "Indent"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LeftJustify"
            Description     =   "Leftv Justify"
            Object.ToolTipText     =   "LeftJustify"
            Object.Tag             =   "LeftJustify"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Center"
            Description     =   "Center"
            Object.ToolTipText     =   "Center"
            Object.Tag             =   "Center"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RightJustify"
            Description     =   "Right Justify"
            Object.ToolTipText     =   "RightJustify"
            Object.Tag             =   "RightJustify"
            ImageIndex      =   12
         EndProperty
      EndProperty
      Begin VB.ComboBox FontCombo 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "FontCombo"
         Top             =   30
         Width           =   2055
      End
      Begin VB.ComboBox FontSizeCombo 
         Height          =   315
         Left            =   2130
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   30
         Width           =   615
      End
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   3375
      Left            =   1050
      TabIndex        =   0
      Top             =   1020
      Width           =   6315
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   -1  'True
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   0   'False
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   5400
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Block Formatting"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   14102
            MinWidth        =   14111
            TextSave        =   ""
            Key             =   "mousemove"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   60
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12434877
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":0E42
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":0F4C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1056
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1160
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":126A
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1374
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":147E
            Key             =   "revert"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1588
            Key             =   "props"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1692
            Key             =   "html"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":19B4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1ABE
            Key             =   "new"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":1BC8
            Key             =   "format"
            Object.Tag             =   "format"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":210A
            Key             =   "bold"
            Object.Tag             =   "bold"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":264C
            Key             =   "italic"
            Object.Tag             =   "italic"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":2B8E
            Key             =   "uline"
            Object.Tag             =   "uline"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":30D0
            Key             =   "cpallette"
            Object.Tag             =   "cpallette"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":3612
            Key             =   "numbers"
            Object.Tag             =   "numbers"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":3B54
            Key             =   "bullets"
            Object.Tag             =   "bullets"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":4096
            Key             =   "outdent"
            Object.Tag             =   "outdent"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":45D8
            Key             =   "indent"
            Object.Tag             =   "indent"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":4B1A
            Key             =   "ljust"
            Object.Tag             =   "ljust"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":505C
            Key             =   "center"
            Object.Tag             =   "center"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbemform.frx":559E
            Key             =   "rjust"
            Object.Tag             =   "rjust"
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu frmFileNewInst 
         Caption         =   "New &Instance"
      End
      Begin VB.Menu hyphz1 
         Caption         =   "-"
      End
      Begin VB.Menu FileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu FileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAsText 
         Caption         =   "Save As &Text..."
      End
      Begin VB.Menu FileNewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "Re&vert..."
      End
      Begin VB.Menu mnuEditHtmlSrc 
         Caption         =   "&Edit HTML Source..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "P&roperties..."
      End
      Begin VB.Menu FileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu EditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu EditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu hyph1x 
         Caption         =   "-"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu EditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu hyph2 
         Caption         =   "-"
      End
      Begin VB.Menu EditSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditFind 
         Caption         =   "&Find Text"
         Shortcut        =   ^F
      End
      Begin VB.Menu hyph3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUnformat 
         Caption         =   "Remove All Formatting"
      End
      Begin VB.Menu mnuEditHyperlink 
         Caption         =   "Hyperlink..."
      End
      Begin VB.Menu mnuUnLink 
         Caption         =   "Un-Link"
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Begin VB.Menu ViewSub 
         Caption         =   "&Borders"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu ViewSub 
         Caption         =   "&Document Details"
         Index           =   1
      End
      Begin VB.Menu EditSnapToGrid 
         Caption         =   "&Snap To Grid"
      End
      Begin VB.Menu hyph21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelHtml 
         Caption         =   "&View Selected Source..."
      End
   End
   Begin VB.Menu Insert 
      Caption         =   "&Insert"
      Begin VB.Menu InsertSub 
         Caption         =   "&Picture..."
         Index           =   0
      End
      Begin VB.Menu InsertSub 
         Caption         =   "&Hyperlink..."
         Index           =   1
         Shortcut        =   ^L
      End
      Begin VB.Menu InsertButton 
         Caption         =   "&Button"
      End
      Begin VB.Menu hyph1 
         Caption         =   "-"
      End
      Begin VB.Menu InsertHTML 
         Caption         =   "HTML &Code..."
      End
   End
   Begin VB.Menu Format 
      Caption         =   "F&ormat"
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   15
      End
   End
   Begin VB.Menu Table 
      Caption         =   "T&able"
      Begin VB.Menu TableSub 
         Caption         =   "Insert Table..."
         Index           =   0
      End
      Begin VB.Menu TableSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu TableSub 
         Caption         =   "Insert Row"
         Index           =   2
      End
      Begin VB.Menu TableSub 
         Caption         =   "Insert Column"
         Index           =   3
      End
      Begin VB.Menu TableSub 
         Caption         =   "Insert Cell"
         Index           =   4
      End
      Begin VB.Menu TableSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu TableSub 
         Caption         =   "Delete Rows"
         Index           =   6
      End
      Begin VB.Menu TableSub 
         Caption         =   "Delete Columns"
         Index           =   7
      End
      Begin VB.Menu TableSub 
         Caption         =   "Delete Cells"
         Index           =   8
      End
      Begin VB.Menu TableSub 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu TableSub 
         Caption         =   "Merge Cells"
         Index           =   10
      End
      Begin VB.Menu TableSub 
         Caption         =   "Split Cell"
         Index           =   11
      End
   End
   Begin VB.Menu D2D 
      Caption         =   "&2D"
      Begin VB.Menu D2DSub 
         Caption         =   "Set Position Attribute To Absolute"
         Index           =   0
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Bring To Front"
         Index           =   1
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Send To Back"
         Index           =   2
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Bring Forward"
         Index           =   3
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Send Back"
         Index           =   4
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Bring Above Text"
         Index           =   5
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Send Below Text"
         Index           =   6
      End
      Begin VB.Menu D2DSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Lock Element"
         Index           =   8
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "(debug)"
      Begin VB.Menu a 
         Caption         =   "a"
      End
      Begin VB.Menu b 
         Caption         =   "b"
      End
      Begin VB.Menu c 
         Caption         =   "c"
      End
      Begin VB.Menu d 
         Caption         =   "d"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1999 Microsoft Corporation.
' All rights reserved.
' Author: Rick Jesse
Option Explicit


Dim DHTMLEditInitialized As Boolean
Dim fontNames(0 To 5) ' list of fonts is the fontComboBox
Dim fontSizes(0 To 6) ' DHTMLEdit font sizes are 1-7
' Tables for toolbar commands
Dim buttonCmds(1 To 12) As DHTMLEDITCMDID
Dim buttonNames(1 To 12) As String
' Tables for menus commands
Dim editMenuCmds(0 To 8) As DHTMLEDITCMDID
Dim insertMenuCmds(0 To 1) As DHTMLEDITCMDID
Dim tableMenuCmds(0 To 11) As DHTMLEDITCMDID
Dim twoDMenuCmds(0 To 8) As DHTMLEDITCMDID
' Document path name
Dim sDocPath As String
' State variables for dynamic context menu
Dim ctxtIs2DCapable As Boolean
Dim ctxtIsAbsPos As Boolean
Dim ctxtIsTable As Boolean
Dim ctxtStdItemCount As Long
Dim ctxt2DItemCount As Long
Dim ctxtTableItemCount As Long

Private Enum General
    DE_E_INVALIDARG = &H5
    DE_E_ACCESS_DENIED = &H46
    DE_E_PATH_NOT_FOUND = &H80070003
    DE_E_FILE_NOT_FOUND = &H80070002
    DE_E_UNEXPECTED = &H8000FFFF
    DE_E_DISK_FULL = &H80070027
    DE_E_NOTSUPPORTED = &H80040100
    DE_E_FILTER_FRAMESET = &H80100001
    DE_E_FILTER_SERVERSCRIPT = &H80100002
    DE_E_FILTER_MULTIPLETAGS = &H80100004
    DE_E_FILTER_SCRIPTLISTING = &H80100008
    DE_E_FILTER_SCRIPTLABEL = &H80100010
    DE_E_FILTER_SCRIPTTEXTAREA = &H80100020
    DE_E_FILTER_SCRIPTSELECT = &H80100040
    DE_E_URL_SYNTAX = &H800401E4
    DE_E_INVALID_URL = &H800C0002
    DE_E_NO_SESSION = &H800C0003
    DE_E_CANNOT_CONNECT = &H800C0004
    DE_E_RESOURCE_NOT_FOUND = &H800C0005
    DE_E_OBJECT_NOT_FOUND = &H800C0006
    DE_E_DATA_NOT_AVAILABLE = &H800C0007
    DE_E_DOWNLOAD_FAILURE = &H800C0008
    DE_E_AUTHENTICATION_REQUIRED = &H800C0009
    DE_E_NO_VALID_MEDIA = &H800C000A
    DE_E_CONNECTION_TIMEOUT = &H800C000B
    DE_E_INVALID_REQUEST = &H800C000C
    DE_E_UNKNOWN_PROTOCOL = &H800C000D
    DE_E_SECURITY_PROBLEM = &H800C000E
    DE_E_CANNOT_LOAD_DATA = &H800C000F
    DE_E_CANNOT_INSTANTIATE_OBJECT = &H800C0010
    DE_E_REDIRECT_FAILED = &H800C0014
    DE_E_REDIRECT_TO_DIR = &H800C0015
    DE_E_CANNOT_LOCK_REQUEST = &H8
End Enum

Private Const APP_TITLE = "DHTML Edit"

Private Sub GetElementUnderInsertionPoint()
Dim rg As IHTMLTxtRange
Dim ctlRg As IHTMLControlRange

Select Case DHTMLEdit1.DOM.selection.Type
   
   Case "None", "Text"
      ' This reduces the selection to just the insertion
      ' point. The parentElement method will then return the
      ' element directly under the mouse pointer.
      Set rg = DHTMLEdit1.DOM.selection.createRange
'      rg.collapse
'      MsgBox rg.parentElement.outerHTML
      MsgBox rg.HTMLText
   Case "Control"
      ' A form or image is selected. The commonParentElement
      ' will return the site selected element.
      Set ctlRg = DHTMLEdit1.DOM.selection.createRange
'      MsgBox ctlRg.commonParentElement.outerHTML
'      MsgBox ctlRg.commonParentElement.tagName
'VIEW SELECTED SOURCE:
      MsgBox ctlRg.Item(0).outerHTML
End Select


End Sub
Private Function Loadfile()

    sDocPath = ""
    DisableToolbar
    
    If Not SaveChanges = vbCancel Then
    
        On Error Resume Next
        DHTMLEdit1.LoadDocument Replace(command$, """", ""), False
        
        If Err.Number < 0 Then
            Dim errMsg As String
            Select Case Err.Number
                Case DE_E_INVALIDARG
                    errMsg = "Invalid argument"
                Case DE_E_PATH_NOT_FOUND
                    errMsg = "Path not found"
                Case DE_E_FILE_NOT_FOUND
                    errMsg = "File not found"
                Case DE_E_ACCESS_DENIED
                    errMsg = "Access denied"
                Case DE_E_UNEXPECTED
                    errMsg = "Unexpected error"
                Case DE_E_FILTER_FRAMESET
                    errMsg = "Document contains a frameset"
                Case DE_E_FILTER_SERVERSCRIPT
                    errMsg = "Document is primarily server side script"
                Case Else
                    errMsg = "Unknown error"
            End Select
            
            MsgBox "Error occurred while loading document: " & errMsg & ".", vbCritical
            DHTMLEdit1.NewDocument
        End If
    End If
    
    If DHTMLEdit1.Busy = False Then
        On Error Resume Next
        ' Force a DisplayChanged event to update toolbar
        ' in case user canceled file open dialog or error occurred
        DHTMLEdit1.DOM.selection.createTextRange.collapse
    End If
    SetFormCaption

End Function

Private Sub WaitWhileBusy()

Do While MainForm.DHTMLEdit1.Busy
    DoEvents
Loop

End Sub
Private Sub a_Click()

'Dim m As MSHTML.HTMLBody

'style:background-color
'DHTMLEdit1.ExecCommand DECMD_SETBACKCOLOR, OLECMDEXECOPT_DODEFAULT, "red"
'MainForm.DHTMLEdit1.DOM.ExecCommand "refresh"

'Set m = DHTMLEdit1.DOM.body


'MsgBox m.document.

'Set m = Nothing

Dim P As Object, i, s

On Error Resume Next
Dim idx
Set P = DHTMLEdit1.DOM.All.tags("img")
MsgBox P.length
For idx = 0 To P.length - 1
    P.Item(idx).outerHTML = ""
Next

'DHTMLEdit1.DocumentHTML = RX_RemoveTagWithContents(DHTMLEdit1.DocumentHTML, "img", True)

End Sub
Private Sub b_Click()

'DHTMLEdit1.DOM.Title = "XXX"

Dim P As Object, i, s
Set P = DHTMLEdit1.DOM.All.tags("p")
For i = 0 To P.length - 1
    P.Item(i).outerText = "X"
Next

'DHTMLEdit1.DOM.All.tags("title").Item(0).innerText = "ZZZ"

DHTMLEdit1.DOM.bgColor = "red"
'MsgBox DHTMLEdit1.DocumentTitle
MsgBox DHTMLEdit1.DocumentHTML

End Sub
Private Sub c_Click()
'IHTMLElement::insertAdjacent

MsgBox DHTMLEdit1.DOM.body.outerText
'MsgBox DHTMLEdit1.DOM.images.Item(0)

MsgBox DHTMLEdit1.DOM.selection.createRange.Text

End Sub

Private Function GuessFileName() As String
Dim nPos As Long, sTemp As String

If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
        'there is a loaded file, use its name
        GuessFileName = DHTMLEdit1.CurrentDocumentPath

Else
        'try to guess a name
        sTemp = Left(DHTMLEdit1.DOM.body.outerText, 1024)
        nPos = InStr(1, sTemp, vbCrLf)
        
        If nPos <> 0 Then
                sTemp = Left(sTemp, nPos - 1)
                
                'Illegal filename chars
                '"\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
                sTemp = Replace(sTemp, "\", "")
                sTemp = Replace(sTemp, "/", "")
                sTemp = Replace(sTemp, ":", "")
                sTemp = Replace(sTemp, "*", "")
                sTemp = Replace(sTemp, "?", "")
                sTemp = Replace(sTemp, "<", "")
                sTemp = Replace(sTemp, ">", "")
                sTemp = Replace(sTemp, "|", "")
                sTemp = Replace(sTemp, """", "")
        Else
        
                sTemp = Left(sTemp, 255)
        
        End If

        'If Len(sTemp) <= 1 Then sTemp = ""

        GuessFileName = sTemp

End If

End Function
Private Sub d_Click()

'MsgBox DHTMLEdit1.ExecCommand(DECMD_GETBACKCOLOR, OLECMDEXECOPT_DODEFAULT)
'MsgBox DHTMLEdit1.FilterSourceCode(DHTMLEdit1.DocumentHTML)
'MsgBox DHTMLEdit1.CurrentDocumentPath
GetElementUnderInsertionPoint

End Sub
Private Sub D2DSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    Dim state As DHTMLEDITCMDF
    
    cmd = twoDMenuCmds(Index)
           
    If Not cmd = 0 Then
        DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
    End If

    state = DHTMLEdit1.QueryStatus(DECMD_MAKE_ABSOLUTE)
    
    If state = DECMDF_LATCHED Then
        D2DSub(0).Caption = "Set Position Attribute To 1D"
        D2DSub(0).Enabled = True
    ElseIf state = DECMDF_ENABLED Then
        D2DSub(0).Caption = "Set Position Attribute To Absolute"
        D2DSub(0).Enabled = True
    Else
        D2DSub(0).Caption = "Set Position Attribute To Absolute"
        D2DSub(0).Enabled = False
    End If

    
End Sub

Private Sub DHTMLEdit1_DocumentComplete()
    If Not DHTMLEditInitialized Then
        Dim fmt As DEGetBlockFmtNamesParam
        Dim i As Long
        Dim fontSize As Long
        Dim fmtName As Variant
        
        ' Create the block fmt names holder
        Set fmt = CreateObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam.1")
        
        ' Get the localized strings for the DECMD_SETBLOCKFMT command
        DHTMLEdit1.ExecCommand DECMD_GETBLOCKFMTNAMES, OLECMDEXECOPT_DONTPROMPTUSER, fmt
        
        ' Put the strings into the Format menu
        i = 0
        For Each fmtName In fmt.Names
            FormatSub(i).Caption = fmtName
            i = i + 1
        Next
        
        UpdateFontCombos
        
        FontSizeCombo.ListIndex = fontSize - 1
        
    End If
    DHTMLEditInitialized = True
End Sub

Private Sub DHTMLEdit1_OnMouseOver()
Dim sElement As String

'Do While DHTMLEdit1.Busy
'    DoEvents
'Loop
    
If DHTMLEdit1.Busy Then Exit Sub

Dim e As IHTMLEventObj
Set e = DHTMLEdit1.DOM.parentWindow.event

sElement = e.srcElement.outerHTML


Select Case e.srcElement.tagName
    
    Case "A"
        StatusBar1.Panels("mousemove").Text = GetLinkHref(sElement)
        'doesn't work with image links!
    
    Case "IMG"
        'check if it is Linked
        If e.srcElement.parentElement.tagName = "A" Then
            sElement = e.srcElement.parentElement.outerHTML
            StatusBar1.Panels("mousemove").Text = GetLinkHref(sElement)
        End If

    Case Else
        StatusBar1.Panels("mousemove").Text = ""  'sElement

End Select


End Sub
Private Sub EditCopy_Click()
DHTMLEdit1.ExecCommand DECMD_COPY, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub EditCut_Click()
DHTMLEdit1.ExecCommand DECMD_CUT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub EditFind_Click()
DHTMLEdit1.ExecCommand DECMD_FINDTEXT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub EditPaste_Click()
DHTMLEdit1.ExecCommand DECMD_PASTE, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub EditRedo_Click()

DHTMLEdit1.ExecCommand DECMD_REDO, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub EditSelAll_Click()
DHTMLEdit1.ExecCommand DECMD_SELECTALL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub EditSnapToGrid_Click()

Dim bState As Boolean
    
bState = DHTMLEdit1.SnapToGrid
bState = Not bState
DHTMLEdit1.SnapToGrid = bState

EditSnapToGrid.Checked = bState

End Sub
Private Sub EditUndo_Click()

DHTMLEdit1.ExecCommand DECMD_UNDO, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub FileExit_Click()
    
        Unload Me

End Sub
Private Sub FileNew_Click()

    If Not SaveChanges = vbCancel Then
        sDocPath = ""
        DHTMLEdit1.NewDocument
        SetFormCaption
    End If

End Sub
Private Sub FileOpen_Click()


    sDocPath = ""
    DisableToolbar
    
    If Not SaveChanges = vbCancel Then
    
        On Error Resume Next
        DHTMLEdit1.LoadDocument "", True
        
        If Err.Number < 0 Then
            Dim errMsg As String
            Select Case Err.Number
                Case DE_E_INVALIDARG
                    errMsg = "Invalid argument"
                Case DE_E_PATH_NOT_FOUND
                    errMsg = "Path not found"
                Case DE_E_FILE_NOT_FOUND
                    errMsg = "File not found"
                Case DE_E_ACCESS_DENIED
                    errMsg = "Access denied"
                Case DE_E_UNEXPECTED
                    errMsg = "Unexpected error"
                Case DE_E_FILTER_FRAMESET
                    errMsg = "Document contains a frameset"
                Case DE_E_FILTER_SERVERSCRIPT
                    errMsg = "Document is primarily server side script"
                Case Else
                    errMsg = "Unknown error"
            End Select
            
            MsgBox "Error occurred while loading document: " & errMsg & ".", vbCritical
            DHTMLEdit1.NewDocument
        End If
    End If
    
    If DHTMLEdit1.Busy = False Then
        On Error Resume Next
        ' Force a DisplayChanged event to update toolbar
        ' in case user canceled file open dialog or error occurred
        DHTMLEdit1.DOM.selection.createTextRange.collapse
    End If
    SetFormCaption
    
End Sub
Private Sub FileSave_Click()


    If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
        SaveDocument False
    Else
        SaveDocument True
    End If


End Sub
Private Sub FileSaveAs_Click()
    SaveDocument True
End Sub

Private Sub FontCombo_Click()
    Dim fn As String
    Dim state As DHTMLEDITCMDF
    
    fn = fontNames(FontCombo.ListIndex)
    
    If (DHTMLEditInitialized) Then
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTNAME)
        If state >= DECMDF_ENABLED Then
            DHTMLEdit1.ExecCommand DECMD_SETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, fn
        End If
    End If
    
End Sub


Private Sub FontCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim state As DHTMLEDITCMDF

    ' return if not the return key
    If Not KeyCode = vbKeyReturn Then
        Exit Sub
    End If
    
    If (DHTMLEditInitialized) Then
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTNAME)
        If state >= DECMDF_ENABLED Then
            ' set the font to what user has typed into the font name combo box
            DHTMLEdit1.ExecCommand DECMD_SETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, FontCombo.Text
        End If
    End If
    
End Sub

Private Sub FontSizeCombo_Click()
    Dim fs As Long
    Dim state As DHTMLEDITCMDF
    
    fs = FontSizeCombo.ListIndex
    fs = fs + 1
    
    If (DHTMLEditInitialized) Then
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTSIZE)
        If state >= DECMDF_ENABLED Then
            DHTMLEdit1.ExecCommand DECMD_SETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, fs
        End If
    End If
    
End Sub

Private Sub FontSizeCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim state As DHTMLEDITCMDF
    Dim fs As String

    ' return if not the return key
    If Not KeyCode = vbKeyReturn Then
        Exit Sub
    End If
    
    If (DHTMLEditInitialized) Then
    
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTSIZE)
        
        If state >= DECMDF_ENABLED Then
        
            ' remember what's in the combox box so we can reset if the user
            ' typed in something invalid
            
            ' if its mixed selected, display an empty string
            If state = DECMDF_NINCHED Then
                fs = ""
            Else
                fs = DHTMLEdit1.ExecCommand(DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER)
                fs = fs - 1
            End If
            
            ' validate what the user type in
            If IsNumeric(FontSizeCombo.Text) = False Then
                ' didn't type in a valid number
                FontSizeCombo.Text = fs
            ElseIf FontSizeCombo.Text < 1 Or FontSizeCombo.Text > 7 Then
                ' number is out of range
                FontSizeCombo.Text = fs
            Else
                ' set the font size to the number the user typed in
                DHTMLEdit1.ExecCommand DECMD_SETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, FontSizeCombo.Text
            End If
        End If
    End If
    
End Sub

Private Sub Form_Activate()

'DHTMLEdit1.SetFocus

End Sub

Private Sub Form_Load()

    InitCommonControls
    MainForm.Caption = APP_TITLE

    DHTMLEdit1.Top = Me.Toolbar1.Height + Me.Toolbar2.Height
    DHTMLEdit1.Left = 0
    DHTMLEdit1.SourceCodePreservation = False   'Allow the control to change the document's original white spacing

    DHTMLEditInitialized = False

    ' Initialize the font name and size combo boxes
    fontNames(0) = "Times New Roman"
    fontNames(1) = "Arial"
    fontNames(2) = "Tahoma"
    fontNames(3) = "Courier"
    fontNames(4) = "Verdana"
    fontNames(5) = "Wingdings"

    fontSizes(0) = "1"
    fontSizes(1) = "2"
    fontSizes(2) = "3"
    fontSizes(3) = "4"
    fontSizes(4) = "5"
    fontSizes(5) = "6"
    fontSizes(6) = "7"

    FontCombo.AddItem fontNames(0)
    FontCombo.AddItem fontNames(1)
    FontCombo.AddItem fontNames(2)
    FontCombo.AddItem fontNames(3)
    FontCombo.AddItem fontNames(4)
    FontCombo.AddItem fontNames(5)
    FontCombo.ListIndex = 0

    FontSizeCombo.AddItem fontSizes(0)
    FontSizeCombo.AddItem fontSizes(1)
    FontSizeCombo.AddItem fontSizes(2)
    FontSizeCombo.AddItem fontSizes(3)
    FontSizeCombo.AddItem fontSizes(4)
    FontSizeCombo.AddItem fontSizes(5)
    FontSizeCombo.AddItem fontSizes(6)
    FontSizeCombo.ListIndex = 0

    InitToolbarTable
    InitMenuTables

    DisableToolbar

Toolbar2.ImageList = ImageList2
Toolbar2.Buttons("new").Image = ImageList2.ListImages("new").Index
Toolbar2.Buttons("open").Image = ImageList2.ListImages("open").Index
Toolbar2.Buttons("save").Image = ImageList2.ListImages("save").Index
Toolbar2.Buttons("props").Image = ImageList2.ListImages("props").Index
Toolbar2.Buttons("html").Image = ImageList2.ListImages("html").Index
Toolbar2.Buttons("revert").Image = ImageList2.ListImages("revert").Index

Toolbar2.Buttons("cut").Image = ImageList2.ListImages("cut").Index
Toolbar2.Buttons("copy").Image = ImageList2.ListImages("copy").Index
Toolbar2.Buttons("paste").Image = ImageList2.ListImages("paste").Index
Toolbar2.Buttons("undo").Image = ImageList2.ListImages("undo").Index
Toolbar2.Buttons("redo").Image = ImageList2.ListImages("redo").Index

If command$ <> "" Then Loadfile

End Sub
Private Sub Form_Resize()
    
    If Not MainForm.WindowState = vbMinimized Then
        DHTMLEdit1.Width = MainForm.ScaleWidth
        DHTMLEdit1.Height = MainForm.ScaleHeight - Me.Toolbar1.Height - Me.Toolbar2.Height - Me.StatusBar1.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If DHTMLEdit1.IsDirty Then

    If Not SaveChanges = vbCancel Then
        
    Else
    
        Cancel = True
    End If
End If

Unload frmEditSource
Unload frmHidden

End Sub
Private Sub FormatSub_Click(Index As Integer)
    Dim state As DHTMLEDITCMDF
    Dim Format As String
    
    state = DHTMLEdit1.QueryStatus(DECMD_SETBLOCKFMT)
    
    If state >= DECMDF_ENABLED Then
        DHTMLEdit1.ExecCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, FormatSub(Index).Caption
    End If
    
End Sub

Private Sub frmFileNewInst_Click()
On Error Resume Next

Shell RemoveSlash(App.Path) & "\" & App.EXEName & ".exe", vbMaximizedFocus

End Sub
Private Sub Insert_Click()
    Dim cmdIndex As Long
    
    For cmdIndex = LBound(insertMenuCmds) To UBound(insertMenuCmds)
        UpdateMenu InsertSub(cmdIndex), insertMenuCmds(cmdIndex)
    Next cmdIndex
    
    If DHTMLEdit1.DOM.selection.Type = "Control" Then ' a control, table, ActiveX control is selected
        InsertButton.Enabled = False
        InsertHTML.Enabled = False
    Else
        InsertButton.Enabled = True
        InsertHTML.Enabled = True
    End If

End Sub

Private Sub InsertButton_Click()
    Dim doc As Object
    Dim selection As Object
    Dim tr As Object
    ' This routine inserts a button at the current selection
    
    ' Get the DHTML Document object
    Set doc = DHTMLEdit1.DOM
    ' Get the DHTML Selection object
    Set selection = doc.selection
    ' Create a TextRange on the current selection
    Set tr = selection.createRange
    
    tr.pasteHTML ("<BUTTON TITLE=Button>Button!</BUTTON>")
    
End Sub

Private Sub InsertHTML_Click()
    InsertHTMLDlg.Show vbModal, Me
End Sub

Private Sub InsertSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    
    cmd = insertMenuCmds(Index)
           
    If Not cmd = 0 Then
        DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
    End If


End Sub

Private Sub mnuAbout_Click()
    
    frmAbout.Show vbModal, Me

End Sub
Private Sub mnuEdit_Click()
    
UpdateMenu EditUndo, DECMD_UNDO
UpdateMenu EditRedo, DECMD_REDO
UpdateMenu EditCut, DECMD_CUT
UpdateMenu EditCopy, DECMD_COPY
UpdateMenu EditPaste, DECMD_PASTE
UpdateMenu mnuEditDelete, DECMD_DELETE
UpdateMenu EditSelAll, DECMD_SELECTALL
UpdateMenu EditFind, DECMD_FINDTEXT
UpdateMenu mnuUnLink, DECMD_UNLINK
UpdateMenu mnuEditUnformat, DECMD_REMOVEFORMAT

'if we can unlink then it is a hyper link !
UpdateMenu mnuEditHyperlink, DECMD_UNLINK

End Sub
Private Sub mnuEditDelete_Click()

DHTMLEdit1.ExecCommand DECMD_DELETE, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuEditHtmlSrc_Click()
Dim vTemp As Variant

g_IsDirty = MainForm.DHTMLEdit1.IsDirty

frmEditSource.Show vbModal

If g_IsDirty = True And MainForm.DHTMLEdit1.IsDirty = False Then
    'not-ready error appears sometimes here, so we better wait
    Do While MainForm.DHTMLEdit1.Busy
        DoEvents
    Loop
    'DHTMLEdit.IsDirty does not get auto-updated, so we force it to update
    vTemp = MainForm.DHTMLEdit1.DOM.bgColor
    MainForm.DHTMLEdit1.DOM.bgColor = "@!~%^$"  'a value that's unlikely to exist
    MainForm.DHTMLEdit1.DOM.bgColor = vTemp
End If


'MsgBox g_IsDirty
'MsgBox DHTMLEdit1.IsDirty

End Sub

Private Sub mnuEditHyperlink_Click()

'Do While DHTMLEdit1.Busy
'    DoEvents
'Loop

DHTMLEdit1.ExecCommand DECMD_HYPERLINK, OLECMDEXECOPT_DODEFAULT

End Sub
Private Sub mnuEditUnformat_Click()

DHTMLEdit1.ExecCommand DECMD_REMOVEFORMAT, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuFileRevert_Click()
Dim Result As VbMsgBoxResult
Dim sCurrentFile As String

On Error GoTo Err_Revert

sCurrentFile = DHTMLEdit1.CurrentDocumentPath

'Only if we have a filename
If Len(sCurrentFile) = 0 Then
    'no file name!
    Beep
    Exit Sub
End If

If DHTMLEdit1.IsDirty Then
    Result = MsgBox("Loose All Changes to  """ & sCurrentFile & """  Since Last Save?", vbOKCancel + vbQuestion + vbDefaultButton2, "Revert")
Else
    'no need to ask...
    Result = vbOK
End If

If Result = vbCancel Then
    'do nothing
ElseIf Result = vbOK Then
    DHTMLEdit1.LoadDocument sCurrentFile
End If

Exit Sub
Err_Revert:
    MsgBox Err.Description, vbCritical, "Error in [Revert]"
    Err.Clear

End Sub
Private Sub mnuFileSaveAsText_Click()
Dim iFile As Integer
Dim sDocumentText As String, sSuggestedFileName As String


If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
    sSuggestedFileName = ChangeFileExtension(DHTMLEdit1.CurrentDocumentPath, "txt")
    CommonDialog1.FileName = sSuggestedFileName
Else
    CommonDialog1.FileName = ""
End If


CommonDialog1.Flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
CommonDialog1.Filter = "Text Files (*.txt)|*.txt|All Files|*.*"
CommonDialog1.CancelError = True

On Error Resume Next

CommonDialog1.ShowSave
If Err Then
    Err.Clear
Else
    sDocumentText = DHTMLEdit1.DOM.body.outerText
    iFile = FreeFile
    Open CommonDialog1.FileName For Output Access Write Shared As #iFile
    Print #iFile, sDocumentText
    Close #iFile
End If


End Sub
Private Sub mnuProperties_Click()
Dim sBaseURL  As String
Dim sOldTitle As String

sBaseURL = MainForm.DHTMLEdit1.BaseURL
sOldTitle = MainForm.DHTMLEdit1.DOM.Title

frmProperties.Show vbModal

'Do While MainForm.DHTMLEdit1.Busy
'    DoEvents
'Loop
'DHTMLEdit.IsDirty does not get auto-updated, so we force it to update
'MainForm.DHTMLEdit1.DOM.bgColor = MainForm.DHTMLEdit1.DOM.bgColor

'Dim s, i, ii, ss, sss
'sss = "<title>" & MainForm.DHTMLEdit1.DOM.Title & "</title>"
's = DHTMLEdit1.DocumentHTML
'i = InStr(1, s, "<title>", vbTextCompare)
'ii = InStr(1, s, "</title>", vbTextCompare)
'ss = Mid(s, i, 8 + ii - i)
'DHTMLEdit1.DocumentHTML = Replace(s, ss, sss)
'

If sOldTitle <> MainForm.DHTMLEdit1.DOM.Title Then
    DHTMLEdit1.DocumentHTML = RX_GenericReplace(DHTMLEdit1.DocumentHTML, "<title>.*?</title>", "<title>" & MainForm.DHTMLEdit1.DOM.Title & "</title>", True)
    MainForm.DHTMLEdit1.BaseURL = sBaseURL
End If

End Sub
Private Sub mnuUnLink_Click()

DHTMLEdit1.ExecCommand DECMD_UNLINK, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuViewSelHtml_Click()
Dim sSelHtml As String

Select Case DHTMLEdit1.DOM.selection.Type
   
   Case "None", "Text"
      'MsgBox DHTMLEdit1.DOM.selection.createRange.parentElement.outerHTML
      sSelHtml = DHTMLEdit1.DOM.selection.createRange.HTMLText
   
   Case "Control" ' including IMG
      sSelHtml = DHTMLEdit1.DOM.selection.createRange.Item(0).outerHTML
    
End Select


frmViewSelHtml.rtfHTML.Text = sSelHtml
ColorizeCode frmViewSelHtml.rtfHTML
frmViewSelHtml.Show vbModal, Me

End Sub
Private Sub Table_Click()
    Dim cmdIndex As Long
    
    For cmdIndex = LBound(tableMenuCmds) To UBound(tableMenuCmds)
        UpdateMenu TableSub(cmdIndex), tableMenuCmds(cmdIndex)
    Next cmdIndex

End Sub

Private Sub D2D_Click()
    Dim cmdIndex As Long
    Dim state As DHTMLEDITCMDF
    
    For cmdIndex = LBound(twoDMenuCmds) To UBound(twoDMenuCmds)
        UpdateMenu D2DSub(cmdIndex), twoDMenuCmds(cmdIndex)
    Next cmdIndex

    state = DHTMLEdit1.QueryStatus(DECMD_LOCK_ELEMENT)
    If state = DECMDF_LATCHED Then
        D2DSub(8).Checked = True
    Else
        D2DSub(8).Checked = False
    End If

End Sub

Private Sub TableSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    
    If Index = 0 Then
        InsertTableDlg.Show vbModal, Me
    Else
        cmd = tableMenuCmds(Index)
               
        If Not cmd = 0 Then
            DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
        End If
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    ' Handle toolbar commands
    Select Case Button.Key
        Case "Bold"
            DHTMLEdit1.ExecCommand DECMD_BOLD, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Italic"
            DHTMLEdit1.ExecCommand DECMD_ITALIC, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Underline"
            DHTMLEdit1.ExecCommand DECMD_UNDERLINE, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Numbers"
            DHTMLEdit1.ExecCommand DECMD_ORDERLIST, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Bullets"
            DHTMLEdit1.ExecCommand DECMD_UNORDERLIST, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Outdent"
            DHTMLEdit1.ExecCommand DECMD_OUTDENT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Indent"
            DHTMLEdit1.ExecCommand DECMD_INDENT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "LeftJustify"
             DHTMLEdit1.ExecCommand DECMD_JUSTIFYLEFT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Center"
            DHTMLEdit1.ExecCommand DECMD_JUSTIFYCENTER, OLECMDEXECOPT_DONTPROMPTUSER
        Case "RightJustify"
            DHTMLEdit1.ExecCommand DECMD_JUSTIFYRIGHT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Color"
'            Dim foreColor As String
'            On Error GoTo cleanup
'            CommonDialog1.Color = 0
'            CommonDialog1.CancelError = True
'            CommonDialog1.Flags = cdlCCFullOpen
'            CommonDialog1.ShowColor
'            foreColor = ""
'            foreColor = FormatRGBString(CommonDialog1.Color)
'            DHTMLEdit1.ExecCommand DECMD_SETFORECOLOR, OLECMDEXECOPT_DONTPROMPTUSER, foreColor
             frmHidden.Tag = "forecolor"
             PopupMenu frmHidden.mnuColors, vbPopupMenuRightButton, Button.Left, Button.Top + Button.Height + Toolbar1.Height + 35 '+ Toolbar2.Height
        Case "BackColor"
'            Dim sBackColor As String
'            On Error GoTo Cleanup
'            CommonDialog1.Color = 0
'            CommonDialog1.CancelError = True
'            CommonDialog1.Flags = cdlCCFullOpen
'            CommonDialog1.ShowColor
'            sBackColor = ""
'            sBackColor = FormatRGBString(CommonDialog1.Color)
'            DHTMLEdit1.ExecCommand DECMD_SETBACKCOLOR, OLECMDEXECOPT_DONTPROMPTUSER, sBackColor
             frmHidden.Tag = "backcolor"
             PopupMenu frmHidden.mnuColors, vbPopupMenuRightButton, Button.Left, Button.Top + Button.Height + Toolbar1.Height + 35 '+ Toolbar2.Height
        End Select
    
Exit Sub
Cleanup:
    Err.Clear

End Sub
Private Sub DHTMLEdit1_ContextMenuAction(ByVal itemIndex As Long)
GoTo AvoidMe
    ' Handle user selection on the custom context menu
   Select Case itemIndex
    Case 0
        DHTMLEdit1.ExecCommand DECMD_CUT, OLECMDEXECOPT_DODEFAULT
    Case 1
        DHTMLEdit1.ExecCommand DECMD_COPY, OLECMDEXECOPT_DODEFAULT
    Case 2
        DHTMLEdit1.ExecCommand DECMD_PASTE, OLECMDEXECOPT_DODEFAULT
    Case 4
        DHTMLEdit1.ExecCommand DECMD_SELECTALL, OLECMDEXECOPT_DODEFAULT
    Case 6
        DHTMLEdit1.ExecCommand DECMD_FONT, OLECMDEXECOPT_PROMPTUSER
    End Select
    
    If ctxtIs2DCapable Then
        Select Case itemIndex
        Case ctxtStdItemCount + 2
            DHTMLEdit1.ExecCommand DECMD_MAKE_ABSOLUTE, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
    If ctxtIsTable Then
        Select Case itemIndex
        Case ctxtStdItemCount + ctxt2DItemCount + 2
            DHTMLEdit1.ExecCommand DECMD_INSERTROW, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 3
            DHTMLEdit1.ExecCommand DECMD_INSERTCOL, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 5
            DHTMLEdit1.ExecCommand DECMD_DELETEROWS, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 6
            DHTMLEdit1.ExecCommand DECMD_DELETECOLS, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
AvoidMe:

End Sub
Private Sub DHTMLEdit1_ShowContextMenu(ByVal x As Long, ByVal y As Long)
GoTo Avoid
    Dim cmdState As DHTMLEDITCMDF
    Dim strings() As String
    Dim states() As OLE_TRISTATE
    
   ' Create dynamic context menu that consists of
   ' a "standard" set of items and items that depend
   ' on the currently selected element.
   ' Look at the current selection and
   ' if its a table then add menu items for add/delete rows and cols
   ' if its 2DCapable then add items to toggle its absolute position attribute
   
    ctxtIs2DCapable = False
    ctxtIsAbsPos = False
    ctxtIsTable = False
        
    ' Determine if the selected element is 2D capable
    cmdState = DHTMLEdit1.QueryStatus(DECMD_MAKE_ABSOLUTE)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIs2DCapable = True
    End If
    
    'Use DECMD_SEND_TO_BACK to determine if this element is abs positioned
    cmdState = DHTMLEdit1.QueryStatus(DECMD_SEND_TO_BACK)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIsAbsPos = True
    End If
    
    'Use DECMD_INSERTROW to determine if this element is a table
    cmdState = DHTMLEdit1.QueryStatus(DECMD_INSERTROW)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIsTable = True
    End If
    
    ctxtStdItemCount = 6
    
    If ctxtIs2DCapable Then
        ctxt2DItemCount = 2 '1 Item + 1 Separator
    Else
        ctxt2DItemCount = 0
    End If
    
    
    If ctxtIsTable Then
        ctxtTableItemCount = 6 '4 Items + 2 Separators
    Else
        ctxtTableItemCount = 0
    End If
    
    
    ReDim strings(0 To ctxtStdItemCount + ctxt2DItemCount + ctxtTableItemCount)
    ReDim states(0 To ctxtStdItemCount + ctxt2DItemCount + ctxtTableItemCount)
    
    strings(0) = "Cut"
    strings(1) = "Copy"
    strings(2) = "Paste"
    strings(3) = ""
    strings(4) = "Select All"
    strings(5) = ""
    strings(6) = "Font..."
        
    cmdState = DHTMLEdit1.QueryStatus(DECMD_CUT)
    If cmdState >= DECMDF_ENABLED Then
         states(0) = Unchecked
     Else
         states(0) = Gray
    End If
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_COPY)
    If cmdState >= DECMDF_ENABLED Then
         states(1) = Unchecked
     Else
         states(1) = Gray
    End If
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_PASTE)
    If cmdState >= DECMDF_ENABLED Then
         states(2) = Unchecked
     Else
         states(2) = Gray
    End If
        
    states(3) = Unchecked
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_SELECTALL)
    If cmdState >= DECMDF_ENABLED Then
         states(4) = Unchecked
     Else
         states(4) = Gray
    End If
    
    states(5) = Unchecked
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_FONT)
    If cmdState >= DECMDF_ENABLED Then
         states(6) = Unchecked
     Else
         states(6) = Gray
    End If
    
    If ctxtIs2DCapable Then
        strings(ctxtStdItemCount + 1) = ""
        states(ctxtStdItemCount + 1) = Unchecked
        If ctxtIsAbsPos Then
            strings(ctxtStdItemCount + 2) = "Make 1D"
        Else
            strings(ctxtStdItemCount + 2) = "Make 2D"
        End If
        states(ctxtStdItemCount + 2) = Unchecked
    End If
    
    If ctxtIsTable Then
        strings(ctxtStdItemCount + ctxt2DItemCount + 1) = ""
        states(ctxtStdItemCount + ctxt2DItemCount + 1) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 2) = "Insert Row"
        states(ctxtStdItemCount + ctxt2DItemCount + 2) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 3) = "Insert Column"
        states(ctxtStdItemCount + ctxt2DItemCount + 3) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 4) = ""
        states(ctxtStdItemCount + ctxt2DItemCount + 4) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 5) = "Delete Row"
        states(ctxtStdItemCount + ctxt2DItemCount + 5) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 6) = "Delete Column"
        states(ctxtStdItemCount + ctxt2DItemCount + 6) = Unchecked
        
    End If
    
    DHTMLEdit1.SetContextMenu strings, states
    
Avoid:

'ctxtIsTable = False
'Use DECMD_INSERTROW to determine if this element is a table
'cmdState = DHTMLEdit1.QueryStatus(DECMD_INSERTROW)
'If cmdState >= DECMDF_ENABLED Then
'    ctxtIsTable = True
'End If

'If ctxtIsTable Then
'    PopupMenu Table, vbPopupMenuRightButton
'Else
    PopupMenu mnuEdit, vbPopupMenuRightButton
'End If

End Sub
Private Sub DHTMLEdit1_DisplayChanged()
On Error GoTo DHTMLEdit1_DisplayChanged_Error
    Dim state As DHTMLEDITCMDF
    Dim cmd As DHTMLEDITCMDID
    Dim Button As String
    Dim cmds As Long
    
    ' DHTMLEdit indicates the UI should be updated
    ' First update the Toolbar
    For cmds = 1 To 12
        cmd = buttonCmds(cmds)
        Button = buttonNames(cmds)
        state = DHTMLEdit1.QueryStatus(buttonCmds(cmds))
        
        If (state >= DECMDF_ENABLED) Then
            Toolbar1.Buttons(Button).Enabled = True
        Else
            Toolbar1.Buttons(Button).Enabled = False
        End If
            
        If (state = DECMDF_LATCHED) Then
            Toolbar1.Buttons(Button).Value = tbrPressed
        Else
            Toolbar1.Buttons(Button).Value = tbrUnpressed
        End If
    Next cmds
    
    UpdateFontCombos
    
    ' Update the Format menu with the localized strings returned from
    ' the DECMD_GETBLOCKFMT command
    state = DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    If state >= DECMDF_ENABLED Then
        Dim blockFmt As String
        blockFmt = DHTMLEdit1.ExecCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        StatusBar1.Panels(1) = blockFmt
    End If
    
Exit Sub
DHTMLEdit1_DisplayChanged_Error:
    LogErrors Err.Number, Err.Description, Err.Source
    Err.Clear

End Sub
Private Sub InitToolbarTable()
    ' Initialize parallel arrays for mapping
    ' toolbar buttons to DHTMLEdit commands
    
    ' The toolbar buttons are named in the properties
    ' dialog of the toolbar control. We'll use these
    ' names to select on when the user selects a button
    
    buttonNames(1) = "Bold"
    buttonNames(2) = "Italic"
    buttonNames(3) = "Underline"
    buttonNames(4) = "Numbers"
    buttonNames(5) = "Bullets"
    buttonNames(6) = "Outdent"
    buttonNames(7) = "Indent"
    buttonNames(8) = "LeftJustify"
    buttonNames(9) = "Center"
    buttonNames(10) = "RightJustify"
    buttonNames(11) = "Color"
    buttonNames(12) = "BackColor"
    ' This array is parallel to the names array
    ' We'll use the to dispatch a command when the
    ' user selects a button from the toolbar
    
    buttonCmds(1) = DECMD_BOLD
    buttonCmds(2) = DECMD_ITALIC
    buttonCmds(3) = DECMD_UNDERLINE
    buttonCmds(4) = DECMD_ORDERLIST
    buttonCmds(5) = DECMD_UNORDERLIST
    buttonCmds(6) = DECMD_INDENT
    buttonCmds(7) = DECMD_OUTDENT
    buttonCmds(8) = DECMD_JUSTIFYLEFT
    buttonCmds(9) = DECMD_JUSTIFYCENTER
    buttonCmds(10) = DECMD_JUSTIFYRIGHT
    buttonCmds(11) = DECMD_SETFORECOLOR
    buttonCmds(12) = DECMD_SETBACKCOLOR
End Sub
Private Sub InitMenuTables()
    
    
    ' Initialize Insert menu command table
    insertMenuCmds(0) = DECMD_IMAGE
    insertMenuCmds(1) = DECMD_HYPERLINK
    
    ' Initialize Insert menu command table
    tableMenuCmds(0) = DECMD_INSERTTABLE
    tableMenuCmds(1) = 0
    tableMenuCmds(2) = DECMD_INSERTROW
    tableMenuCmds(3) = DECMD_INSERTCOL
    tableMenuCmds(4) = DECMD_INSERTCELL
    tableMenuCmds(5) = 0
    tableMenuCmds(6) = DECMD_DELETEROWS
    tableMenuCmds(7) = DECMD_DELETECOLS
    tableMenuCmds(8) = DECMD_DELETECELLS
    tableMenuCmds(9) = 0
    tableMenuCmds(10) = DECMD_MERGECELLS
    tableMenuCmds(11) = DECMD_SPLITCELL
     
    ' Initialize 2D menu command table
    twoDMenuCmds(0) = DECMD_MAKE_ABSOLUTE
    twoDMenuCmds(1) = DECMD_BRING_TO_FRONT
    twoDMenuCmds(2) = DECMD_SEND_TO_BACK
    twoDMenuCmds(3) = DECMD_BRING_FORWARD
    twoDMenuCmds(4) = DECMD_SEND_BACKWARD
    twoDMenuCmds(5) = DECMD_BRING_ABOVE_TEXT
    twoDMenuCmds(6) = DECMD_SEND_BELOW_TEXT
    twoDMenuCmds(7) = 0
    twoDMenuCmds(8) = DECMD_LOCK_ELEMENT
    End Sub

Private Sub UpdateMenu(menu As Control, command As DHTMLEDITCMDID)

Dim state As DHTMLEDITCMDF

If Not command = 0 Then
    state = DHTMLEdit1.QueryStatus(command)
    
    If (state >= DECMDF_ENABLED) Then
        menu.Enabled = True
    Else
        menu.Enabled = False
    End If
End If

End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key

    Case "new"
        FileNew_Click
        
    Case "open"
        FileOpen_Click
    
    Case "save"
        FileSave_Click
    
    Case "props"
         mnuProperties_Click
    
    Case "html"
         mnuEditHtmlSrc_Click
    Case "revert"
         mnuFileRevert_Click
    
    Case "cut"
         EditCut_Click

    Case "copy"
         EditCopy_Click

    Case "paste"
         EditPaste_Click

    Case "undo"
         EditUndo_Click
    
    Case "redo"
         EditRedo_Click

    Case Else
    
End Select


End Sub
Private Sub ViewSub_Click(Index As Integer)
    Dim state As Boolean
    
    ' Toggle different properties on DHTMLEdit.
    ' Check the menu items if the properties are set
    ' to true
    Select Case Index
        Case 0
            state = DHTMLEdit1.ShowBorders
            state = Not state
            DHTMLEdit1.ShowBorders = state
            ViewSub(Index).Checked = state
        Case 1
            state = DHTMLEdit1.ShowDetails
            state = Not state
            DHTMLEdit1.ShowDetails = state
            ViewSub(Index).Checked = state
    End Select
        
End Sub

Private Sub Format_Click()
    Dim state As DHTMLEDITCMDF
    Dim Format As String
    Dim menuItem As Variant
    
    state = DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    
    If state >= DECMDF_ENABLED Then
        Format = DHTMLEdit1.ExecCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        
        For Each menuItem In FormatSub
            
            ' enable menu item
            menuItem.Enabled = True

            ' Check the menu that reflects the
            ' current formatting
            If menuItem.Caption = Format Then
                menuItem.Checked = True
            Else
                menuItem.Checked = False
            End If
            
        Next
    ElseIf state = DECMDF_DISABLED Then
        ' disable format menu menuItems
        For Each menuItem In FormatSub
            menuItem.Enabled = False
            menuItem.Checked = False
        Next
    End If
End Sub

Private Sub SetFormCaption()

WaitWhileBusy

If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
    MainForm.Caption = APP_TITLE & " - " & DHTMLEdit1.DocumentTitle & " [" & DHTMLEdit1.CurrentDocumentPath & "]"
Else
    MainForm.Caption = APP_TITLE
End If

End Sub
Private Sub DisableToolbar()
    
    FontCombo.Text = ""
    FontCombo.Enabled = False
    FontSizeCombo.Text = ""
    FontSizeCombo.Enabled = False
    
    Dim b As Object
    For Each b In Toolbar1.Buttons
        b.Enabled = False
    Next

    DoEvents 'give toolbar a chance to update itself
End Sub

Private Sub UpdateFontCombos()
On Error GoTo UpdateFontCombos_Error
    Dim state As DHTMLEDITCMDF
    
    ' Update the font name combo box on the toolbar
    state = DHTMLEdit1.QueryStatus(DECMD_GETFONTNAME)
    If state = DECMDF_ENABLED Or state = DECMDF_LATCHED Then
        Dim fontName As String
        fontName = DHTMLEdit1.ExecCommand(DECMD_GETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER)
        FontCombo.Text = fontName
        FontCombo.Enabled = True
    Else
        FontCombo.Text = ""
        If state = DECMDF_NINCHED Then
            FontCombo.Enabled = True
        Else
            FontCombo.Enabled = False
        End If
        
    End If

    ' Update the font size combo box on the toolbar
    state = DHTMLEdit1.QueryStatus(DECMD_GETFONTSIZE)
    If state = DECMDF_ENABLED Or state = DECMDF_LATCHED Then
        Dim fontSize As Long
        fontSize = DHTMLEdit1.ExecCommand(DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER)
        If fontSize >= 1 Then
            FontSizeCombo.Text = fontSize
        Else
            FontSizeCombo.Text = ""
        End If
        FontSizeCombo.Enabled = True
    Else
        FontSizeCombo.Text = ""
        If state = DECMDF_NINCHED Then
            FontSizeCombo.Enabled = True
        Else
            FontSizeCombo.Enabled = False
        End If
    End If
    
Exit Sub
UpdateFontCombos_Error:
    LogErrors Err.Number, Err.Description, Err.Source
    Err.Clear

End Sub
Private Function SaveChanges() As Long
    
    Dim retVal As Long
    If DHTMLEdit1.IsDirty Then
            
        retVal = MsgBox("The current document has changed." & vbCrLf & vbCrLf & "Do you want to save changes?", vbQuestion Or vbYesNoCancel, "Confirm")
    
        Select Case retVal
            Case vbCancel
                SaveChanges = vbCancel
            Case vbYes
                Dim saveSuccess As Boolean
                saveSuccess = False
                If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
                    saveSuccess = SaveDocument(False)
                Else
                    saveSuccess = SaveDocument(True)
                End If
                
                If saveSuccess = True Then
                    SaveChanges = vbOK
                Else
                    SaveChanges = vbCancel
                End If
            
            Case vbNo
                SaveChanges = vbNo
        End Select
    End If
End Function

Private Function SaveDocument(ByVal PromptUser As Boolean) As Boolean

    SaveDocument = True
    
    DisableToolbar
    
    If PromptUser = True Then
        On Error Resume Next
'        MsgBox "*" & GuessFileName() & "*"
        DHTMLEdit1.SaveDocument GuessFileName(), True
    
    Else
        If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
            On Error Resume Next
            DHTMLEdit1.SaveDocument DHTMLEdit1.CurrentDocumentPath
        Else
            Err.Clear
            SaveDocument = False
        End If
    End If
    
    If Err.Number < 0 Then
        Dim errMsg As String
        Select Case Err.Number
            Case DE_E_INVALIDARG
                errMsg = "Invalid argument"
            Case DE_E_PATH_NOT_FOUND
                errMsg = "Path not found"
            Case DE_E_DISK_FULL
                errMsg = "Disk is full"
            Case DE_E_ACCESS_DENIED
                errMsg = "Access denied"
            Case DE_E_UNEXPECTED
                errMsg = "Unexpected error"
            Case Else
                errMsg = "Unknown error"
        End Select
        SaveDocument = False
        MsgBox "Error occurred while saving document: " & errMsg & ".", vbCritical
    End If
        
    On Error Resume Next
    ' Force a DisplayChanged event to update toolbar
    ' in case user canceled file save dialog
    DHTMLEdit1.DOM.selection.createTextRange.collapse
    SetFormCaption
End Function

