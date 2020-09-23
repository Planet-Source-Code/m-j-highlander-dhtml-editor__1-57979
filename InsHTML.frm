VERSION 5.00
Begin VB.Form InsertHTMLDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert HTML"
   ClientHeight    =   4395
   ClientLeft      =   2550
   ClientTop       =   2400
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "InsHTML.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   3870
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3870
      Width           =   1215
   End
   Begin VB.TextBox HTMLText 
      Height          =   3195
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "InsHTML.frx":000C
      Top             =   510
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "HTML source:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "InsertHTMLDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1999 Microsoft Corporation.
' All rights reserved.
Private Sub CmdCancel_Click()
    Unload Me
    MainForm.DHTMLEdit1.SetFocus
End Sub

Private Sub cmdOK_Click()

    Dim doc As Object
    Dim sel As Object
    Dim tr As Object
    
    ' get the DHTML Document object
    Set doc = MainForm.DHTMLEdit1.DOM
    ' get the IE4 selection object
    Set sel = doc.selection
    ' create a TextRange from the current selection
    Set tr = sel.createrange
    
    ' paste our html into the range
    tr.pasteHTML (HTMLText.Text)
    Unload Me
End Sub

Private Sub Form_Activate()

HTMLText.SelStart = 0
HTMLText.SelLength = Len(HTMLText.Text)
HTMLText.SetFocus


End Sub
