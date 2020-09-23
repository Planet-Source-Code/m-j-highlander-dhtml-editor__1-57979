VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmViewSelHtml 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML"
   ClientHeight    =   4785
   ClientLeft      =   2745
   ClientTop       =   3075
   ClientWidth     =   7545
   Icon            =   "ViewSelHtml.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7545
   Begin RichTextLib.RichTextBox rtfHTML 
      Height          =   3915
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6906
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"ViewSelHtml.frx":000C
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Clipboard"
      Height          =   465
      Left            =   3990
      TabIndex        =   1
      Top             =   4200
      Width           =   1605
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5820
      TabIndex        =   0
      Top             =   4200
      Width           =   1605
   End
End
Attribute VB_Name = "frmViewSelHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdCopy_Click()

If Trim(rtfHTML.Text) <> "" Then

    Clipboard.Clear
    Clipboard.SetText rtfHTML.Text, vbCFText

End If


End Sub

Private Sub Form_Load()

Me.Icon = MainForm.Icon

End Sub
