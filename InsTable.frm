VERSION 5.00
Begin VB.Form InsertTableDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   2280
   ClientLeft      =   675
   ClientTop       =   1050
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TableCaption 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox CellAttrs 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox TableAttrs 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Cols 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Rows 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton CancelCmd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton OkCmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label CaptionLabel 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label CellTagLabel 
      Caption         =   "Cell Tag Attributes:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label TableTagLabel 
      Caption         =   "Table Tag Attributes:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label ColLabel 
      Caption         =   "Number of columns:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label RowLabel 
      Caption         =   "Number of rows:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "InsertTableDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1999 Microsoft Corporation.
' All rights reserved.
Private tableParam As DEInsertTableParam

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' create the table parameter object
    Set tableParam = CreateObject("DEInsertTableParam.DEInsertTableParam.1")
    
    Rows = tableParam.NumRows
    Cols = tableParam.NumCols
    TableAttrs = tableParam.TableAttrs
    CellAttrs = tableParam.CellAttrs
    TableCaption = tableParam.Caption

End Sub

Private Sub OkCmd_Click()
    
    If Rows = "" Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Rows) = False Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf Rows <= 0 Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    End If
       
    If Cols = "" Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Cols) = False Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf Cols <= 0 Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    End If
    
    tableParam.NumRows = Rows
    tableParam.NumCols = Cols
    
    If Len(TableAttrs.Text) Then
        tableParam.TableAttrs = TableAttrs.Text
    Else
        tableParam.TableAttrs = ""
    End If
    
    If Len(CellAttrs.Text) Then
        tableParam.CellAttrs = CellAttrs.Text
    Else
        tableParam.CellAttrs = ""
    End If
    
    If Len(TableCaption.Text) Then
        tableParam.Caption = TableCaption.Text
    Else
        tableParam.Caption = ""
    End If
    
    MainForm.DHTMLEdit1.ExecCommand DECMD_INSERTTABLE, OLECMDEXECOPT_DONTPROMPTUSER, tableParam
    Unload Me
End Sub

