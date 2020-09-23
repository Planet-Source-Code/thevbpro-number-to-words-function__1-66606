VERSION 5.00
Begin VB.Form frmNumberToWords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert a Number to Words"
   ClientHeight    =   4920
   ClientLeft      =   1350
   ClientTop       =   1590
   ClientWidth     =   6585
   Icon            =   "Num2Word.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3780
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "&Do It"
      Default         =   -1  'True
      Height          =   495
      Left            =   300
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox txtNumber 
      Height          =   435
      Left            =   360
      MaxLength       =   36
      TabIndex        =   0
      Top             =   600
      Width           =   5835
   End
   Begin VB.Label lblNumToWords 
      BorderStyle     =   1  'Fixed Single
      Height          =   2595
      Left            =   300
      TabIndex        =   3
      Top             =   1500
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Number in words:"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   300
      Width           =   5775
   End
End
Attribute VB_Name = "frmNumberToWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtNumber_GotFocus()
    With txtNumber
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then Exit Sub
    
    If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtNumber_Change()
    lblNumToWords = ""
End Sub

Private Sub cmdDoIt_Click()
    lblNumToWords = NumToWords(txtNumber)
End Sub

Private Sub cmdExit_Click()
    End
End Sub
