VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KeyCode "
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Main.frx":0442
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "NumLock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblTitleChr 
      Caption         =   "KeyCode:"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTitleAsc 
      Caption         =   "Ascii:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCode 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblASCII 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

lblASCII.Caption = Chr(KeyCode)
lblCode.Caption = KeyCode
    
Select Case KeyCode
Case 112
    lblASCII.Caption = "F1"
    lblCode.Caption = KeyCode
Case 113
    lblASCII.Caption = "F2"
    lblCode.Caption = KeyCode
Case 114
    lblASCII.Caption = "F3"
    lblCode.Caption = KeyCode
Case 115
    lblASCII.Caption = "F4"
    lblCode.Caption = KeyCode
Case 116
    lblASCII.Caption = "F5"
    lblCode.Caption = KeyCode
Case 117
    lblASCII.Caption = "F6"
    lblCode.Caption = KeyCode
Case 118
    lblASCII.Caption = "F7"
    lblCode.Caption = KeyCode
Case 119
    lblASCII.Caption = "F8"
    lblCode.Caption = KeyCode
Case 120
    lblASCII.Caption = "F9"
    lblCode.Caption = KeyCode
Case 121
    lblASCII.Caption = "F10"
    lblCode.Caption = KeyCode
Case 122
    lblASCII.Caption = "F11"
    lblCode.Caption = KeyCode
Case 123
    lblASCII.Caption = "F12"
    lblCode.Caption = KeyCode
Case 145
    lblASCII.Caption = "Scroll Lock"
    lblCode.Caption = KeyCode
Case 19
    lblASCII.Caption = "Pause"
    lblCode.Caption = KeyCode
Case 9
    lblASCII.Caption = "{TAB}"
    lblCode.Caption = KeyCode
Case 20
    lblASCII.Caption = "Caps Lock"
    lblCode.Caption = KeyCode
Case 16
    lblASCII.Caption = "Shift"
    lblCode.Caption = KeyCode
Case 18
    lblASCII.Caption = "Alt"
    lblCode.Caption = KeyCode
Case 17
    lblASCII.Caption = "Control"
    lblCode.Caption = KeyCode
Case 32
    lblASCII.Caption = "{Space}"
    lblCode.Caption = KeyCode
Case 13
    lblASCII.Caption = "Enter"
    lblCode.Caption = KeyCode
Case 8
    lblASCII.Caption = "{Back Space}"
    lblCode.Caption = KeyCode
Case 45
    lblASCII.Caption = "Insert"
    lblCode.Caption = KeyCode
Case 36
    lblASCII.Caption = "Home"
    lblCode.Caption = KeyCode
Case 33
    lblASCII.Caption = "PgUp"
    lblCode.Caption = KeyCode
Case 46
    lblASCII.Caption = "Delete"
    lblCode.Caption = KeyCode
Case 35
    lblASCII.Caption = "End"
    lblCode.Caption = KeyCode
Case 34
    lblASCII.Caption = "PgDn"
    lblCode.Caption = KeyCode
Case 38
    lblASCII.Caption = "{UP}"
    lblCode.Caption = KeyCode
Case 37
    lblASCII.Caption = "{LEFT}"
    lblCode.Caption = KeyCode
Case 39
    lblASCII.Caption = "{RIGHT}"
    lblCode.Caption = KeyCode
Case 40
    lblASCII.Caption = "{DOWN}"
    lblCode.Caption = KeyCode
Case 103
    lblASCII.Caption = "7"
    lblCode.Caption = KeyCode
Case 27
    lblASCII.Caption = "Esc"
    lblCode.Caption = KeyCode
Case 144
    lblASCII.Caption = "NumLock"
    lblCode.Caption = KeyCode
Case 111
    lblASCII.Caption = "/"
    lblCode.Caption = KeyCode
Case 106
    lblASCII.Caption = "*"
    lblCode.Caption = KeyCode
Case 109
    lblASCII.Caption = "-"
    lblCode.Caption = KeyCode
Case 104
    lblASCII.Caption = "8"
    lblCode.Caption = KeyCode
Case 105
    lblASCII.Caption = "9"
    lblCode.Caption = KeyCode
Case 100
    lblASCII.Caption = "4"
    lblCode.Caption = KeyCode
Case 101
    lblASCII.Caption = "5"
    lblCode.Caption = KeyCode
Case 102
    lblASCII.Caption = "6"
    lblCode.Caption = KeyCode
Case 97
    lblASCII.Caption = "1"
    lblCode.Caption = KeyCode
Case 98
    lblASCII.Caption = "2"
    lblCode.Caption = KeyCode
Case 99
    lblASCII.Caption = "3"
    lblCode.Caption = KeyCode
Case 96
    lblASCII.Caption = "0"
    lblCode.Caption = KeyCode
Case 110
    lblASCII.Caption = "."
    lblCode.Caption = KeyCode
Case 107
    lblASCII.Caption = "+"
    lblCode.Caption = KeyCode
Case 12
    lblASCII.Caption = "5"
    lblCode.Caption = KeyCode
Case Else
    lblASCII.Caption = Chr(KeyCode)
    lblCode.Caption = KeyCode
End Select

Select Case KeyCode
Case 96 To 111, 12
    Image1.Visible = True
    Label1.Visible = True
    Label1.Caption = "figure block"
Case 144
    Image1.Visible = True
    Label1.Visible = True
    Label1.Caption = "NumLock"
Case Else
    Image1.Visible = False
    Label1.Visible = False
End Select
End Sub

