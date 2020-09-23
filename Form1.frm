VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   420
      Left            =   45
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1185
      Width           =   2400
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   75
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   690
      Width           =   2370
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   195
      Width           =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function BkColor(txt As TextBox)
Dim Control As Control
'call from the GotFocus property of each textbox
'ex: BkColor Text1

For Each Control In Me
If TypeOf Control Is TextBox Then Control.BackColor = vbWhite
Next Control
txt.BackColor = &HC0FFFF
End Function

Private Sub Text1_GotFocus()
BkColor Text1
End Sub

Private Sub Text2_GotFocus()
BkColor Text2
End Sub

Private Sub Text3_GotFocus()
BkColor Text3
End Sub

