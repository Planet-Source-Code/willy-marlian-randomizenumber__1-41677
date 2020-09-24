VERSION 5.00
Begin VB.Form frmRandomizeNumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Randomize"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtMaxNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   360
      TabIndex        =   2
      Text            =   "10"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtNumber 
      Height          =   4575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdRandomize 
      Caption         =   "Randomize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmRandomizeNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdHelp_Click()
  MsgBox "Randomize & Print Unique Number" & vbCr & _
    "---------------------------------------------" & vbCr & _
    "This Program Created by Willy Marlian" & vbCr & _
    "Christmas Edition 2002"
End Sub

Private Sub cmdRandomize_Click()
  txtNumber = "Max: " & txtMaxNumber
  RandomizeNumber txtMaxNumber
  'txtNumber = txtNumber & vbCrLf & 1
End Sub

Private Sub RandomizeNumber(intMaxNumber As Integer)
'Randomize & Print Unique Number
Dim N(999) As Integer
Dim i, j, k As Integer
  ' Begin - Pengisian Array N Secara Acak
  Randomize
  N(1) = Int(Rnd * intMaxNumber) + 1
  i = 2
  While i <= intMaxNumber
    j = 1
    N(i) = Int(Rnd * intMaxNumber) + 1
    While j < i
      If N(i) = N(j) Then
        N(i) = Int(Rnd * intMaxNumber) + 1
        j = 1
      Else
        j = j + 1
      End If
    Wend
    i = i + 1
  Wend
  ' End - Pengisian Array N Secara Acak

  ' Print
  For i = 1 To intMaxNumber
    txtNumber = txtNumber & vbCrLf & N(i)
  Next i
End Sub
