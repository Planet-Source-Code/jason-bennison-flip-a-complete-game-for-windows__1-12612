VERSION 4.00
Begin VB.Form FrmGuessTheWord 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   1110
   ClientTop       =   7200
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   945
   Left            =   1050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Top             =   6855
   Width           =   9615
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   5280
      TabIndex        =   0
      Top             =   90
      Width           =   3015
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Guess the Word:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "FrmGuessTheWord"
Attribute VB_Creatable = False
Attribute VB_Exposed = False







Private Sub cmdOK_Click()
GWD = TxtName.Text
TxtName.Text = UCase(Trim(""))
FrmGuessTheWord.Hide
End Sub


Private Sub Form_Activate()
FrmGuessTheWord.Left = Flip.Left + 50
FrmGuessTheWord.Top = Flip.Top + Flip.Height - FrmGuessTheWord.Height - 50

End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
Dim VX
VX = Chr(KeyAscii)
KeyAscii = Asc(UCase(VX))
End Sub


