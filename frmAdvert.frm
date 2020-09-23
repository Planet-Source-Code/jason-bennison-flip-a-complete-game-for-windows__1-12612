VERSION 4.00
Begin VB.Form frmAdvert 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ordering Flip!"
   ClientHeight    =   6030
   ClientLeft      =   195
   ClientTop       =   555
   ClientWidth     =   9210
   ControlBox      =   0   'False
   Height          =   6435
   Icon            =   "frmAdvert.frx":0000
   Left            =   135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9210
   Top             =   210
   Width           =   9330
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   120
      Picture         =   "frmAdvert.frx":030A
      ScaleHeight     =   630
      ScaleWidth      =   1290
      TabIndex        =   6
      Top             =   3360
      Width           =   1290
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   120
      Picture         =   "frmAdvert.frx":15BC
      ScaleHeight     =   630
      ScaleWidth      =   1290
      TabIndex        =   5
      Top             =   4080
      Width           =   1290
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   120
      Picture         =   "frmAdvert.frx":286E
      ScaleHeight     =   630
      ScaleWidth      =   1290
      TabIndex        =   4
      Top             =   4800
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load Order Form"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      HelpContextID   =   10
      Left            =   3480
      TabIndex        =   3
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return to Flip!"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Flip!"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   120
      X2              =   7800
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To order Flip!, Follow these 3 simple steps -"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   9015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3. Play Flip! and enjoy!."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   5040
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2. When you get your Validation Card, follow the instructions on the Card."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4320
      Width           =   7455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1. Fill in and send the Order Form with payment to the address on the form."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   7455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAdvert.frx":3B20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   9015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ordering Flip! is Simple."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmAdvert"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
frmAdvert.Hide
End Sub


Private Sub cmdReturn_Click()
JJ = 1
JK = 0
frmAdvert.Hide
End Sub


Private Sub Command1_Click()
CommonDialog1.HelpFile = "Info.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub

Private Sub Form_Activate()
frmAdvert.Top = (Screen.Height / 2) - (frmAdvert.Height / 2)
frmAdvert.Left = (Screen.Width / 2) - (frmAdvert.Width / 2)
End Sub

Private Sub Form_Paint()
Static Fd As Integer
If Fd <> 1 Then
Fd = 1
End If
End Sub


