VERSION 4.00
Begin VB.Form Flip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6900
   ClientLeft      =   1125
   ClientTop       =   2265
   ClientWidth     =   9615
   ClipControls    =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   7590
   Icon            =   "Flip.frx":0000
   Left            =   1065
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   6900
   ScaleWidth      =   9615
   Top             =   1635
   Width           =   9735
   Begin VB.Timer Timer5 
      Interval        =   3000
      Left            =   4680
      Top             =   2640
   End
   Begin VB.PictureBox pic2player 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3240
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   116
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Pic1player 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2520
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   115
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Piccols 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6960
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   112
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox PicCards 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6240
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   111
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picMov 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4080
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   110
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox PicFlag 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4800
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   109
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox PicNums 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5520
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   108
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox PicRS 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7800
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   107
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10750
      Left            =   4320
      Top             =   2640
   End
   Begin VB.PictureBox Pic3x8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   105
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox pic4x10 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   960
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   104
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox pic6x12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   103
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11520
      TabIndex        =   102
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   19
      Left            =   3720
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   100
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox Picture10 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8520
      Picture         =   "Flip.frx":030A
      ScaleHeight     =   450
      ScaleWidth      =   735
      TabIndex        =   99
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   42
      Left            =   3240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   87
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "On"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   86
      Top             =   4320
      Width           =   375
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   67
      Left            =   8400
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   85
      Top             =   3600
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Interval        =   1825
      Left            =   3960
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   3240
      Top             =   2640
   End
   Begin VB.TextBox Text3 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   84
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fault On/Off"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   83
      Top             =   5760
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3600
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   82
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   66
      Left            =   7680
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   71
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   59
      Left            =   6600
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   70
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   68
      Left            =   5280
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   69
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   69
      Left            =   4320
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   68
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   70
      Left            =   4560
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   67
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   64
      Left            =   6120
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   66
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5400
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   65
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   72
      Left            =   8640
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   64
      Top             =   2760
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   71
      Left            =   8640
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   63
      Top             =   3240
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   65
      Left            =   6240
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   62
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   53
      Left            =   8400
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   61
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   48
      Left            =   6600
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   60
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   60
      Left            =   1920
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   59
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   62
      Left            =   8400
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   58
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   63
      Left            =   5520
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   57
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   49
      Left            =   6480
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   56
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   46
      Left            =   3000
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   55
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   2400
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   54
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   2280
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   53
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   3000
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   52
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   3720
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   51
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   600
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   50
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   1200
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   49
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   1080
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   48
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   24
      Left            =   600
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   47
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   61
      Left            =   8400
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   46
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1440
      MousePointer    =   2  'Cross
      Negotiate       =   -1  'True
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   45
      Tag             =   "1"
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   4440
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   44
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   36
      Left            =   960
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   43
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   25
      Left            =   960
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   42
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   960
      MousePointer    =   2  'Cross
      Negotiate       =   -1  'True
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   41
      Tag             =   "2"
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   34
      Left            =   4440
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   40
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2280
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   39
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   58
      Left            =   960
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   38
      Top             =   3480
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   720
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   37
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   720
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   36
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   47
      Left            =   6120
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   35
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   37
      Left            =   5160
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   34
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   57
      Left            =   7320
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   33
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   44
      Left            =   7680
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   32
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   32
      Left            =   360
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   45
      Left            =   6120
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   33
      Left            =   960
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   29
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   5880
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   28
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   5880
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   27
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   20
      Left            =   720
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   26
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   1080
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   25
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   840
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   24
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   960
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   23
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   840
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   22
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   54
      Left            =   480
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   30
      Left            =   7920
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   20
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   55
      Left            =   8280
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   19
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   43
      Left            =   5280
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   18
      Top             =   2760
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   31
      Left            =   720
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   17
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   56
      Left            =   5640
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   39
      Left            =   5520
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   51
      Left            =   5400
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   14
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   40
      Left            =   1440
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   52
      Left            =   5760
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   12
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   28
      Left            =   960
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   840
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   41
      Left            =   -240
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   29
      Left            =   360
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   26
      Left            =   5160
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   38
      Left            =   360
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   50
      Left            =   5880
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   27
      Left            =   480
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox G 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   35
      Left            =   480
      MousePointer    =   2  'Cross
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtWord 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1635
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox TxtBackground 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2640
      Width           =   600
   End
   Begin VB.Shape shpplay2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Height          =   525
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   4095
   End
   Begin VB.Shape Shpplay1 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Height          =   525
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   4215
   End
   Begin VB.Label lblplayer2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   5520
      TabIndex        =   114
      Top             =   5640
      Width           =   105
   End
   Begin VB.Label lblplayer1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   360
      TabIndex        =   113
      Top             =   5640
      Width           =   105
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpContext     =   1
      HelpFile        =   "flip"
      HelpKey         =   "F1"
   End
   Begin VB.Label lblNana 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Left            =   240
      TabIndex        =   106
      Top             =   5640
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.OLE OLE1 
      Class           =   "midfile"
      Height          =   495
      Left            =   1800
      OleObjectBlob   =   "Flip.frx":06D4
      TabIndex        =   101
      Top             =   1920
      Width           =   30
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1320
      TabIndex        =   98
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label Lbl3 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1560
      TabIndex        =   97
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label lbl5 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2040
      TabIndex        =   96
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1800
      TabIndex        =   95
      Top             =   4440
      Width           =   150
   End
   Begin VB.Label lbl6 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2280
      TabIndex        =   94
      Top             =   4680
      Width           =   150
   End
   Begin VB.Label lbl7 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2520
      TabIndex        =   93
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label lbl8 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2760
      TabIndex        =   92
      Top             =   4680
      Width           =   150
   End
   Begin VB.Label lbl9 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3000
      TabIndex        =   91
      Top             =   4680
      Width           =   150
   End
   Begin VB.Label lbl10 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "10"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3240
      TabIndex        =   90
      Top             =   4560
      Width           =   300
   End
   Begin VB.Label lbl11 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "11"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3600
      TabIndex        =   89
      Top             =   4560
      Width           =   300
   End
   Begin VB.Label lbl12 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "12"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3960
      TabIndex        =   88
      Top             =   4560
      Width           =   300
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   480
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   49
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":24EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3332
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4178
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4552
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":492C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":50E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":54BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5894
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6048
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6422
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":67FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":738A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7764
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":82F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":86CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":925A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTries 
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   81
      Top             =   5760
      Width           =   2895
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   49
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   397
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9634
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A3E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B210
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C036
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DC86
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":EAAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F8D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":106F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":114AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1225C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1300E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15924
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":166D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":17488
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":182CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":19114
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":19F5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1ADA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1BBE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1CA2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1D872
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1E6B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1F4FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":20344
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2118A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":21FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":22E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":23C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":24AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":258E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2672E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":27574
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":283BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":29200
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2A046
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2AE8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2BCD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2CB18
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2D8CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2E688
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":2F442
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":301F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":30FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":31D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":32B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":33938
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3477E
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":355C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3640A
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":37250
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":38096
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":38EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":39D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3AB68
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3B9AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3C7F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3D63A
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3E480
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":3F2C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4010C
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":40F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":41D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":42BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":43A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4486A
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":456B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":464F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4733C
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":48182
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":48FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":49E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4AC54
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4BA9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4C8E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4D726
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4E56C
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":4F3B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":501F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5103E
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":51E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":52CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":53B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":54956
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5579C
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":565E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":57428
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5826E
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":590B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":59EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5AD40
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5BB86
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5C9CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5D812
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5E658
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":5F49E
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":602E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6112A
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":61F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":62DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":63BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":64A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":65888
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":666CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":67514
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6835A
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":691A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":69FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6AE2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6BC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6CAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6D882
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6E644
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":6F40A
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":701D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":70F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":71D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":72B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":738C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7468E
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":75454
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7621A
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":76FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":77DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":78B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":79932
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7A6F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7B53E
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7C384
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7D1CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7E010
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7EE56
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":7FC9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":80AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":81928
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8276E
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":835B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":843FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":85240
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":86086
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":86ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":87D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":88B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8999E
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8A7E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8B62A
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8C470
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8D2B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8E0FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8EF42
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":8FD88
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":90BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":91A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9285A
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":936A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":944E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9532C
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":96172
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":96FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":97D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":98B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":998F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9A6B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9B46E
            Key             =   ""
         EndProperty
         BeginProperty ListImage167 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9C22C
            Key             =   ""
         EndProperty
         BeginProperty ListImage168 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9CFEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage169 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9DDA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage170 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9EBD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage171 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":9FA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage172 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A083E
            Key             =   ""
         EndProperty
         BeginProperty ListImage173 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage174 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A2496
            Key             =   ""
         EndProperty
         BeginProperty ListImage175 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A32BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage176 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A40EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage177 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A4F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage178 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A5D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage179 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A6B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage180 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A79DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage181 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A8820
            Key             =   ""
         EndProperty
         BeginProperty ListImage182 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":A9662
            Key             =   ""
         EndProperty
         BeginProperty ListImage183 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":AA4A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage184 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":AB2E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage185 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":AC128
            Key             =   ""
         EndProperty
         BeginProperty ListImage186 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":ACF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage187 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":ADDA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage188 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":AEBEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage189 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":AFA28
            Key             =   ""
         EndProperty
         BeginProperty ListImage190 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B086A
            Key             =   ""
         EndProperty
         BeginProperty ListImage191 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B16AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage192 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B24EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage193 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B3330
            Key             =   ""
         EndProperty
         BeginProperty ListImage194 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B40EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage195 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B4EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage196 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B5C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage197 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B6A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage198 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B77D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage199 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B858C
            Key             =   ""
         EndProperty
         BeginProperty ListImage200 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":B9346
            Key             =   ""
         EndProperty
         BeginProperty ListImage201 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BA100
            Key             =   ""
         EndProperty
         BeginProperty ListImage202 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BAF32
            Key             =   ""
         EndProperty
         BeginProperty ListImage203 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BBD64
            Key             =   ""
         EndProperty
         BeginProperty ListImage204 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BCB96
            Key             =   ""
         EndProperty
         BeginProperty ListImage205 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BD9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage206 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BE7FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage207 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":BF62C
            Key             =   ""
         EndProperty
         BeginProperty ListImage208 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C045E
            Key             =   ""
         EndProperty
         BeginProperty ListImage209 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C1290
            Key             =   ""
         EndProperty
         BeginProperty ListImage210 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C20C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage211 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C2EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage212 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C3D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage213 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C4B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage214 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C598A
            Key             =   ""
         EndProperty
         BeginProperty ListImage215 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C67BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage216 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C75EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage217 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C8420
            Key             =   ""
         EndProperty
         BeginProperty ListImage218 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":C9252
            Key             =   ""
         EndProperty
         BeginProperty ListImage219 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CA084
            Key             =   ""
         EndProperty
         BeginProperty ListImage220 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CAEB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage221 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CBCE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage222 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CCB1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage223 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CD94C
            Key             =   ""
         EndProperty
         BeginProperty ListImage224 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CE77E
            Key             =   ""
         EndProperty
         BeginProperty ListImage225 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":CF5B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage226 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D03F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage227 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D1234
            Key             =   ""
         EndProperty
         BeginProperty ListImage228 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D2076
            Key             =   ""
         EndProperty
         BeginProperty ListImage229 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D2EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage230 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage231 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D4B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage232 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D597E
            Key             =   ""
         EndProperty
         BeginProperty ListImage233 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D67B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage234 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D75F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage235 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D843C
            Key             =   ""
         EndProperty
         BeginProperty ListImage236 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":D9282
            Key             =   ""
         EndProperty
         BeginProperty ListImage237 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DA0C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage238 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DAF0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage239 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DBD54
            Key             =   ""
         EndProperty
         BeginProperty ListImage240 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DCB9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage241 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DD9E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage242 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DE806
            Key             =   ""
         EndProperty
         BeginProperty ListImage243 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":DF62C
            Key             =   ""
         EndProperty
         BeginProperty ListImage244 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage245 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E1278
            Key             =   ""
         EndProperty
         BeginProperty ListImage246 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E209E
            Key             =   ""
         EndProperty
         BeginProperty ListImage247 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E2EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage248 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E3CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage249 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E4B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage250 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E5936
            Key             =   ""
         EndProperty
         BeginProperty ListImage251 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E675C
            Key             =   ""
         EndProperty
         BeginProperty ListImage252 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E7582
            Key             =   ""
         EndProperty
         BeginProperty ListImage253 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E83A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage254 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E91CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage255 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":E9FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage256 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":EAE1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage257 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":EBC40
            Key             =   ""
         EndProperty
         BeginProperty ListImage258 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":ECA76
            Key             =   ""
         EndProperty
         BeginProperty ListImage259 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":ED8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage260 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":EE6E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage261 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":EF518
            Key             =   ""
         EndProperty
         BeginProperty ListImage262 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage263 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F1184
            Key             =   ""
         EndProperty
         BeginProperty ListImage264 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F1FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage265 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F2DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage266 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F3BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage267 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F4974
            Key             =   ""
         EndProperty
         BeginProperty ListImage268 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F5736
            Key             =   ""
         EndProperty
         BeginProperty ListImage269 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F64F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage270 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F72BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage271 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F807C
            Key             =   ""
         EndProperty
         BeginProperty ListImage272 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F8E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage273 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":F9C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage274 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FAA46
            Key             =   ""
         EndProperty
         BeginProperty ListImage275 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FB88C
            Key             =   ""
         EndProperty
         BeginProperty ListImage276 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FC6D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage277 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FD518
            Key             =   ""
         EndProperty
         BeginProperty ListImage278 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FE35E
            Key             =   ""
         EndProperty
         BeginProperty ListImage279 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FF1A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage280 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":FFFEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage281 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":100E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage282 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":101C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage283 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":102A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage284 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1038AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage285 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1046D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage286 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":105502
            Key             =   ""
         EndProperty
         BeginProperty ListImage287 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10632C
            Key             =   ""
         EndProperty
         BeginProperty ListImage288 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":107156
            Key             =   ""
         EndProperty
         BeginProperty ListImage289 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":107F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage290 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":108DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage291 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":109BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage292 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10AA0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage293 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10B838
            Key             =   ""
         EndProperty
         BeginProperty ListImage294 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10C666
            Key             =   ""
         EndProperty
         BeginProperty ListImage295 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10D494
            Key             =   ""
         EndProperty
         BeginProperty ListImage296 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10E2C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage297 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10F0F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage298 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":10FF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage299 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":110D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage300 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":111B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage301 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1129A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage302 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1137D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage303 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":114604
            Key             =   ""
         EndProperty
         BeginProperty ListImage304 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":115432
            Key             =   ""
         EndProperty
         BeginProperty ListImage305 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":116260
            Key             =   ""
         EndProperty
         BeginProperty ListImage306 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11708E
            Key             =   ""
         EndProperty
         BeginProperty ListImage307 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":117EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage308 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":118CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage309 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":119B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage310 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11A946
            Key             =   ""
         EndProperty
         BeginProperty ListImage311 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11B774
            Key             =   ""
         EndProperty
         BeginProperty ListImage312 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11C5A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage313 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11D3D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage314 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11E1FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage315 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11F02C
            Key             =   ""
         EndProperty
         BeginProperty ListImage316 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":11FE5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage317 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":120C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage318 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":121AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage319 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1228E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage320 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":123712
            Key             =   ""
         EndProperty
         BeginProperty ListImage321 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":124540
            Key             =   ""
         EndProperty
         BeginProperty ListImage322 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12536E
            Key             =   ""
         EndProperty
         BeginProperty ListImage323 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12619C
            Key             =   ""
         EndProperty
         BeginProperty ListImage324 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":126FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage325 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":127DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage326 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12824E
            Key             =   ""
         EndProperty
         BeginProperty ListImage327 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":129080
            Key             =   ""
         EndProperty
         BeginProperty ListImage328 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":129EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage329 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12ACE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage330 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12BB16
            Key             =   ""
         EndProperty
         BeginProperty ListImage331 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12C948
            Key             =   ""
         EndProperty
         BeginProperty ListImage332 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12D77A
            Key             =   ""
         EndProperty
         BeginProperty ListImage333 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12E5AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage334 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":12F3DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage335 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":130210
            Key             =   ""
         EndProperty
         BeginProperty ListImage336 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":131042
            Key             =   ""
         EndProperty
         BeginProperty ListImage337 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":131E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage338 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":132CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage339 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":133AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage340 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13490A
            Key             =   ""
         EndProperty
         BeginProperty ListImage341 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13573C
            Key             =   ""
         EndProperty
         BeginProperty ListImage342 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13656E
            Key             =   ""
         EndProperty
         BeginProperty ListImage343 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1373A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage344 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1381D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage345 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":139004
            Key             =   ""
         EndProperty
         BeginProperty ListImage346 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":139E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage347 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13AC68
            Key             =   ""
         EndProperty
         BeginProperty ListImage348 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13BA9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage349 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13C8CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage350 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13D6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage351 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13E530
            Key             =   ""
         EndProperty
         BeginProperty ListImage352 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":13F362
            Key             =   ""
         EndProperty
         BeginProperty ListImage353 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":140194
            Key             =   ""
         EndProperty
         BeginProperty ListImage354 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":140FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage355 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":141DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage356 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":142C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage357 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":143A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage358 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14488E
            Key             =   ""
         EndProperty
         BeginProperty ListImage359 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1456C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage360 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1464F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage361 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":147324
            Key             =   ""
         EndProperty
         BeginProperty ListImage362 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":148156
            Key             =   ""
         EndProperty
         BeginProperty ListImage363 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":148F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage364 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":149DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage365 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14ABEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage366 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14BA32
            Key             =   ""
         EndProperty
         BeginProperty ListImage367 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14C878
            Key             =   ""
         EndProperty
         BeginProperty ListImage368 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14D6AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage369 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14E4DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage370 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":14F322
            Key             =   ""
         EndProperty
         BeginProperty ListImage371 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":150168
            Key             =   ""
         EndProperty
         BeginProperty ListImage372 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":150FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage373 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":151DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage374 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":152C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage375 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":153A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage376 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15489E
            Key             =   ""
         EndProperty
         BeginProperty ListImage377 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1556D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage378 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":156502
            Key             =   ""
         EndProperty
         BeginProperty ListImage379 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":157334
            Key             =   ""
         EndProperty
         BeginProperty ListImage380 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":158166
            Key             =   ""
         EndProperty
         BeginProperty ListImage381 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":158F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage382 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":159DCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage383 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15ABFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage384 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15BA2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage385 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15C860
            Key             =   ""
         EndProperty
         BeginProperty ListImage386 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15D692
            Key             =   ""
         EndProperty
         BeginProperty ListImage387 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15E4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage388 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":15F31A
            Key             =   ""
         EndProperty
         BeginProperty ListImage389 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":160160
            Key             =   ""
         EndProperty
         BeginProperty ListImage390 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":160FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage391 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":161DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage392 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":162C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage393 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":163A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage394 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1648B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage395 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":1656F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage396 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":16653E
            Key             =   ""
         EndProperty
         BeginProperty ListImage397 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Flip.frx":167384
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape InnerBord 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   30
      Left            =   960
      Top             =   1800
      Width           =   5790
   End
   Begin VB.Shape ToolBar 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   525
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   9495
   End
   Begin VB.Label lblD 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   720
      TabIndex        =   80
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblE 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   480
      TabIndex        =   79
      Top             =   4560
      Width           =   195
   End
   Begin VB.Label lblF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "F"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   720
      TabIndex        =   78
      Top             =   4560
      Width           =   195
   End
   Begin VB.Label Label8 
      Caption         =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   -360
      TabIndex        =   77
      Top             =   2760
      Width           =   210
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   480
      TabIndex        =   76
      Top             =   3840
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   -240
      TabIndex        =   75
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   720
      TabIndex        =   74
      Top             =   3840
      Width           =   195
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   480
      TabIndex        =   73
      Top             =   4200
      Width           =   210
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1095
      TabIndex        =   72
      Top             =   4440
      Width           =   150
   End
   Begin VB.Shape Surround 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   6270
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuShortTunes 
         Caption         =   "Short Tunes"
      End
      Begin VB.Menu mnuNoTunes 
         Caption         =   "No Tunes"
      End
      Begin VB.Menu mnucontinuous 
         Caption         =   "Timed Game"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisableWilds 
         Caption         =   "Disable Wild Cards"
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHiScore 
         Caption         =   "Hi-Score &Table"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSize 
      Caption         =   "&Skill Level"
      Begin VB.Menu mnu3X8 
         Caption         =   "&Beginner       (3 x 8)"
      End
      Begin VB.Menu mnu4X10 
         Caption         =   "&Intermediate  (4 x 10)"
      End
      Begin VB.Menu mnu5X12 
         Caption         =   "&Expert           (6 x 12)"
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu1Player 
         Caption         =   "1 Player Game"
      End
      Begin VB.Menu mnu2Player 
         Caption         =   "2 Player Game"
      End
   End
   Begin VB.Menu MnuObjects 
      Caption         =   "&Objects"
      Begin VB.Menu MnuAnything 
         Caption         =   "&Anything that moves"
      End
      Begin VB.Menu mnuFlags 
         Caption         =   "Flags of the &World"
      End
      Begin VB.Menu mnuNumbers 
         Caption         =   "N&umbers"
      End
      Begin VB.Menu mnuPlayCards 
         Caption         =   "&Playing Cards   "
      End
      Begin VB.Menu mnuColors 
         Caption         =   "&Colors       (Fixed)"
      End
      Begin VB.Menu mnuColors2 
         Caption         =   "&Colors       (Randomized)"
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserWord 
         Caption         =   "User Word"
      End
      Begin VB.Menu mnuComputerWord 
         Caption         =   "Computer Generated Words"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuShowHelp 
         Caption         =   "Contents     F1"
      End
      Begin VB.Menu mnuOrdering 
         Caption         =   "Ordering Info."
      End
      Begin VB.Menu mnuLoadOrder 
         Caption         =   "Load Order Form"
      End
      Begin VB.Menu mnuValidate 
         Caption         =   "&Validate"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MnuGuess 
         Caption         =   "show guess"
      End
   End
End
Attribute VB_Name = "Flip"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer
Dim AB As Integer
Dim AST As Integer
Dim AX As Integer
Dim C As Integer
Dim CH As Integer
Dim CI(1 To 37) As Variant
Dim CJ(1 To 14) As Variant
Dim CK(1 To 14) As Variant
Dim CL As Integer
Dim CT As Integer
Dim CW As Integer
Dim CX As Integer
Dim D(1 To 72)
Dim DB(1 To 302) As Variant
Dim DC(1 To 302) As Variant
Dim DD(1 To 302) As Variant
Dim DE(1 To 302) As Variant
Dim DL As Integer
Dim DM As Integer
Dim DN As Integer
Dim DW As Integer
Dim DX As Integer
Dim DY As Integer
Dim DZ As Integer
Dim E(1 To 72)
Dim EX As Integer
Dim F(1 To 72)
Dim FT(1 To 72)
Dim FU(1 To 72)
Dim FV(1 To 72)
Dim FW As Integer
Dim GA1 As Integer
Dim GA2 As Integer
Dim GO As Integer
Dim GW As Integer
Dim H As Integer
Dim HAA
Dim HK As Integer
Dim I As Integer
Dim IA As Integer
Dim IB As Integer
Dim IC As Integer
Dim J As Integer
Dim K(1 To 72)
Dim L As Integer
Dim LT As Integer
Dim M As Integer
Dim MHD As Integer
Dim N(1 To 72)
Dim O As Integer
Dim OPT As Integer
Dim P As Integer
Dim PA(1 To 4)
Dim Player1 As String * 10
Dim Player2 As String * 10
Dim PL1 As Integer
Dim PL2 As Integer
Dim PX As Integer
Dim Q(0 To 104)
Dim R As Integer
Dim S(1 To 72)
Dim SA As Integer
Dim SE As Integer
Dim SP As Integer
Dim ST As Integer
Dim SW As Integer
Dim T(1 To 72)
Dim U(1 To 72)
Dim UW As Integer
Dim V As Integer
Dim W As Integer
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim SX As Integer
Dim Word As String
Dim Word1 As String
Dim Word2 As String
Dim Tem As Integer
Dim ZX(1 To 72)

Private Sub CmdNewGame_Click()
ChangeLevels
End Sub


Private Sub Command2_Click()
Tem = Tem + 1
If Tem >= 2 Then
Tem = 0
End If
If Tem = 1 Then
Command2.Caption = "ON"
If R = 1 Or R = 2 Then
Q(5) = 1
Else
Q(19) = 1
End If
End If
If Tem = 0 Then
Command2.Caption = "OFF"
If R = 1 Or R = 2 Then
Q(19) = 0
Else
Q(19) = 0
End If
End If
End Sub


Private Sub Command3_Click()
For J = 1 To 72
Q(J) = 2
Next
End Sub

Private Sub Form_Activate()
Flip.Left = (Screen.Width / 2) - (Flip.Width / 2)
Flip.Top = (Screen.Height / 2) - (Flip.Height / 2)
End Sub

Private Sub Form_Load()
frmOpen.Hide
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Visible = False
End With
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Command2.Visible = False
Command3.Visible = False
Wordlist
GetVars
'Text3.Visible = False
Text3.Text = LN

G(0).Visible = False
Pic3x8.Picture = ImageList2.ListImages(3).Picture
pic4x10.Picture = ImageList2.ListImages(6).Picture
pic6x12.Picture = ImageList2.ListImages(7).Picture
PicRS.Picture = ImageList2.ListImages(9).Picture
picMov.Picture = ImageList2.ListImages(12).Picture
PicFlag.Picture = ImageList2.ListImages(13).Picture
PicNums.Picture = ImageList2.ListImages(20).Picture
PicCards.Picture = ImageList2.ListImages(15).Picture
Piccols.Picture = ImageList2.ListImages(17).Picture
Pic1player.Picture = ImageList2.ListImages(22).Picture
pic2player.Picture = ImageList2.ListImages(23).Picture
mnu4X10.Checked = True
MnuAnything.Checked = True
mnuComputerWord.Checked = True
mnu1Player.Checked = True
mnuUserWord.Enabled = False
mnuShortTunes.Checked = True
txtWord.Enabled = False
Timer3.Enabled = False
lblplayer1.Visible = False
lblplayer2.Visible = False
txtWord.BackColor = RGB(255, 255, 255)
TxtBackground.BackColor = RGB(255, 255, 255)
With Toolbar
.Left = 125
.Top = 85
.Width = 9250
.Height = 525
End With
With lblNana
.Left = 1
.Height = 650
.Top = 5540
.Width = 9500
End With
With Shpplay1
.Visible = False
.Height = 525
.Left = 240
.Width = 4095
.Top = 5670
End With
With shpplay2
.Visible = False
.Height = 525
.Left = 5400
.Width = 4095
.Top = 5670
End With
ChangeLevels
For J = 1 To Z
G(J).Picture = ImageList2.ListImages(1).Picture
Next
frmOpen.Hide
End Sub
Private Sub Form_Paint()
Static Fd As Integer
If Fd <> 1 Then
Fd = 1
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmOpen.MMControl1.Command = "close"
End Sub


Private Sub G_Click(Index As Integer)
If CT = 0 Then
lblNana.Visible = False
lblTries.Visible = True
End If
Timer3.Enabled = True
CW = Index
If CW = 23 Or CW = 24 Then W = 1
L = L + 1: GO = GO + 1
If OPT = 1 Then
SX = SX - 2
If SX <= 2 Then SX = 2
End If
Timer2.Interval = SX
If PX = 0 Then
lblTries.Caption = "Number of tries: " + Str(GO)
End If
If PX = 1 Then
If SW = 1 Then
PL2 = PL2 + 1
shpplay2.Visible = True
End If
If SW = 0 Then
PL1 = PL1 + 1
Shpplay1.Visible = True
End If
lblplayer1.Visible = True
lblplayer2.Visible = True
lblplayer1.Caption = Player1 + "  " + Str(PL1) + "      " + Str(GA1)
lblplayer2.Caption = Player2 + "  " + Str(PL2) + "      " + Str(GA2)
End If
For CW = 1 To Z
If Index = CW Then
GoTo 50
End If
Next
50 Q(CW) = 1
G(CW).Enabled = False
C = CW
PlayTune
If OPT > 1 Then
Timer2.Interval = 500
GoTo 55
End If
If L = 2 Then
Timer3.Interval = DL
End If
If L = 1 Then
Timer3.Interval = DM
End If
55 MouseDisable
If CT = 1 Then
For J = 1 To Z
Next
End If
End Sub
Private Sub G_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = SWP(1) Then GoTo 20
SWP(2) = SWP(1)
If Q(SWP(2)) > 0 Then GoTo 19
G(SWP(2)).Picture = ImageList2.ListImages(1).Picture
10 G(Index).Picture = ImageList2.ListImages(2).Picture
GoTo 19
19 SWP(1) = Index
20 End Sub





Private Sub mnu1Player_Click()
Setup1Player
End Sub

Private Sub mnu2Player_Click()
Setup2Players
End Sub

Private Sub mnu3X8_Click()
Level1
End Sub

Private Sub mnu4X10_Click()
Level2
End Sub

Private Sub mnu5X12_Click()
Level3
End Sub

Private Sub MnuAbout_Click()
With frmOpen.MMControl1
        .Shareable = False
        .DeviceType = "Sequencer"
        .Command = "close"
        .filename = "76.MID"
        .Command = "Open"
        .Command = "play"
    End With
FrmAbout.Show 1
End Sub

Private Sub MnuAnything_Click()
Option1
End Sub

Private Sub MnuColors_Click()
Option6
End Sub

Private Sub mnuColors2_Click()
Option7
End Sub

Private Sub mnuComputerWord_Click()
UW = 0
ChangeLevels
mnuUserWord.Checked = False
mnuComputerWord.Checked = True
End Sub

Private Sub Mnucontinuous_Click()
If CT = 1 Then GoTo 10
CT = 1
If V = 0 Then
MsgBox "This Program is a Shareware demo, please see Ordering Info in the Help Menu"
CT = 0
GoTo 10
End If
Pic1player.Picture = ImageList2.ListImages(22).Picture
pic2player.Picture = ImageList2.ListImages(23).Picture
mnu2Player.Checked = False
mnu1Player.Checked = True
mnuComputerWord.Checked = False
mnuDisableWilds.Enabled = False
Mnucontinuous.Checked = True
mnuShortTunes.Checked = False
mnuNoTunes.Checked = False
lblplayer1.Visible = False
lblplayer2.Visible = False
Shpplay1.Visible = False
shpplay2.Visible = False
ChangeLevels
10
End Sub

Private Sub mnuDisableWilds_Click()
If OPT >= 2 Then GoTo 10
AX = AX + 1
If AX >= 2 Then AX = 0
If AX = 0 Then
W = 1
mnuDisableWilds.Checked = False
Else
W = 0
mnuDisableWilds.Checked = True
End If
10
End Sub

Private Sub mnuExit_Click()
If V = 1 Then GoTo 11
If V = 0 Then
With frmOpen.MMControl1
        .Shareable = False
        .DeviceType = "Sequencer"
        .Command = "close"
        .filename = "76.MID"
        .Command = "Open"
        .Command = "play"
    End With
frmAdvert.Show 1
If JJ = 1 Then GoTo 13
End If
11 CloseDown
12 End
13
End Sub









Public Function MouseDisable()
For J = 1 To Z
G(J).Enabled = False
Next
MousePointer = 1
Timer3.Interval = DL
End Function
Public Function MouseEnable()
If L = 2 Then
Match
End If
If MHD = 2 Then
For J = 1 To Z
If Q(J) = 0 Then
G(J).Enabled = True
End If
Next
MousePointer = 2
End If
Timer3.Enabled = False
End Function
Public Function Setup6by12()
If Z = 72 Then GoTo 100
Z = 72
lblA.Visible = False
lblB.Visible = False
lblC.Visible = False
lblD.Visible = False
lblE.Visible = False
lblF.Visible = False
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lbl4.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbl7.Visible = False
lbl8.Visible = False
lbl9.Visible = False
lbl10.Visible = False
lbl11.Visible = False
lbl12.Visible = False
InnerBord.Visible = False
Surround.Visible = False
txtWord.Visible = False
TxtBackground.Visible = False
txtWord = ""
For J = 1 To 72
G(J).Visible = False
Next
For J = 40 To 72
G(J).Picture = ImageList2.ListImages(1).Picture
Next
With lbl1
.Left = 770
.Top = 720
End With
With lbl2
.Left = 1490
.Top = 720
End With
With lbl3
.Left = 2210
.Top = 720
End With
With lbl4
.Left = 2930
.Top = 720
End With
With lbl5
.Left = 3650
.Top = 720
End With
With lbl6
.Left = 4370
.Top = 720
End With
With lbl7
.Left = 5090
.Top = 720
End With
With lbl8
.Left = 5790
.Top = 720
End With
With lbl9
.Left = 6540
.Top = 720
End With
With lbl10
.Left = 7170
.Top = 720
End With
With lbl11
.Left = 7870
.Top = 720
End With
With lbl12
.Left = 8570
.Top = 720
End With
With Surround
.BorderWidth = 10
.FillStyle = 0
.Height = 4815
.Left = 240
.Top = 720
.Width = 9150
End With
With lblA
.Left = 240
.Top = 1240
End With
With lblB
.Left = 240
.Top = 2000
End With
With lblC
.Left = 240
.Top = 2640
End With
With lblD
.Left = 240
.Top = 3360
End With
With lblE
.Left = 240
.Top = 4080
End With
With lblF
.Left = 240
.Top = 4800
End With
With TxtBackground
.Left = 480
.Height = 4335
.Top = 1080
.Width = 8655
End With
With txtWord
.Height = 1620
.Left = 480
.Top = 2520
.Width = 8660
End With
With InnerBord
.Left = 480
.Top = 1080
.Width = 8655
.Height = 4335
End With
100 Scramble
P = 0
For A = 1 To 72
P = P + 1
If T(P) = 1 Then
G(A).Left = 480
G(A).Top = 1080
End If
If T(P) = 2 Then
G(A).Left = 1200
G(A).Top = 1080
End If
If T(P) = 3 Then
G(A).Left = 1920
G(A).Top = 1080
End If
If T(P) = 4 Then
G(A).Left = 2640
G(A).Top = 1080
End If
If T(P) = 5 Then
G(A).Left = 3360
G(A).Top = 1080
End If
If T(P) = 6 Then
G(A).Left = 4080
G(A).Top = 1080
End If
If T(P) = 7 Then
G(A).Left = 4800
G(A).Top = 1080
End If
If T(P) = 8 Then
G(A).Left = 5520
G(A).Top = 1080
End If
If T(P) = 9 Then
G(A).Left = 6240
G(A).Top = 1080
End If
If T(P) = 10 Then
G(A).Left = 6960
G(A).Top = 1080
End If
If T(P) = 11 Then
G(A).Left = 7680
G(A).Top = 1080
End If
If T(P) = 12 Then
G(A).Left = 8400
G(A).Top = 1080
End If
If T(P) = 13 Then
G(A).Left = 480
G(A).Top = 1800
End If
If T(P) = 14 Then
G(A).Left = 1200
G(A).Top = 1800
End If
If T(P) = 15 Then
G(A).Left = 1920
G(A).Top = 1800
End If
If T(P) = 16 Then
G(A).Left = 2640
G(A).Top = 1800
End If
If T(P) = 17 Then
G(A).Left = 3360
G(A).Top = 1800
End If
If T(P) = 18 Then
G(A).Left = 4080
G(A).Top = 1800
End If
If T(P) = 19 Then
G(A).Left = 4800
G(A).Top = 1800
End If
If T(P) = 20 Then
G(A).Left = 5520
G(A).Top = 1800
End If
If T(P) = 21 Then
G(A).Left = 6240
G(A).Top = 1800
End If
If T(P) = 22 Then
G(A).Left = 6960
G(A).Top = 1800
End If
If T(P) = 23 Then
G(A).Left = 7680
G(A).Top = 1800
End If
If T(P) = 24 Then
G(A).Left = 8400
G(A).Top = 1800
End If
If T(P) = 25 Then
G(A).Left = 480
G(A).Top = 2520
End If
If T(P) = 26 Then
G(A).Left = 1200
G(A).Top = 2520
End If
If T(P) = 27 Then
G(A).Left = 1920
G(A).Top = 2520
End If
If T(P) = 28 Then
G(A).Left = 2640
G(A).Top = 2520
End If
If T(P) = 29 Then
G(A).Left = 3360
G(A).Top = 2520
End If
If T(P) = 30 Then
G(A).Left = 4080
G(A).Top = 2520
End If
If T(P) = 31 Then
G(A).Left = 4800
G(A).Top = 2520
End If
If T(P) = 32 Then
G(A).Left = 5520
G(A).Top = 2520
End If
If T(P) = 33 Then
G(A).Left = 6240
G(A).Top = 2520
End If
If T(P) = 34 Then
G(A).Left = 6960
G(A).Top = 2520
End If
If T(P) = 35 Then
G(A).Left = 7680
G(A).Top = 2520
End If
If T(P) = 36 Then
G(A).Left = 8400
G(A).Top = 2520
End If
If T(P) = 37 Then
G(A).Left = 480
G(A).Top = 3240
End If
If T(P) = 38 Then
G(A).Left = 1200
G(A).Top = 3240
End If
If T(P) = 39 Then
G(A).Left = 1920
G(A).Top = 3240
End If
If T(P) = 40 Then
G(A).Left = 2640
G(A).Top = 3240
End If
If T(P) = 41 Then
G(A).Left = 3360
G(A).Top = 3240
End If
If T(P) = 42 Then
G(A).Left = 4080
G(A).Top = 3240
End If
If T(P) = 43 Then
G(A).Left = 4800
G(A).Top = 3240
End If
If T(P) = 44 Then
G(A).Left = 5520
G(A).Top = 3240
End If
If T(P) = 45 Then
G(A).Left = 6240
G(A).Top = 3240
End If
If T(P) = 46 Then
G(A).Left = 6960
G(A).Top = 3240
End If
If T(P) = 47 Then
G(A).Left = 7680
G(A).Top = 3240
End If
If T(P) = 48 Then
G(A).Left = 8400
G(A).Top = 3240
End If
If T(P) = 49 Then
G(A).Left = 480
G(A).Top = 3960
End If
If T(P) = 50 Then
G(A).Left = 1200
G(A).Top = 3960
End If
If T(P) = 51 Then
G(A).Left = 1920
G(A).Top = 3960
End If
If T(P) = 52 Then
G(A).Left = 2640
G(A).Top = 3960
End If
If T(P) = 53 Then
G(A).Left = 3360
G(A).Top = 3960
End If
If T(P) = 54 Then
G(A).Left = 4080
G(A).Top = 3960
End If
If T(P) = 55 Then
G(A).Left = 4800
G(A).Top = 3960
End If
If T(P) = 56 Then
G(A).Left = 5520
G(A).Top = 3960
End If
If T(P) = 57 Then
G(A).Left = 6240
G(A).Top = 3960
End If
If T(P) = 58 Then
G(A).Left = 6960
G(A).Top = 3960
End If
If T(P) = 59 Then
G(A).Left = 7680
G(A).Top = 3960
End If
If T(P) = 60 Then
G(A).Left = 8400
G(A).Top = 3960
End If
If T(P) = 61 Then
G(A).Left = 480
G(A).Top = 4680
End If
If T(P) = 62 Then
G(A).Left = 1200
G(A).Top = 4680
End If
If T(P) = 63 Then
G(A).Left = 1920
G(A).Top = 4680
End If
If T(P) = 64 Then
G(A).Left = 2640
G(A).Top = 4680
End If
If T(P) = 65 Then
G(A).Left = 3360
G(A).Top = 4680
End If
If T(P) = 66 Then
G(A).Left = 4080
G(A).Top = 4680
End If
If T(P) = 67 Then
G(A).Left = 4800
G(A).Top = 4680
End If
If T(P) = 68 Then
G(A).Left = 5520
G(A).Top = 4680
End If
If T(P) = 69 Then
G(A).Left = 6240
G(A).Top = 4680
End If
If T(P) = 70 Then
G(A).Left = 6960
G(A).Top = 4680
End If
If T(P) = 71 Then
G(A).Left = 7680
G(A).Top = 4680
End If
If T(P) = 72 Then
G(A).Left = 8400
G(A).Top = 4680
End If
Next
lblA.Visible = True
lblB.Visible = True
lblC.Visible = True
lblD.Visible = True
lblE.Visible = True
lblF.Visible = True
lbl1.Visible = True
lbl2.Visible = True
lbl3.Visible = True
lbl4.Visible = True
lbl5.Visible = True
lbl6.Visible = True
lbl7.Visible = True
lbl8.Visible = True
lbl9.Visible = True
lbl10.Visible = True
lbl11.Visible = True
lbl12.Visible = True
InnerBord.Visible = True
Surround.Visible = True
txtWord.Visible = True
TxtBackground.Visible = True
For J = 1 To 72
G(J).Visible = True
Next
txtWord = Word
End Function


Public Sub Setup4by10()
If Z = 40 Then GoTo 100
Z = 40
txtWord = ""
For J = 1 To 72
G(J).Visible = False
Next
lblA.Visible = False
lblB.Visible = False
lblC.Visible = False
lblD.Visible = False
lblE.Visible = False
lblF.Visible = False
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lbl4.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbl7.Visible = False
lbl8.Visible = False
lbl9.Visible = False
lbl10.Visible = False
lbl11.Visible = False
lbl12.Visible = False
InnerBord.Visible = False
Surround.Visible = False
txtWord.Visible = False
TxtBackground.Visible = False
J = 0
With lblA
.Top = 1800
.Left = 1080
End With
With lblB
.Top = 2520
.Left = 1080
End With
With lblC
.Top = 3240
.Left = 1080
End With
With lblD
.Top = 3960
.Left = 1080
End With
With lbl1
.Left = 1590
.Top = 1200
End With
With lbl2
.Left = 2310
.Top = 1200
End With
With lbl3
.Left = 3030
.Top = 1200
End With
With lbl4
.Left = 3750
.Top = 1200
End With
With lbl5
.Left = 4470
.Top = 1200
End With
With lbl6
.Left = 5190
.Top = 1200
End With
With lbl7
.Left = 5910
.Top = 1200
End With
With lbl8
.Left = 6630
.Top = 1200
End With
With lbl9
.Left = 7350
.Top = 1200
End With
With lbl10
.Left = 8000
.Top = 1200
End With
With Surround
.Height = 3495
.Left = 1080
.Top = 1200
.Width = 7710
End With
With TxtBackground
.Height = 2870
.Left = 1320
.Top = 1560
.Width = 7215
End With
With txtWord
.Height = 1620
.Left = 1320
.Top = 2160
.Width = 7195
End With
With InnerBord
.Height = 2930
.Left = 1320
.Top = 1560
.Width = 7230
End With
100 Scramble
P = 0
For A = 1 To 40 Step 1
P = P + 1
If T(P) = 1 Then
G(A).Left = 1320
G(A).Top = 1560
End If
If T(P) = 2 Then
G(A).Left = 2040
G(A).Top = 1560
End If
If T(P) = 3 Then
G(A).Left = 2760
G(A).Top = 1560
End If
If T(P) = 4 Then
G(A).Left = 3480
G(A).Top = 1560
End If
If T(P) = 5 Then
G(A).Left = 4200
G(A).Top = 1560
End If
If T(P) = 6 Then
G(A).Left = 4920
G(A).Top = 1560
End If
If T(P) = 7 Then
G(A).Left = 5640
G(A).Top = 1560
End If
If T(P) = 8 Then
G(A).Left = 6360
G(A).Top = 1560
End If
If T(P) = 9 Then
G(A).Left = 7080
G(A).Top = 1560
End If
If T(P) = 10 Then
G(A).Left = 7800
G(A).Top = 1560
End If
If T(P) = 11 Then
G(A).Left = 1320
G(A).Top = 2280
End If
If T(P) = 12 Then
G(A).Left = 2040
G(A).Top = 2280
End If
If T(P) = 13 Then
G(A).Left = 2760
G(A).Top = 2280
End If
If T(P) = 14 Then
G(A).Left = 3480
G(A).Top = 2280
End If
If T(P) = 15 Then
G(A).Left = 4200
G(A).Top = 2280
End If
If T(P) = 16 Then
G(A).Left = 4920
G(A).Top = 2280
End If
If T(P) = 17 Then
G(A).Left = 5640
G(A).Top = 2280
End If
If T(P) = 18 Then
G(A).Left = 6360
G(A).Top = 2280
End If
If T(P) = 19 Then
G(A).Left = 7080
G(A).Top = 2280
End If
If T(P) = 20 Then
G(A).Left = 7800
G(A).Top = 2280
End If
If T(P) = 21 Then
G(A).Left = 1320
G(A).Top = 3000
End If
If T(P) = 22 Then
G(A).Left = 2040
G(A).Top = 3000
End If
If T(P) = 23 Then
G(A).Left = 2760
G(A).Top = 3000
End If
If T(P) = 24 Then
G(A).Left = 3480
G(A).Top = 3000
End If
If T(P) = 25 Then
G(A).Left = 4200
G(A).Top = 3000
End If
If T(P) = 26 Then
G(A).Left = 4920
G(A).Top = 3000
End If
If T(P) = 27 Then
G(A).Left = 5640
G(A).Top = 3000
End If
If T(P) = 28 Then
G(A).Left = 6360
G(A).Top = 3000
End If
If T(P) = 29 Then
G(A).Left = 7080
G(A).Top = 3000
End If
If T(P) = 30 Then
G(A).Left = 7800
G(A).Top = 3000
End If
If T(P) = 31 Then
G(A).Left = 1320
G(A).Top = 3720
End If
If T(P) = 32 Then
G(A).Left = 2040
G(A).Top = 3720
End If
If T(P) = 33 Then
G(A).Left = 2760
G(A).Top = 3720
End If
If T(P) = 34 Then
G(A).Left = 3480
G(A).Top = 3720
End If
If T(P) = 35 Then
G(A).Left = 4200
G(A).Top = 3720
End If
If T(P) = 36 Then
G(A).Left = 4920
G(A).Top = 3720
End If
If T(P) = 37 Then
G(A).Left = 5640
G(A).Top = 3720
End If
If T(P) = 38 Then
G(A).Left = 6360
G(A).Top = 3720
End If
If T(P) = 39 Then
G(A).Left = 7080
G(A).Top = 3720
End If
If T(P) = 40 Then
G(A).Left = 7800
G(A).Top = 3720
End If
Next
lblA.Visible = True
lblB.Visible = True
lblC.Visible = True
lblD.Visible = True
lbl1.Visible = True
lbl2.Visible = True
lbl3.Visible = True
lbl4.Visible = True
lbl5.Visible = True
lbl6.Visible = True
lbl7.Visible = True
lbl8.Visible = True
lbl9.Visible = True
lbl10.Visible = True
InnerBord.Visible = True
Surround.Visible = True
txtWord.Visible = True
TxtBackground.Visible = True


For J = 1 To 40
G(J).Visible = True
Next
txtWord = Word
End Sub

Public Function Setup3by8()
If Z = 24 Then GoTo 100
Z = 24
txtWord = ""
For J = 1 To 72
G(J).Visible = False
Next
lblD.Visible = False
lblE.Visible = False
lblF.Visible = False
lbl9.Visible = False
lbl10.Visible = False
lbl11.Visible = False
lbl12.Visible = False
InnerBord.Visible = False
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lbl4.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbl7.Visible = False
lbl8.Visible = False
lblA.Visible = False
lblB.Visible = False
lblC.Visible = False
lblD.Visible = False
Surround.Visible = False
txtWord.Visible = False
TxtBackground.Visible = False
With InnerBord
.Height = 2190
.Left = 1680
.Width = 5790
.Top = 1920
End With
With lbl1
.Left = 1950
.Top = 1560
End With
With lbl2
.Left = 2670
.Top = 1560
End With
With lbl3
.Left = 3390
.Top = 1560
End With
With lbl4
.Left = 4110
.Top = 1560
End With
With lbl5
.Left = 4830
.Top = 1560
End With
With lbl6
.Left = 5550
.Top = 1560
End With
With lbl7
.Left = 6270
.Top = 1560
End With
With lbl8
.Left = 6990
.Top = 1560
End With
With lblA
.Top = 2160
.Left = 1440
End With
With lblB
.Top = 2880
.Left = 1440
End With
With lblC
.Left = 1440
.Top = 3600
End With
With Surround
.Top = 1600
.Width = 6270
.Left = 1440
.Height = 2770
End With
With txtWord
.Top = 2280
.Height = 1620
.Left = 1680
.Width = 5775
End With
With TxtBackground
.Left = 1680
.Height = 2175
.Top = 1920
.Width = 5775
End With
100 Scramble
P = 0
For A = 1 To 24
P = P + 1
If T(P) = 1 Then
G(A).Left = 1680
G(A).Top = 1920
End If
If T(P) = 2 Then
G(A).Left = 2400
G(A).Top = 1920
End If
If T(P) = 3 Then
G(A).Left = 3120
G(A).Top = 1920
End If
If T(P) = 4 Then
G(A).Left = 3840
G(A).Top = 1920
End If
If T(P) = 5 Then
G(A).Left = 4560
G(A).Top = 1920
End If
If T(P) = 6 Then
G(A).Left = 5280
G(A).Top = 1920
End If
If T(P) = 7 Then
G(A).Left = 6000
G(A).Top = 1920
End If
If T(P) = 8 Then
G(A).Left = 6720
G(A).Top = 1920
End If
If T(P) = 9 Then
G(A).Left = 1680
G(A).Top = 2640
End If
If T(P) = 10 Then
G(A).Left = 2400
G(A).Top = 2640
End If
If T(P) = 11 Then
G(A).Left = 3120
G(A).Top = 2640
End If
If T(P) = 12 Then
G(A).Left = 3840
G(A).Top = 2640
End If
If T(P) = 13 Then
G(A).Left = 4560
G(A).Top = 2640
End If
If T(P) = 14 Then
G(A).Left = 5280
G(A).Top = 2640
End If
If T(P) = 15 Then
G(A).Left = 6000
G(A).Top = 2640
End If
If T(P) = 16 Then
G(A).Left = 6720
G(A).Top = 2640
End If
If T(P) = 17 Then
G(A).Left = 1680
G(A).Top = 3360
End If
If T(P) = 18 Then
G(A).Left = 2400
G(A).Top = 3360
End If
If T(P) = 19 Then
G(A).Left = 3120
G(A).Top = 3360
End If
If T(P) = 20 Then
G(A).Left = 3840
G(A).Top = 3360
End If
If T(P) = 21 Then
G(A).Left = 4560
G(A).Top = 3360
End If
If T(P) = 22 Then
G(A).Left = 5280
G(A).Top = 3360
End If
If T(P) = 23 Then
G(A).Left = 6000
G(A).Top = 3360
End If
If T(P) = 24 Then
G(A).Left = 6720
G(A).Top = 3360
End If
Next
InnerBord.Visible = True
lbl1.Visible = True
lbl2.Visible = True
lbl3.Visible = True
lbl4.Visible = True
lbl5.Visible = True
lbl6.Visible = True
lbl7.Visible = True
lbl8.Visible = True
lblA.Visible = True
lblB.Visible = True
lblC.Visible = True
Surround.Visible = True
txtWord.Visible = True
TxtBackground.Visible = True
For J = 1 To 24
G(J).Visible = True
Next
txtWord = Word
End Function

Public Function ChangeLevels()
If AB >= 1 Then
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
.filename = "75.MID"
.Command = "Open"
.Command = "play"
End With
End If
2 AB = AB + 1
lblNana.Visible = False
txtWord.Text = " "
NewGame
If R = 1 Then
Setup3by8
End If
If R = 2 Then
Setup4by10
End If
If R = 3 Then
Setup6by12
End If
L = 0
If CT = 1 Then
TimedGame
GoTo 11
End If
Dictionary
If PX = 0 Then
lblTries.Visible = True
End If
If PX = 1 Then
lblplayer1.Visible = True
lblplayer2.Visible = True
Shpplay1.Visible = True
shpplay2.Visible = False
txtWord.Text = Word2
'Text3.Text = Word2
End If
If V = 0 Then
If OPT >= 2 Then
For J = 1 To Z
Q(J) = 2
G(J).Enabled = False
Next
MsgBox "This Program is a Shareware demo, please see Ordering Info in the Help Menu"
GoTo 11
End If
If R = 3 Then
For J = 1 To Z
Q(J) = 2
G(J).Enabled = False
Next
MsgBox "This Program is a Shareware demo, please see Ordering Info in the Help Menu"
GoTo 11
End If
End If
11
End Function
Public Function Scramble()
Randomize Timer
For IA = 1 To Z: If T(IA) = 0 Then GoTo 30
10 GoTo 30
20 IC = 0
For J = 1 To (IA - 1): If T(J) = IB Then IC = 1
Next
If IC = 0 Then GoTo 40
30 IB = (Int(Rnd * Z) + 1)
If IA = 1 Then GoTo 40
GoTo 20
40 T(IA) = IB
Next
End Function

Public Function GuessTheWord()
For J = 1 To Z
If Q(J) >= 2 Then G(J).Visible = False
Next
FrmGuessTheWord.Show 1
GWD = UCase(GWD)
If GWD = "" Then GoTo 70
If PX = 0 Then
If GWD = Word Then MHD = 1
If GWD <> Word Then MHD = 0
End If
If PX = 1 Then
If SW = 1 Then
If GWD = Word2 Then
MHD = 1
GoTo 50
End If
If GWD <> Word2 Then
MHD = 0
GoTo 60
End If
End If
If SW = 0 Then
If GWD = Word1 Then
MHD = 1
GoTo 50
End If
If GWD <> Word1 Then
MHD = 0
GoTo 60
End If
End If
End If
If MHD = 1 Then GoTo 50
If MHD = 0 Then GoTo 60
If MHD = 2 Then
MHD = 0
GoTo 60
End If
50 PlayCorrect
Timer3.Enabled = False
Timer1.Enabled = True
GoTo 70
60 Timer4.Enabled = True
For J = 1 To Z
If Q(J) = 0 Then
G(J).Enabled = False
End If
Next
PlayWrong
70
For J = 1 To Z
If Q(J) >= 2 Then
G(J).Visible = True
G(J).Enabled = False
End If
Next
End Function
Public Function PlayCorrect()
If SA = 1 Then GoTo 9001
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
.filename = "74.MID"
.Command = "Open"
.Command = "play"
End With
9001
End Function
Public Function PlayWrong()
If SA = 1 Then GoTo 9001
If CT = 1 Then GoTo 9001
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
.filename = "73.MID"
.Command = "Open"
.Command = "play"
End With
9001
End Function

Public Function Match()
L = 0
If AX = 1 Then GoTo 5
If W = 1 Then
WProc
GoTo 15
End If
5 For DX = 1 To Z Step 2
If Q(DX) = 1 And Q((DX + 1)) = 1 Then
Q(DX) = 2: Q((DX + 1)) = 2
CX = CX + 2
If PX = 1 Then
If SW = 0 Then
SW = 1
GoTo 6
End If
If SW = 1 Then
SW = 0
End If
End If
6 If K(DX) = 1 Then GoTo 15
U(DX) = 3
PlayTuneB
U(DX) = 0
If CT = 1 Then GoTo 7
GuessTheWord
7 K(DX) = 1
GoTo 15
End If
Next
For J = 1 To Z
If Q(J) = 1 Then
G(J).Picture = ImageList2.ListImages(1).Picture
Q(J) = 0
SX = SX + 2
End If
Next
For J = 1 To Z
N(J) = 0
Next
GoTo 15
15 If PX = 1 Then
Switch
End If
End Function
Public Function PlayTune()
If SA = 1 Then GoTo 9001
If CT = 1 Then GoTo 9001
With frmOpen.MMControl1
.Wait = False
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
End With
If FW = 1 Then GoTo 240
10 If C > 2 Then GoTo 30
If N(1) = 1 Then GoTo 20
If ST = 0 Then GoTo 11
frmOpen.MMControl1.filename = "1C.MID"
GoTo 12
11 frmOpen.MMControl1.filename = "1A.MID"
12 N(1) = 1
GoTo 9000
20 If ST = 0 Then GoTo 21
frmOpen.MMControl1.filename = "1D.MID"
GoTo 9000
21 frmOpen.MMControl1.filename = "1B.MID"
22 GoTo 9000
30 If C > 4 Then GoTo 50
If N(3) = 1 Then GoTo 40
If ST = 0 Then GoTo 31
frmOpen.MMControl1.filename = "2C.MID"
GoTo 32
31 frmOpen.MMControl1.filename = "2A.MID"
32 N(3) = 1
GoTo 9000
40 If ST = 0 Then GoTo 41
frmOpen.MMControl1.filename = "2D.MID"
GoTo 9000
41 frmOpen.MMControl1.filename = "2B.MID"
42 GoTo 9000
50 If C > 6 Then GoTo 70
If N(5) = 1 Then GoTo 60
If ST = 0 Then GoTo 51
frmOpen.MMControl1.filename = "3C.MID"
GoTo 52
51 frmOpen.MMControl1.filename = "3A.MID"
52 N(5) = 1
GoTo 9000
60 If ST = 0 Then GoTo 61
frmOpen.MMControl1.filename = "3D.MID"
GoTo 9000
61 frmOpen.MMControl1.filename = "3B.MID"
62 GoTo 9000
70 If C > 8 Then GoTo 90
If N(7) = 1 Then GoTo 80
If ST = 0 Then GoTo 71
frmOpen.MMControl1.filename = "4C.MID"
GoTo 72
71 frmOpen.MMControl1.filename = "4A.MID"
72 N(7) = 1
GoTo 9000
80 If ST = 0 Then GoTo 81
frmOpen.MMControl1.filename = "4D.MID"
GoTo 9000
81 frmOpen.MMControl1.filename = "4B.MID"
82 GoTo 9000
90 If C > 10 Then GoTo 110
If N(9) = 1 Then GoTo 100
If ST = 0 Then GoTo 91
frmOpen.MMControl1.filename = "5C.MID"
GoTo 92
91 frmOpen.MMControl1.filename = "5A.MID"
92 N(9) = 1
GoTo 9000
100 If ST = 0 Then GoTo 101
frmOpen.MMControl1.filename = "5D.MID"
GoTo 9000
101 frmOpen.MMControl1.filename = "5B.MID"
102 GoTo 9000
110 If C > 12 Then GoTo 130
If N(11) = 1 Then GoTo 120
If ST = 0 Then GoTo 111
frmOpen.MMControl1.filename = "6C.MID"
GoTo 112
111 frmOpen.MMControl1.filename = "6A.MID"
112 N(11) = 1
GoTo 9000
120 If ST = 0 Then GoTo 121
frmOpen.MMControl1.filename = "6D.MID"
GoTo 9000
121 frmOpen.MMControl1.filename = "6B.MID"
122 GoTo 9000
130 If C > 14 Then GoTo 150
If N(13) = 1 Then GoTo 140
If ST = 0 Then GoTo 131
frmOpen.MMControl1.filename = "7C.MID"
GoTo 132
131 frmOpen.MMControl1.filename = "7A.MID"
132 N(13) = 1
GoTo 9000
140 If ST = 0 Then GoTo 141
frmOpen.MMControl1.filename = "7D.MID"
GoTo 9000
141 frmOpen.MMControl1.filename = "7B.MID"
142 GoTo 9000
150 If C > 16 Then GoTo 170
If N(15) = 1 Then GoTo 160
If ST = 0 Then GoTo 151
151 frmOpen.MMControl1.filename = "8C.MID"
GoTo 152
frmOpen.MMControl1.filename = "8A.MID"
152 N(15) = 1
GoTo 9000
160 If ST = 0 Then GoTo 161
frmOpen.MMControl1.filename = "8D.MID"
GoTo 9000
161 frmOpen.MMControl1.filename = "8B.MID"
162 GoTo 9000
170 If C > 18 Then GoTo 190
If N(17) = 1 Then GoTo 180
If ST = 0 Then GoTo 171
frmOpen.MMControl1.filename = "9C.MID"
GoTo 172
171 frmOpen.MMControl1.filename = "9A.MID"
172 N(17) = 1
GoTo 9000
180 If ST = 0 Then GoTo 181
frmOpen.MMControl1.filename = "9D.MID"
GoTo 9000
181 frmOpen.MMControl1.filename = "9B.MID"
182 GoTo 9000
190 If C > 20 Then GoTo 210
If N(19) = 1 Then GoTo 200
If ST = 0 Then GoTo 191
frmOpen.MMControl1.filename = "10C.MID"
GoTo 192
191 frmOpen.MMControl1.filename = "10A.MID"
192 N(19) = 1
GoTo 9000
200 If ST = 0 Then GoTo 201
frmOpen.MMControl1.filename = "10D.MID"
GoTo 9000
201 frmOpen.MMControl1.filename = "10B.MID"
202 GoTo 9000
210 If C > 22 Then GoTo 230
If N(21) = 1 Then GoTo 220
If ST = 0 Then GoTo 211
frmOpen.MMControl1.filename = "11C.MID"
GoTo 212
211 frmOpen.MMControl1.filename = "11A.MID"
212 N(21) = 1
GoTo 9000
220 If ST = 0 Then GoTo 221
frmOpen.MMControl1.filename = "11D.MID"
GoTo 9000
221 frmOpen.MMControl1.filename = "11B.MID"
223 GoTo 9000
230 If AX = 1 Then GoTo 244
If C > 24 Then GoTo 250
If Q(23) = 1 Then Q(23) = 2
If Q(24) = 1 Then Q(24) = 2
FW = 1
If N(23) = 1 Then GoTo 240
If ST = 0 Then GoTo 231
frmOpen.MMControl1.filename = "12C.MID"
GoTo 233
231 frmOpen.MMControl1.filename = "12A.MID"
233 N(23) = 1
If L = 2 Then GoTo 240
GoTo 9000
240 If L = 1 Then
N(23) = 0
GoTo 230
End If
If ST = 0 Then GoTo 241
frmOpen.MMControl1.filename = "12D.MID"
GoTo 243
241 frmOpen.MMControl1.filename = "12B.MID"
243 FW = 0
GoTo 9000
244 If C > 24 Then GoTo 250
If N(23) = 1 Then GoTo 247
If ST = 0 Then GoTo 245
frmOpen.MMControl1.filename = "12C.MID"
GoTo 246
245 frmOpen.MMControl1.filename = "12a.MID"
246 N(23) = 1
GoTo 9000
247 If ST = 0 Then GoTo 248
frmOpen.MMControl1.filename = "12D.MID"
GoTo 9000
248 frmOpen.MMControl1.filename = "12B.MID"
249 GoTo 9000
250 If C > 26 Then GoTo 270
If N(25) = 1 Then GoTo 260
If ST = 0 Then GoTo 251
frmOpen.MMControl1.filename = "13C.MID"
GoTo 252
251 frmOpen.MMControl1.filename = "13a.MID"
252 N(25) = 1
GoTo 9000
260 If ST = 0 Then GoTo 261
frmOpen.MMControl1.filename = "13D.MID"
GoTo 9000
261 frmOpen.MMControl1.filename = "13B.MID"
262 GoTo 9000
270 If C > 28 Then GoTo 290
If N(27) = 1 Then GoTo 280
If ST = 0 Then GoTo 271
frmOpen.MMControl1.filename = "14C.MID"
GoTo 272
271 frmOpen.MMControl1.filename = "14a.MID"
272 N(27) = 1
GoTo 9000
280 If ST = 0 Then GoTo 281
frmOpen.MMControl1.filename = "14D.MID"
GoTo 9000
281 frmOpen.MMControl1.filename = "14B.MID"
282 GoTo 9000
290 If C > 30 Then GoTo 310
If N(29) = 1 Then GoTo 300
If ST = 0 Then GoTo 291
frmOpen.MMControl1.filename = "15C.MID"
GoTo 292
291 frmOpen.MMControl1.filename = "15a.MID"
292 N(29) = 1
GoTo 9000
300 If ST = 0 Then GoTo 301
frmOpen.MMControl1.filename = "15D.MID"
GoTo 9000
301 frmOpen.MMControl1.filename = "15B.MID"
302 GoTo 9000
310 If C > 32 Then GoTo 330
If N(31) = 1 Then GoTo 320
If ST = 0 Then GoTo 311
frmOpen.MMControl1.filename = "16C.MID"
GoTo 312
311 frmOpen.MMControl1.filename = "16A.MID"
312 N(31) = 1
GoTo 9000
320 If ST = 0 Then GoTo 321
frmOpen.MMControl1.filename = "16D.MID"
GoTo 9000
321 frmOpen.MMControl1.filename = "16b.MID"
322 GoTo 9000
330 If C > 34 Then GoTo 350
If N(33) = 1 Then GoTo 340
If ST = 0 Then GoTo 331
frmOpen.MMControl1.filename = "17C.MID"
GoTo 332
331 frmOpen.MMControl1.filename = "17a.MID"
332 N(33) = 1
GoTo 9000
340 If ST = 0 Then GoTo 341
frmOpen.MMControl1.filename = "17D.MID"
GoTo 9000
341 frmOpen.MMControl1.filename = "17B.MID"
342 GoTo 9000
350 If C > 36 Then GoTo 370
If N(35) = 1 Then GoTo 360
If ST = 0 Then GoTo 351
frmOpen.MMControl1.filename = "18C.MID"
GoTo 352
351 frmOpen.MMControl1.filename = "18a.MID"
352 N(35) = 1
GoTo 9000
360 If ST = 0 Then GoTo 361
frmOpen.MMControl1.filename = "18D.MID"
GoTo 9000
361 frmOpen.MMControl1.filename = "18B.MID"
362 GoTo 9000
370 If C > 38 Then GoTo 390
If N(37) = 1 Then GoTo 380
If ST = 0 Then GoTo 371
frmOpen.MMControl1.filename = "19C.MID"
GoTo 372
371 frmOpen.MMControl1.filename = "19a.MID"
372 N(37) = 1
GoTo 9000
380 If ST = 0 Then GoTo 381
frmOpen.MMControl1.filename = "19D.MID"
GoTo 9000
381 frmOpen.MMControl1.filename = "19B.MID"
382 GoTo 9000
390 If C > 40 Then GoTo 410
If N(39) = 1 Then GoTo 400
If ST = 0 Then GoTo 391
frmOpen.MMControl1.filename = "20C.MID"
GoTo 392
391 frmOpen.MMControl1.filename = "20a.MID"
392 N(39) = 1
GoTo 9000
400 If ST = 0 Then GoTo 401
frmOpen.MMControl1.filename = "20D.MID"
GoTo 9000
401 frmOpen.MMControl1.filename = "20B.MID"
402 GoTo 9000
410 If C > 42 Then GoTo 430
If N(41) = 1 Then GoTo 420
If ST = 0 Then GoTo 411
frmOpen.MMControl1.filename = "21C.MID"
GoTo 412
411 frmOpen.MMControl1.filename = "21a.MID"
412 N(41) = 1
GoTo 9000
420 If ST = 0 Then GoTo 421
frmOpen.MMControl1.filename = "21D.MID"
GoTo 9000
421 frmOpen.MMControl1.filename = "21B.MID"
422 GoTo 9000
430 If C > 44 Then GoTo 450
If N(43) = 1 Then GoTo 440
If ST = 0 Then GoTo 431
frmOpen.MMControl1.filename = "22C.MID"
GoTo 432
431 frmOpen.MMControl1.filename = "22a.MID"
432 N(43) = 1
GoTo 9000
440 If ST = 0 Then GoTo 441
frmOpen.MMControl1.filename = "22D.MID"
GoTo 9000
441 frmOpen.MMControl1.filename = "22B.MID"
442 GoTo 9000
450 If C > 46 Then GoTo 470
If N(45) = 1 Then GoTo 460
If ST = 0 Then GoTo 451
frmOpen.MMControl1.filename = "23C.MID"
GoTo 452
451 frmOpen.MMControl1.filename = "23a.MID"
452 N(45) = 1
GoTo 9000
460 If ST = 0 Then GoTo 461
frmOpen.MMControl1.filename = "23D.MID"
GoTo 9000
461 frmOpen.MMControl1.filename = "23B.MID"
462 GoTo 9000
470 If C > 48 Then GoTo 490
If N(47) = 1 Then GoTo 480
If ST = 0 Then GoTo 471
frmOpen.MMControl1.filename = "24C.MID"
GoTo 472
471 frmOpen.MMControl1.filename = "24a.MID"
472 N(47) = 1
GoTo 9000
480 If ST = 0 Then GoTo 481
frmOpen.MMControl1.filename = "24D.MID"
GoTo 9000
481 frmOpen.MMControl1.filename = "24B.MID"
482 GoTo 9000
490 If C > 50 Then GoTo 510
If N(49) = 1 Then GoTo 500
If ST = 0 Then GoTo 491
frmOpen.MMControl1.filename = "25C.MID"
GoTo 492
491 frmOpen.MMControl1.filename = "25a.MID"
492 N(49) = 1
GoTo 9000
500 If ST = 0 Then GoTo 501
frmOpen.MMControl1.filename = "25D.MID"
GoTo 9000
501 frmOpen.MMControl1.filename = "25B.MID"
502 GoTo 9000
510 If C > 52 Then GoTo 530
If N(51) = 1 Then GoTo 520
If ST = 0 Then GoTo 511
frmOpen.MMControl1.filename = "26C.MID"
GoTo 512
511 frmOpen.MMControl1.filename = "26a.MID"
512 N(51) = 1
GoTo 9000
520 If ST = 0 Then GoTo 521
frmOpen.MMControl1.filename = "26D.MID"
GoTo 9000
521 frmOpen.MMControl1.filename = "26B.MID"
522 GoTo 9000
530 If C > 54 Then GoTo 550
If N(53) = 1 Then GoTo 540
If ST = 0 Then GoTo 531
frmOpen.MMControl1.filename = "27C.MID"
GoTo 532
531 frmOpen.MMControl1.filename = "27a.MID"
532 N(53) = 1
GoTo 9000
540 If ST = 0 Then GoTo 541
frmOpen.MMControl1.filename = "27D.MID"
GoTo 9000
541 frmOpen.MMControl1.filename = "27B.MID"
542 GoTo 9000
550 If C > 56 Then GoTo 570
If N(55) = 1 Then GoTo 560
If ST = 0 Then GoTo 551
frmOpen.MMControl1.filename = "28C.MID"
GoTo 552
551 frmOpen.MMControl1.filename = "28a.MID"
552 N(55) = 1
GoTo 9000
560 If ST = 0 Then GoTo 561
frmOpen.MMControl1.filename = "28D.MID"
GoTo 9000
561 frmOpen.MMControl1.filename = "28B.MID"
562 GoTo 9000
570 If C > 58 Then GoTo 590
If N(57) = 1 Then GoTo 580
If ST = 0 Then GoTo 571
frmOpen.MMControl1.filename = "29C.MID"
GoTo 572
571 frmOpen.MMControl1.filename = "29a.MID"
572 N(57) = 1
GoTo 9000
580 If ST = 0 Then GoTo 581
frmOpen.MMControl1.filename = "29D.MID"
GoTo 9000
581 frmOpen.MMControl1.filename = "29B.MID"
582 GoTo 9000
590 If C > 60 Then GoTo 610
If N(59) = 1 Then GoTo 600
If ST = 0 Then GoTo 591
frmOpen.MMControl1.filename = "30C.MID"
GoTo 592
591 frmOpen.MMControl1.filename = "30a.MID"
592 N(59) = 1
GoTo 9000
600 If ST = 0 Then GoTo 601
frmOpen.MMControl1.filename = "30D.MID"
GoTo 9000
601 frmOpen.MMControl1.filename = "30B.MID"
602 GoTo 9000
610 If C > 62 Then GoTo 630
If N(61) = 1 Then GoTo 620
If ST = 0 Then GoTo 611
frmOpen.MMControl1.filename = "31C.MID"
GoTo 612
611 frmOpen.MMControl1.filename = "31a.MID"
612 N(61) = 1
GoTo 9000
620 If ST = 0 Then GoTo 621
frmOpen.MMControl1.filename = "31D.MID"
GoTo 9000
621 frmOpen.MMControl1.filename = "31B.MID"
622 GoTo 9000
630 If C > 64 Then GoTo 650
If N(63) = 1 Then GoTo 640
If ST = 0 Then GoTo 631
frmOpen.MMControl1.filename = "32C.MID"
GoTo 632
631 frmOpen.MMControl1.filename = "32a.MID"
632 N(63) = 1
GoTo 9000
640 If ST = 0 Then GoTo 641
frmOpen.MMControl1.filename = "32D.MID"
GoTo 9000
641 frmOpen.MMControl1.filename = "32B.MID"
642 GoTo 9000
650 If C > 66 Then GoTo 670
If N(65) = 1 Then GoTo 660
If ST = 0 Then GoTo 651
frmOpen.MMControl1.filename = "33C.MID"
GoTo 652
651 frmOpen.MMControl1.filename = "33a.MID"
652 N(65) = 1
GoTo 9000
660 If ST = 0 Then GoTo 661
frmOpen.MMControl1.filename = "33D.MID"
GoTo 9000
661 frmOpen.MMControl1.filename = "33B.MID"
662 GoTo 9000
670 If C > 68 Then GoTo 690
If N(67) = 1 Then GoTo 680
If ST = 0 Then GoTo 671
frmOpen.MMControl1.filename = "34C.MID"
GoTo 672
671 frmOpen.MMControl1.filename = "34a.MID"
672 N(67) = 1
GoTo 9000
680 If ST = 0 Then GoTo 681
frmOpen.MMControl1.filename = "34D.MID"
GoTo 9000
681 frmOpen.MMControl1.filename = "34B.MID"
682 GoTo 9000
690 If C > 70 Then GoTo 710
If N(69) = 1 Then GoTo 700
If ST = 0 Then GoTo 691
frmOpen.MMControl1.filename = "35C.MID"
GoTo 692
691 frmOpen.MMControl1.filename = "35a.MID"
692 N(69) = 1
GoTo 9000
700 If ST = 0 Then GoTo 701
frmOpen.MMControl1.filename = "35D.MID"
GoTo 9000
701 frmOpen.MMControl1.filename = "35B.MID"
702 GoTo 9000
710 If C > 72 Then GoTo 710
If N(71) = 1 Then GoTo 720
If ST = 0 Then GoTo 711
frmOpen.MMControl1.filename = "36C.MID"
GoTo 712
711 frmOpen.MMControl1.filename = "36a.MID"
712 N(71) = 1
GoTo 9000
720 If ST = 0 Then GoTo 721
frmOpen.MMControl1.filename = "36D.MID"
GoTo 9000
721 frmOpen.MMControl1.filename = "36B.MID"
9000 frmOpen.MMControl1.Command = "Open"
frmOpen.MMControl1.Command = "play"
9001
End Function
Public Function GetVars()
OPT = 1
MHD = 2
SX = 100
CT = 0
PX = 0
FC = Len(FA)
FB = FreeFile
Fd = 1
UW = 0
DL = 1825
DM = 2325
PA(4) = 12
R = 2
ST = 1
X = 1680: Y = 1920
H = 72
DY = 289
DZ = 326
DN = 362
M = 8
V = 0
D(1) = 1
D(2) = 1
E(1) = 8
E(2) = 8
S(1) = D(1)
S(2) = D(2)
FT(1) = DY
FT(2) = DY
FU(1) = DZ
FU(2) = DZ
FV(1) = DN
FV(2) = DN
For J = 3 To (H - 1) Step 2
D(J) = D(J - 2) + M
D(J + 1) = D(J - 2) + M
E(J) = E(J - 2) + M
E(J + 1) = E(J - 2) + M
S(J) = D(J)
S(J + 1) = D(J + 1)
DY = DY + 1
DZ = DZ + 1
DN = DN + 1
FT(J) = DY
FT(J + 1) = DY
FU(J) = DZ
FU(J + 1) = DZ
FV(J) = DN
FV(J + 1) = DN
Next
Open "C:\Dsort.DAT" For Random As FB Len = FC
Fd = 1
Get #FB, Fd, FA
If FA.BScore3 < 2 Then
LB = "JaySoft"
LC = "JaySoft"
LD = "JaySoft"
LE = 100
LF = 200
LG = 300
LH = "Anon"
LI = "Anon"
LJ = "Anon"
LK = 101
LL = 201
LM = 301
JL = 999
JM = 999
JN = 999
LO = "JaySoft"
LP = "JaySoft"
LQ = "JaySoft"
Else
LA = FA.Name
LB = FA.Begin1
LC = FA.Inter1
LD = FA.Expert1
LE = FA.BScore1
LF = FA.BScore2
LG = FA.BScore3
LH = FA.Begin2
LI = FA.Inter2
LJ = FA.Expert2
LK = FA.WScore1
LL = FA.WScore2
LM = FA.WScore2
LO = FA.TName1
LP = FA.Tname2
LQ = FA.Tname3
JL = FA.BTime1
JM = FA.BTime2
JN = FA.BTime3
LN = FA.Validate
End If
If LN = Trim("30169862") Then
V = 1
Upgrade
End If
End Function
Public Function PlayTuneB()
If SA = 1 Then GoTo 9001
If CT = 1 Then GoTo 9001
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
End With
If U(1) = 3 Then GoTo 10
If U(3) = 3 Then GoTo 30
If U(5) = 3 Then GoTo 50
If U(7) = 3 Then GoTo 70
If U(9) = 3 Then GoTo 90
If U(11) = 3 Then GoTo 110
If U(13) = 3 Then GoTo 130
If U(15) = 3 Then GoTo 150
If U(17) = 3 Then GoTo 170
If U(19) = 3 Then GoTo 190
If U(21) = 3 Then GoTo 210
If U(23) = 3 Then GoTo 230
If U(25) = 3 Then GoTo 250
If U(27) = 3 Then GoTo 270
If U(29) = 3 Then GoTo 290
If U(31) = 3 Then GoTo 310
If U(33) = 3 Then GoTo 330
If U(35) = 3 Then GoTo 350
If U(37) = 3 Then GoTo 370
If U(39) = 3 Then GoTo 390
If U(41) = 3 Then GoTo 410
If U(43) = 3 Then GoTo 430
If U(45) = 3 Then GoTo 450
If U(47) = 3 Then GoTo 470
If U(49) = 3 Then GoTo 490
If U(51) = 3 Then GoTo 510
If U(53) = 3 Then GoTo 530
If U(55) = 3 Then GoTo 550
If U(57) = 3 Then GoTo 570
If U(59) = 3 Then GoTo 590
If U(61) = 3 Then GoTo 610
If U(63) = 3 Then GoTo 630
If U(65) = 3 Then GoTo 650
If U(67) = 3 Then GoTo 670
If U(69) = 3 Then GoTo 690
If U(71) = 3 Then GoTo 710
GoTo 9000
10 frmOpen.MMControl1.filename = "1.MID"
GoTo 9000
30 frmOpen.MMControl1.filename = "2.MID"
GoTo 9000
50 frmOpen.MMControl1.filename = "3.MID"
GoTo 9000
70 frmOpen.MMControl1.filename = "4.MID"
GoTo 9000
90 frmOpen.MMControl1.filename = "5.MID"
GoTo 9000
110 frmOpen.MMControl1.filename = "6.MID"
GoTo 9000
130 frmOpen.MMControl1.filename = "7.MID"
GoTo 9000
150 frmOpen.MMControl1.filename = "8.MID"
GoTo 9000
170 frmOpen.MMControl1.filename = "9.MID"
GoTo 9000
190 frmOpen.MMControl1.filename = "10.MID"
GoTo 9000
210 frmOpen.MMControl1.filename = "11.MID"
GoTo 9000
230 frmOpen.MMControl1.filename = "12.MID"
GoTo 9000
250 frmOpen.MMControl1.filename = "13.MID"
GoTo 9000
270 frmOpen.MMControl1.filename = "14.MID"
GoTo 9000
290 frmOpen.MMControl1.filename = "15.MID"
GoTo 9000
310 frmOpen.MMControl1.filename = "16.MID"
GoTo 9000
330 frmOpen.MMControl1.filename = "17.MID"
GoTo 9000
350 frmOpen.MMControl1.filename = "18.MID"
GoTo 9000
370 frmOpen.MMControl1.filename = "19.MID"
GoTo 9000
390 frmOpen.MMControl1.filename = "20.MID"
GoTo 9000
410 frmOpen.MMControl1.filename = "21.MID"
GoTo 9000
430 frmOpen.MMControl1.filename = "22.MID"
GoTo 9000
450 frmOpen.MMControl1.filename = "23.MID"
GoTo 9000
470 frmOpen.MMControl1.filename = "24.MID"
GoTo 9000
490 frmOpen.MMControl1.filename = "25.MID"
GoTo 9000
510 frmOpen.MMControl1.filename = "26.MID"
GoTo 9000
530 frmOpen.MMControl1.filename = "27.MID"
GoTo 9000
550 frmOpen.MMControl1.filename = "28.MID"
GoTo 9000
570 frmOpen.MMControl1.filename = "29.MID"
GoTo 9000
590 frmOpen.MMControl1.filename = "30.MID"
GoTo 9000
610 frmOpen.MMControl1.filename = "31.MID"
GoTo 9000
630 frmOpen.MMControl1.filename = "32.MID"
GoTo 9000
650 frmOpen.MMControl1.filename = "33.MID"
GoTo 9000
670 frmOpen.MMControl1.filename = "34.MID"
GoTo 9000
690 frmOpen.MMControl1.filename = "35.MID"
GoTo 9000
710 frmOpen.MMControl1.filename = "36.MID"
GoTo 9000
9000 With frmOpen.MMControl1
    .Command = "Open"
    .Command = "play"
End With
9001
End Function

Private Sub mnuFlags_Click()
Option2
End Sub

Private Sub MnuGuess_Click()
FrmGuessTheWord.Show
End Sub

Private Sub mnuHiScore_Click()
JU = 1
frmHiScore.Show 1
Unload frmHiScore
JU = 0
End Sub



Private Sub mnuLoadOrder_Click()
With frmOpen.MMControl1
        .Shareable = False
        .DeviceType = "Sequencer"
        .Command = "close"
        .filename = "38.MID"
        .Command = "Open"
        .Command = "play"
    End With
CommonDialog1.HelpFile = "Info.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub



Private Sub mnuNewGame_Click()
ChangeLevels
End Sub

Private Sub mnuNoTunes_Click()
If SA = 1 Then GoTo 11
ST = 0: CT = 0: LT = 0: SA = 1
Mnucontinuous.Checked = False
mnuShortTunes.Checked = False
mnuNoTunes.Checked = True
mnuComputerWord.Enabled = True
mnuDisableWilds.Enabled = True
DL = 1000
DM = 1000
ChangeLevels
11
End Sub

Private Sub mnuNumbers_Click()
Option3
End Sub

Private Sub mnuOrdering_Click()
With frmOpen.MMControl1
        .Shareable = False
        .DeviceType = "Sequencer"
        .Command = "close"
        .filename = "76.MID"
        .Command = "Open"
        .Command = "play"
    End With
    JK = 1
frmAdvert.Show 1
If JJ = 1 Then GoTo 13
11 CloseDown
12 End
13
End Sub

Private Sub mnuPlayCards_Click()
Option5
End Sub



Private Sub mnuShortTunes_Click()
If ST = 1 Then GoTo 11
ST = 1: CT = 0: LT = 0: SA = 0
If OPT = 1 Then
AX = 0
W = 1
mnuDisableWilds.Enabled = True
mnuDisableWilds.Checked = False
End If
Mnucontinuous.Checked = False
mnuShortTunes.Checked = True
mnuNoTunes.Checked = False
mnuComputerWord.Enabled = True
DL = 1820
DM = 2325
ChangeLevels
11
End Sub

Private Sub mnuShowHelp_Click()
CommonDialog1.HelpFile = "flip.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub

Private Sub mnuUserWord_Click()
UW = 1
mnuUserWord.Checked = True
mnuComputerWord.Checked = False
ChangeLevels
End Sub

Private Sub mnuValidate_Click()
Dim UIP
UIP = InputBox("Please enter the Validation Code exactly as shown on your card. If you don't have a card, please click Cancel and see Help, Ordering Info.", "Validating Flip!")
If UIP = "" Then Exit Sub
If UIP = Trim("30169862") Then
LN = UIP
V = 1
'Text2.Text = Str(V)
MsgBox "Flip! has been validated, everything will now work!  Thank you for purchasing Flip!", vbInformation
Upgrade
End If
If UIP <> "30169862" Then
MsgBox "That is not a valid Code", vbCritical
End
End If
End Sub




Private Sub Pic1player_Click()
Setup1Player
End Sub

Private Sub pic2player_Click()
Setup2Players
End Sub

Private Sub Pic3x8_Click()
Level1
End Sub

Private Sub pic4x10_Click()
Level2
End Sub


Private Sub pic6x12_Click()
Level3
End Sub

Private Sub PicCards_Click()
Option5
End Sub

Private Sub Piccols_Click()
Option6
End Sub

Private Sub PicFlag_Click()
Option2
End Sub


Private Sub picMov_Click()
Option1
End Sub

Private Sub PicNums_Click()
Option3
End Sub

Private Sub PicRS_Click()
PicRS.Picture = ImageList2.ListImages(10).Picture
ChangeLevels
PicRS.Picture = ImageList2.ListImages(9).Picture
If CT = 1 Then
lblTries.Visible = False
With lblNana
.Visible = True
.Caption = "Clear the Board:"
.ForeColor = RGB(255, 255, 0)
End With
DL = 500
DM = 1000
Timer5.Enabled = True
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
.filename = "37.MID"
.Command = "Open"
.Command = "play"
End With
End If
End Sub

Private Sub Picture10_Click()
If V = 1 Then GoTo 11
If V = 0 Then
With frmOpen.MMControl1
        .Shareable = False
        .DeviceType = "Sequencer"
        .Command = "close"
        .filename = "76.MID"
        .Command = "Open"
        .Command = "play"
    End With
    JK = 1
frmAdvert.Show 1
If JJ = 1 Then GoTo 13
End If
11 CloseDown
12 End
13
End Sub
Private Sub Timer1_Timer()
If CT = 1 Then GoTo 9000
lblplayer1.Visible = False
lblplayer2.Visible = False
Shpplay1.Visible = False
shpplay2.Visible = False
lblTries.Visible = False
lblNana.Visible = True
If PX = 1 Then Timer1.Interval = 4000
HK = HK + 1
CL = CL + 1
If PX = 0 Then
If CL = 2 Then CL = 0
End If
If PX = 1 Then
If CL = 3 Then CL = 0
End If
If CL = 1 Then
For J = 1 To Z
G(J).Visible = True
Q(J) = 3
Next
lblNana.Caption = " "
End If
If CL = 0 Then
For J = 1 To Z
G(J).Visible = False
Next
If PX = 0 Then
lblNana.Caption = "The Word"
txtWord.Text = Word
End If
If PX = 1 Then
lblNana.Caption = Player1 + "- Your Word"
txtWord.Text = Word2
End If
End If
If CL = 2 Then
For J = 1 To Z
G(J).Visible = False
Next
lblNana.Caption = Player2 + "- Your Word"
txtWord.Text = Word1
End If
If HK = 1 Then
For J = 1 To Z
Q(J) = 1
G(J).Enabled = False
G(J).Visible = True
Next
End If
If EX = 1 Then GoTo 15
CheckScores
15
If HK > 100 Then HK = 100
If JK = 0 Then
If JU = 0 Then

End If
End If
9000
End Sub
Private Sub Timer2_Timer()
JJ = 0
If OPT = 1 Then GoTo 10
If OPT = 2 Then GoTo 20
If OPT = 3 Then GoTo 30
If OPT = 4 Then GoTo 40
If OPT = 5 Then GoTo 50
If OPT = 6 Then GoTo 60
If OPT = 7 Then GoTo 60
10 Animated
GoTo 100
20 For DW = 1 To Z
If Q(DW) >= 1 Then
G(DW).Picture = ImageList1.ListImages(FT(DW)).Picture
End If
Next
GoTo 100
30 For DW = 1 To Z
If Q(DW) >= 1 Then
G(DW).Picture = ImageList1.ListImages(FU(DW)).Picture
End If
Next
GoTo 100
40 For DW = 1 To Z
If Q(DW) >= 1 Then
G(DW).Picture = ImageList1.ListImages(FT(DW)).Picture
End If
Next
GoTo 100
50 For DW = 1 To Z
If Q(DW) >= 1 Then
G(DW).Picture = ImageList1.ListImages(FV(DW)).Picture
End If
Next
GoTo 100
60 For DW = 1 To Z
If Q(DW) >= 1 Then
G(DW).Picture = ImageList1.ListImages(325).Picture
End If
Next
GoTo 100
100
End Sub

Public Function Dictionary()
If UW = 1 Then
GetUserWords
GoTo 99
End If
10 Randomize
If PX = 0 Then
If R = 1 Then
O = Int(Rnd * 301) + 1
If O > 301 Then GoTo 10
Word = DC(O)
txtWord = Word
Text3.Text = Word
End If
If R = 2 Then
SP = Int(Rnd * 2) + 1
If SP = 1 Then
O = Int(Rnd * 301) + 1
If O > 301 Then GoTo 10
Word = DC(O)
txtWord = Word
Text3.Text = Word
End If
If SP = 2 Then
O = Int(Rnd * 301) + 1
Word = DD(O)
txtWord = Word
Text3.Text = Word
End If
End If
If R = 3 Then
SP = Int(Rnd * 3) + 1
If SP = 1 Then
O = Int(Rnd * 301) + 1
Word = DC(O)
txtWord = Word
Text3.Text = Word
End If
If SP = 2 Then
O = Int(Rnd * 301) + 1
Word = DD(O)
txtWord = Word
Text3.Text = Word
End If
If SP = 3 Then
O = Int(Rnd * 301) + 1
Word = DE(O)
txtWord = Word
Text3.Text = Word
End If
End If
If Word = "" Then GoTo 10
End If
11 If PX = 1 Then
If R = 1 Then
O = Int(Rnd * 301) + 1
If O > 301 Then GoTo 10
Word1 = DC(O)
End If
If R = 2 Then
SP = Int(Rnd * 2) + 1
If SP = 1 Then
O = Int(Rnd * 301) + 1
If O > 301 Then GoTo 10
Word1 = DC(O)
End If
If SP = 2 Then
O = Int(Rnd * 301) + 1
Word1 = DD(O)
End If
End If
If R = 3 Then
SP = Int(Rnd * 3) + 1
If SP = 1 Then
O = Int(Rnd * 301) + 1
Word1 = DC(O)
End If
If SP = 2 Then
O = Int(Rnd * 301) + 1
Word1 = DD(O)
End If
If SP = 3 Then
O = Int(Rnd * 301) + 1
Word1 = DE(O)
End If
End If
If Word1 = "" Then GoTo 11
End If
12 If R = 1 Then
O = Int(Rnd * 301) + 1
If O > 301 Then GoTo 10
Word2 = DC(O)
End If
If R = 2 Then
SP = Int(Rnd * 2) + 1
If SP = 1 Then
O = Int(Rnd * 301) + 1
If O > 301 Then GoTo 10
Word2 = DC(O)
End If
If SP = 2 Then
O = Int(Rnd * 301) + 1
Word2 = DD(O)
End If
End If
If R = 3 Then
SP = Int(Rnd * 3) + 1
If SP = 1 Then
O = Int(Rnd * 301) + 1
Word2 = DC(O)
End If
If SP = 2 Then
O = Int(Rnd * 301) + 1
Word2 = DD(O)
End If
If SP = 3 Then
O = Int(Rnd * 301) + 1
Word2 = DE(O)
End If
If Word2 = "" Then GoTo 12
End If
99
End Function
Public Function WProc()
For J = 1 To Z
If Q(J) = 1 Then
Q(J) = 2
End If
Next
If Q(23) = 1 Then Q(23) = 2
If Q(24) = 1 Then Q(24) = 2
U(23) = 3
PlayTuneB
U(23) = 0
W = 0
If PX = 1 Then
If SW = 0 Then
SW = 1
GoTo 6
End If
If SW = 1 Then
SW = 0
End If
End If
6 GuessTheWord
End Function


Private Sub Timer3_Timer()
If MHD = 1 Then GoTo 10
If Timer1.Enabled = True Then GoTo 10
MouseEnable
10
End Sub















Public Function Option2()
If OPT = 2 Then GoTo 10
OPT = 2
AX = 1
W = 0
MnuAnything.Checked = False
mnuFlags.Checked = True
mnuNumbers.Checked = False
mnuPlayCards.Checked = False
MnuColors.Checked = False
mnuColors2.Checked = False
picMov.Picture = ImageList2.ListImages(11).Picture
PicFlag.Picture = ImageList2.ListImages(14).Picture
PicNums.Picture = ImageList2.ListImages(20).Picture
PicCards.Picture = ImageList2.ListImages(15).Picture
Piccols.Picture = ImageList2.ListImages(17).Picture
ChangeLevels

10
End Function

Public Function Option1()
If OPT = 1 Then GoTo 10
OPT = 1
AX = 0
MnuAnything.Checked = True
mnuFlags.Checked = False
mnuNumbers.Checked = False
mnuPlayCards.Checked = False
MnuColors.Checked = False
mnuColors2.Checked = False
picMov.Picture = ImageList2.ListImages(12).Picture
PicFlag.Picture = ImageList2.ListImages(13).Picture
PicNums.Picture = ImageList2.ListImages(20).Picture
PicCards.Picture = ImageList2.ListImages(15).Picture
Piccols.Picture = ImageList2.ListImages(17).Picture
ChangeLevels

10
End Function

Public Function Level1()
If R = 1 Then GoTo 10
R = 1
Pic3x8.Picture = ImageList2.ListImages(4).Picture
pic4x10.Picture = ImageList2.ListImages(5).Picture
pic6x12.Picture = ImageList2.ListImages(7).Picture
mnu3X8.Checked = True
mnu4X10.Checked = False
mnu5X12.Checked = False
ChangeLevels

10
End Function

Public Function Level2()
If R = 2 Then GoTo 10
R = 2
Pic3x8.Picture = ImageList2.ListImages(3).Picture
pic4x10.Picture = ImageList2.ListImages(6).Picture
pic6x12.Picture = ImageList2.ListImages(7).Picture
mnu3X8.Checked = False
mnu4X10.Checked = True
mnu5X12.Checked = False
ChangeLevels

10
End Function


Public Function Level3()
If R = 3 Then GoTo 10
R = 3
Pic3x8.Picture = ImageList2.ListImages(3).Picture
pic4x10.Picture = ImageList2.ListImages(5).Picture
pic6x12.Picture = ImageList2.ListImages(8).Picture
mnu3X8.Checked = False
mnu4X10.Checked = False
mnu5X12.Checked = True
ChangeLevels
10
End Function

Public Function CheckScores()
If EX = 1 Then GoTo 500
EX = 1
If CT = 1 Then GoTo 300
If PX = 1 Then GoTo 301
If R = 1 Then
If GO > LE And GO < LK Then GoTo 301
End If
If R = 2 Then
If GO > LF And GO < LL Then GoTo 301
End If
If R = 3 Then
If GO > LG And GO < LM Then GoTo 301
End If
If R = 1 Then
If GO < LE Then
LE = GO
GoTo 100
End If
If GO > LK Then
LK = GO
GoTo 150
End If
End If
If R = 2 Then
If GO < LF Then
LF = GO
GoTo 100
End If
If GO > LL Then
LL = GO
GoTo 150
End If
End If
If R = 3 Then
If GO < LG Then
LG = GO
GoTo 100
End If
If GO > LM Then
LM = GO
GoTo 150
End If
End If
100 HAA = InputBox("Congratulations! Please enter your name", "For the Best Scores Table")
If R = 1 Then
LB = StrConv(HAA, vbProperCase)
GoTo 499
End If
If R = 2 Then
LC = StrConv(HAA, vbProperCase)
GoTo 499
End If
If R = 3 Then
LD = StrConv(HAA, vbProperCase)
GoTo 499
End If
150 HAA = InputBox("Congratulations! Please enter your name", "For the Worst Scores Table")
If R = 1 Then
LH = StrConv(HAA, vbProperCase)
GoTo 499
End If
If R = 2 Then
LI = StrConv(HAA, vbProperCase)
GoTo 499
End If
If R = 3 Then
LJ = StrConv(HAA, vbProperCase)
GoTo 499
End If
300 If R = 1 Then
If SE < JL Then
JL = SE
HAA = InputBox("Congratulations! Please enter your name", "For the Best Times Score Table")
LO = StrConv(HAA, vbProperCase)
GoTo 499
End If
End If
If R = 2 Then
If SE < JM Then
JM = SE
HAA = InputBox("Congratulations! Please enter your name", "For the Best Times Score Table")
LP = StrConv(HAA, vbProperCase)
GoTo 499
End If
End If
If R = 3 Then
If SE < JN Then
JN = SE
HAA = InputBox("Congratulations! Please enter your name", "For the Best Times Score Table")
LQ = StrConv(HAA, vbProperCase)
GoTo 499
End If
End If
lblNana.Visible = True
lblNana.Caption = "Well done, You did it in " + Str(SE) + " Seconds"
GoTo 500
301 If CT = 0 Then
If PX = 0 Then
Randomize
CH = Int(Rnd * 14) + 1
If SW = 1 Then
lblNana.Caption = CK(CH)
End If
End If
If PX = 1 Then
Randomize
CH = Int(Rnd * 14) + 1
If SW = 1 Then
lblNana.Caption = CJ(CH) + Player2
GA2 = GA2 + 1
End If
If SW = 0 Then
lblNana.Caption = CJ(CH) + Player1
GA1 = GA1 + 1
End If
End If
End If
GoTo 500
499 JU = 1
frmHiScore.Show 1
Unload frmHiScore
JU = 0
500
End Function
Public Function NewGame()
GO = 0
SX = 100
SE = 0
SW = 0
PL1 = 0
PL2 = 0
CX = 0
W = 0
L = 0
EX = 0
MHD = 2
HK = 0
CH = 0
CL = 0
Timer1.Interval = 9000
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = True
lblTries.Caption = "Number of tries: " + Str(GO)
lblplayer1.Caption = Player1 + "  " + Str(PL1) + "     " + Str(GA1)
lblplayer2.Caption = Player2 + "  " + Str(PL2) + "     " + Str(GA2)
For J = 1 To 72
Q(J) = 0
G(J).Enabled = True
N(J) = 0
K(J) = 0
U(J) = 0
G(J).Picture = ImageList2.ListImages(1).Picture
Next
lblNana.ForeColor = RGB(255, 255, 255)
Timer5.Interval = 3000
If CT = 0 Then
Timer5.Enabled = False
End If
MousePointer = 2
End Function
Public Function Option3()
If OPT = 3 Then GoTo 10
OPT = 3
AX = 1
MnuAnything.Checked = False
mnuFlags.Checked = False
mnuNumbers.Checked = True
mnuPlayCards.Checked = False
MnuColors.Checked = False
mnuColors2.Checked = False
picMov.Picture = ImageList2.ListImages(11).Picture
PicFlag.Picture = ImageList2.ListImages(13).Picture
PicNums.Picture = ImageList2.ListImages(19).Picture
PicCards.Picture = ImageList2.ListImages(15).Picture
Piccols.Picture = ImageList2.ListImages(17).Picture
ChangeLevels

10
End Function



Public Function Option5()
If OPT = 5 Then GoTo 10
OPT = 5
AX = 1
MnuAnything.Checked = False
mnuFlags.Checked = False
mnuNumbers.Checked = False
mnuPlayCards.Checked = True
MnuColors.Checked = False
mnuColors2.Checked = False
picMov.Picture = ImageList2.ListImages(11).Picture
PicFlag.Picture = ImageList2.ListImages(13).Picture
PicNums.Picture = ImageList2.ListImages(20).Picture
PicCards.Picture = ImageList2.ListImages(16).Picture
Piccols.Picture = ImageList2.ListImages(17).Picture
ChangeLevels

10
End Function

Public Function Option6()
If OPT = 6 Then GoTo 100
OPT = 6
AX = 1
MnuAnything.Checked = False
mnuFlags.Checked = False
mnuNumbers.Checked = False
mnuPlayCards.Checked = False
MnuColors.Checked = True
mnuColors2.Checked = False
picMov.Picture = ImageList2.ListImages(11).Picture
PicFlag.Picture = ImageList2.ListImages(13).Picture
PicNums.Picture = ImageList2.ListImages(20).Picture
PicCards.Picture = ImageList2.ListImages(15).Picture
Piccols.Picture = ImageList2.ListImages(18).Picture
G(1).BackColor = RGB(0, 0, 255)
G(2).BackColor = RGB(0, 0, 255)
G(3).BackColor = RGB(255, 0, 0)
G(4).BackColor = RGB(255, 0, 0)
G(5).BackColor = RGB(0, 255, 0)
G(6).BackColor = RGB(0, 255, 0)
G(7).BackColor = RGB(255, 255, 0)
G(8).BackColor = RGB(255, 255, 0)
G(9).BackColor = RGB(0, 255, 255)
G(10).BackColor = RGB(0, 255, 255)
G(11).BackColor = RGB(255, 255, 255)
G(12).BackColor = RGB(255, 255, 255)
G(13).BackColor = RGB(255, 0, 255)
G(14).BackColor = RGB(255, 0, 255)
G(15).BackColor = RGB(255, 0, 90)
G(16).BackColor = RGB(255, 0, 90)
G(17).BackColor = RGB(255, 255, 113)
G(18).BackColor = RGB(255, 255, 113)
G(19).BackColor = RGB(130, 255, 113)
G(20).BackColor = RGB(130, 255, 113)
G(21).BackColor = RGB(130, 130, 113)
G(22).BackColor = RGB(130, 130, 113)
G(23).BackColor = RGB(130, 130, 0)
G(24).BackColor = RGB(130, 130, 0)
G(25).BackColor = RGB(130, 64, 0)
G(26).BackColor = RGB(130, 64, 0)
G(27).BackColor = RGB(255, 130, 0)
G(28).BackColor = RGB(255, 130, 0)
G(29).BackColor = RGB(130, 64, 64)
G(30).BackColor = RGB(130, 64, 64)
G(31).BackColor = RGB(255, 130, 130)
G(32).BackColor = RGB(255, 130, 130)
G(33).BackColor = RGB(129, 0, 255)
G(34).BackColor = RGB(129, 0, 255)
G(35).BackColor = RGB(0, 0, 0)
G(36).BackColor = RGB(0, 0, 0)
G(37).BackColor = RGB(130, 0, 82)
G(38).BackColor = RGB(130, 0, 82)
G(39).BackColor = RGB(130, 0, 0)
G(40).BackColor = RGB(130, 0, 0)
G(41).BackColor = RGB(255, 124, 66)
G(42).BackColor = RGB(255, 124, 66)
G(43).BackColor = RGB(0, 255, 50)
G(44).BackColor = RGB(0, 255, 50)
G(45).BackColor = RGB(0, 0, 70)
G(46).BackColor = RGB(0, 0, 70)
G(47).BackColor = RGB(194, 129, 128)
G(48).BackColor = RGB(194, 129, 128)
G(49).BackColor = RGB(194, 50, 194)
G(50).BackColor = RGB(194, 50, 194)
G(51).BackColor = RGB(128, 129, 47)
G(52).BackColor = RGB(128, 129, 47)
G(53).BackColor = RGB(220, 165, 155)
G(54).BackColor = RGB(220, 165, 155)
G(55).BackColor = RGB(255, 128, 0)
G(56).BackColor = RGB(255, 128, 0)
G(57).BackColor = RGB(255, 255, 189)
G(58).BackColor = RGB(255, 255, 189)
G(59).BackColor = RGB(61, 192, 99)
G(60).BackColor = RGB(61, 192, 99)
G(61).BackColor = RGB(0, 128, 255)
G(62).BackColor = RGB(0, 128, 255)
G(63).BackColor = RGB(61, 192, 160)
G(64).BackColor = RGB(61, 192, 160)
G(65).BackColor = RGB(0, 255, 182)
G(66).BackColor = RGB(0, 255, 182)
G(67).BackColor = RGB(255, 0, 23)
G(68).BackColor = RGB(255, 0, 23)
G(69).BackColor = RGB(64, 0, 161)
G(70).BackColor = RGB(64, 0, 161)
G(71).BackColor = RGB(220, 165, 0)
G(72).BackColor = RGB(220, 165, 0)
ChangeLevels
100
End Function

Public Function CloseDown()
Fd = 1
FA.Name = LA
FA.Begin1 = LB
FA.Inter1 = LC
FA.Expert1 = LD
FA.BScore1 = LE
FA.BScore2 = LF
FA.BScore3 = LG
FA.Begin2 = LH
FA.Inter2 = LI
FA.Expert2 = LJ
FA.WScore1 = LK
FA.WScore2 = LL
FA.WScore2 = LM
FA.Validate = LN
FA.TName1 = LO
FA.Tname2 = LP
FA.Tname3 = LQ
FA.BTime1 = JL
FA.BTime2 = JM
FA.BTime3 = JN
Put #FB, Fd, FA
Close FB
End Function

Private Sub Timer4_Timer()
If CT = 1 Then GoTo 9000
lblNana.Visible = True
lblTries.Visible = False
Randomize
CH = Int(Rnd * 37) + 1
lblNana.Caption = CI(CH)
lblplayer1.Visible = False
lblplayer2.Visible = False
Shpplay1.Visible = False
shpplay2.Visible = False
For J = 1 To Z
If Q(J) = 0 Then
G(J).Enabled = True
End If
Next
MousePointer = 2
MHD = 2
9000 Timer4.Enabled = False
End Sub



Public Function Option7()
If OPT = 7 Then GoTo 100
OPT = 7
AX = 1
MnuAnything.Checked = False
mnuFlags.Checked = False
mnuNumbers.Checked = False
mnuPlayCards.Checked = False
MnuColors.Checked = False
mnuColors2.Checked = True
picMov.Picture = ImageList2.ListImages(11).Picture
PicFlag.Picture = ImageList2.ListImages(13).Picture
PicNums.Picture = ImageList2.ListImages(20).Picture
PicCards.Picture = ImageList2.ListImages(15).Picture
Piccols.Picture = ImageList2.ListImages(18).Picture
Randomize
For J = 1 To 71 Step 2
Randomize
PA(1) = Int(Rnd * 255) + 1
PA(2) = Int(Rnd * 255) + 1
PA(3) = Int(Rnd * 255) + 1
G(J).BackColor = RGB(PA(1), PA(2), PA(3))
G(J + 1).BackColor = RGB(PA(1), PA(2), PA(3))
Next
ChangeLevels
100
End Function

Public Function Wordlist()
DC(1) = "STONE"
DC(2) = "ROUND"
DC(3) = "BOUND"
DC(4) = "SOUND"
DC(5) = "HANDS"
DC(6) = "HALVE"
DC(7) = "GUYOT"
DC(8) = "GUNGE"
DC(9) = "GUSTY"
DC(10) = "GULLY"
DC(11) = "FOLLY"
DC(12) = "FOCUS"
DC(13) = "FLYER"
DC(14) = "FLUTE"
DC(15) = "FLOUR"
DC(16) = "FLUFF"
DC(17) = "EARLY"
DC(18) = "DRYER"
DC(19) = "DRONE"
DC(20) = "DRIVE"
DC(21) = "DRILL"
DC(22) = "DRAWN"
DC(23) = "DOZEN"
DC(24) = "DOUSE"
DC(25) = "CLOWN"
DC(26) = "STAMP"
DC(27) = "CHILD"
DC(28) = "SHEEP"
DC(29) = "DINGY"
DC(30) = "DINGO"
DC(31) = "DINER"
DC(32) = "DIGIT"
DC(33) = "DIARY"
DC(34) = "DAIRY"
DC(35) = "CURRY"
DC(36) = "HURRY"
DC(37) = "FURRY"
DC(38) = "CUDDY"
DC(39) = "CRUST"
DC(40) = "COSET"
DC(41) = "GROUP"
DC(42) = "FRUIT"
DC(43) = "CLOUD"
DC(44) = "CLOCK"
DC(45) = "CLEAR"
DC(46) = "PIANO"
DC(47) = "CAVIL"
DC(48) = "CARRY"
DC(49) = "MARRY"
DC(50) = "CAROM"
DC(51) = "CARDS"
DC(52) = "CAPON"
DC(53) = "BUSHY"
DC(54) = "BURRO"
DC(55) = "BOUGH"
DC(56) = "ROUGH"
DC(57) = "TOUGH"
DC(58) = "BOSUN"
DC(59) = "BOOTH"
DC(60) = "BINGO"
DC(61) = "BATCH"
DC(62) = "HATCH"
DC(63) = "MATCH"
DC(64) = "CATCH"
DC(65) = "LATCH"
DC(66) = "PATCH"
DC(67) = "YACHT"
DC(68) = "BASAL"
DC(69) = "BANJO"
DC(70) = "AUDIO"
DC(71) = "ARENA"
DC(72) = "ANION"
DC(73) = "ONION"
DC(74) = "ANIME"
DC(75) = "ANGEL"
DC(76) = "ANKLE"
DC(77) = "AMUSE"
DC(78) = "AMICE"
DC(79) = "AMBER"
DC(80) = "ALPHA"
DC(81) = "ALOFT"
DC(82) = "AGENT"
DC(83) = "JOULE"
DC(84) = "JOINT"
DC(85) = "POINT"
DC(86) = "QUOIT"
DC(87) = "CRACK"
DC(88) = "PRANK"
DC(89) = "DRANK"
DC(90) = "KETCH"
DC(91) = "PRUNE"
DC(92) = "MOUSE"
DC(93) = "HOUSE"
DC(94) = "RIGHT"
DC(95) = "PLAIN"
DC(96) = "TRAIN"
DC(97) = "EARTH"
DC(98) = "GREEN"
DC(99) = "BROWN"
DC(100) = "BLACK"
DC(101) = "HORSE"
DC(102) = "QUEEN"
DC(103) = "LADLE"
DC(104) = "LASER"
DC(105) = "RADAR"
DC(106) = "LADEN"
DC(107) = "MONEY"
DC(108) = "MONTH"
DC(109) = "SNAKE"
DC(110) = "MOVIE"
DC(111) = "MUSIC"
DC(112) = "FLUTE"
DC(113) = "MURAL"
DC(114) = "NORTH"
DC(115) = "SOUTH"
DC(116) = "PILOT"
DC(117) = "PITOT"
DC(118) = "PITCH"
DC(119) = "HITCH"
DC(120) = "PITON"
DC(121) = "FRONT"
DC(122) = "POLAR"
DC(123) = "PROBE"
DC(124) = "CHAIR"
DC(125) = "TABLE"
DC(126) = "FLOOR"
DC(127) = "ROCKY"
DC(128) = "DOUGH"
DC(129) = "SALON"
DC(130) = "SARAN"
DC(131) = "SAXON"
DC(132) = "SUGAR"
DC(133) = "BREAD"
DC(134) = "TOAST"
DC(135) = "SCOLD"
DC(136) = "SCRIM"
DC(137) = "OTTER"
DC(138) = "OTHER"
DC(139) = "FIRST"
DC(140) = "SENSE"
DC(141) = "SHARE"
DC(142) = "SMART"
DC(143) = "SPARK"
DC(144) = "ROAST"
DC(145) = "STORM"
DC(146) = "PLANT"
DC(147) = "TEETH"
DC(148) = "TEPID"
DC(149) = "SEVEN"
DC(150) = "EIGHT"
DC(151) = "THREE"
DC(152) = "FIFTY"
DC(153) = "TIBIA"
DC(154) = "SKULL"
DC(155) = "FEMUR"
DC(156) = "THUMB"
DC(157) = "BONES"
DC(158) = "ELBOW"
DC(159) = "TIMER"
DC(160) = "SLEEP"
DC(161) = "AWAKE"
DC(162) = "TRACK"
DC(163) = "TRUCK"
DC(164) = "VALUE"
DC(165) = "ORGAN"
DC(166) = "RIVER"
DC(167) = "TULIP"
DC(168) = "MARCH"
DC(169) = "APRIL"
DC(170) = "TWIST"
DC(171) = "TWIRL"
DC(172) = "TWERP"
DC(173) = "UMBER"
DC(174) = "UNBAR"
DC(175) = "UNDER"
DC(176) = "ABOVE"
DC(177) = "UNITE"
DC(178) = "UNITY"
DC(179) = "LEARN"
DC(180) = "WORLD"
DC(181) = "BATHE"
DC(182) = "SWEEP"
DC(183) = "SURGE"
DC(184) = "PURGE"
DC(185) = "MERGE"
DC(186) = "BARGE"
DC(187) = "STAGE"
DC(188) = "VALET"
DC(189) = "WATER"
DC(190) = "VELUM"
DC(191) = "VENUE"
DC(192) = "FIZZY"
DC(193) = "DRINK"
DC(194) = "SHAKE"
DC(195) = "BREAK"
DC(196) = "BRAKE"
DC(197) = "VOGUE"
DC(198) = "VAGUE"
DC(199) = "WAIST"
DC(200) = "YEAST"
DC(201) = "LEMON"
DC(202) = "TRICK"
DC(203) = "SONGS"
DC(204) = "TUNES"
DC(205) = "DREAM"
DC(206) = "HEAVY"
DC(207) = "HONEY"
DC(207) = "FUNNY"
DC(208) = "BUNNY"
DC(209) = "SUNNY"
DC(210) = "AGAIN"
DC(211) = "RAINY"
DC(212) = "HAPPY"
DC(213) = "CANDY"
DC(214) = "HANDY"
DC(215) = "NIGHT"
DC(216) = "FIGHT"
DC(217) = "MIGHT"
DC(218) = "LIGHT"
DC(219) = "HEART"
DC(220) = "WHOLE"
DC(221) = "TEACH"
DC(222) = "PLACE"
DC(223) = "LEAVE"
DC(224) = "NEVER"
DC(225) = "LEVER"
DC(226) = "FEVER"
DC(227) = "DANCE"
DC(228) = "LANCE"
DC(229) = "COULD"
DC(230) = "MOULD"
DC(231) = "TRULY"
DC(232) = "START"
DC(233) = "HOUND"
DC(234) = "WHERE"
DC(235) = "LOOSE"
DC(236) = "MOOSE"
DC(237) = "SMILE"
DC(238) = "MELON"
DC(239) = "BRICK"
DC(240) = "FETCH"
DC(241) = "VALVE"
DC(242) = "CRISP"
DC(243) = "UNTIL"
DC(244) = "BEANS"
DC(245) = "UNTIL"
DC(246) = "BRAIN"
DC(247) = "RINGS"
DC(248) = "LUCKY"
DC(249) = "SPEAK"
DC(250) = "GOING"
DC(251) = "SCONE"
DC(252) = "SHIRT"
DC(253) = "PANTS"
DC(254) = "CHILL"
DC(255) = "WORDS"
DC(256) = "SPLIT"
DC(257) = "SHIFT"
DC(258) = "SPOOF"
DC(259) = "SPOOK"
DC(260) = "SPOOL"
DC(261) = "SPRAY"
DC(262) = "STEAK"
DC(263) = "STAKE"
DC(264) = "SHOPS"
DC(265) = "CLOTH"
DC(266) = "BOOKS"
DC(267) = "GRACE"
DC(268) = "ALLEY"
DC(269) = "SKATE"
DC(270) = "MAJOR"
DC(271) = "STAVE"
DC(272) = "STAND"
DC(273) = "PAPER"
DC(274) = "ROSES"
DC(275) = "GRASS"
DC(276) = "AMAZE"
DC(277) = "POLKA"
DC(278) = "CHORD"
DC(279) = "RACES"
DC(280) = "MINOR"
DC(281) = "LATER"
DC(282) = "BLOCK"
DC(283) = "DRAKE"
DC(284) = "SMOKE"
DC(285) = "CONGA"
DC(286) = "BOARD"
DC(287) = "FLICK"
DC(288) = "GRADE"
DC(289) = "APPLE"
DC(290) = "THANK"
DC(291) = "STYLE"
DC(292) = "BENCH"
DC(293) = "PAINT"
DC(294) = "TRACE"
DC(295) = "ABAFT"
DC(296) = "USAGE"
DC(297) = "CLASS"
DC(298) = "BRUSH"
DC(299) = "BRASS"
DC(300) = "GRAPH"
DC(301) = "ZONAL"
DD(1) = "BLANCH"
DD(2) = "PLANET"
DD(3) = "BISHOP"
DD(4) = "BIONIC"
DD(5) = "BETRAY"
DD(6) = "BENIGN"
DD(7) = "PURPLE"
DD(8) = "ORANGE"
DD(9) = "BELIEF"
DD(10) = "BEETLE"
DD(11) = "PORTER"
DD(12) = "GOLFER"
DD(13) = "CASTLE"
DD(14) = "KEEPER"
DD(15) = "BASALT"
DD(16) = "APPLES"
DD(17) = "POTATO"
DD(18) = "TOMATO"
DD(19) = "MASTER"
DD(20) = "PUBLIC"
DD(21) = "SCHOOL"
DD(22) = "CACTUS"
DD(23) = "BANKER"
DD(24) = "BOTTLE"
DD(25) = "BALADE"
DD(26) = "RECORD"
DD(27) = "PLAYER"
DD(28) = "TAILOR"
DD(29) = "BAGDER"
DD(30) = "BABOON"
DD(31) = "MONKEY"
DD(32) = "AUTUMN"
DD(33) = "WINTER"
DD(34) = "MONDAY"
DD(35) = "FRIDAY"
DD(36) = "SUNDAY"
DD(37) = "AUGUST"
DD(38) = "IMMUNE"
DD(39) = "AUTHOR"
DD(40) = "ATTEND"
DD(41) = "ARCTIC"
DD(42) = "MEMBER"
DD(43) = "ASTRAL"
DD(44) = "SILVER"
DD(45) = "YELLOW"
DD(46) = "SPHERE"
DD(47) = "CIRCLE"
DD(48) = "TARMAC"
DD(49) = "CEMENT"
DD(50) = "FARMER"
DD(51) = "HUNTER"
DD(52) = "ASCEND"
DD(53) = "DANCER"
DD(54) = "SHOWER"
DD(55) = "FLIGHT"
DD(56) = "CANDLE"
DD(57) = "WINDOW"
DD(58) = "CHERRY"
DD(59) = "POLICE"
DD(60) = "AROUND"
DD(61) = "ARCADE"
DD(62) = "COWBOY"
DD(63) = "CARBON"
DD(64) = "OXYGEN"
DD(65) = "CARTON"
DD(66) = "CARROT"
DD(67) = "PIGEON"
DD(68) = "CARVER"
DD(69) = "CASUAL"
DD(70) = "WORKER"
DD(71) = "CAUDLE"
DD(72) = "CAUGHT"
DD(73) = "CENTER"
DD(74) = "HEATER"
DD(75) = "TIMBER"
DD(76) = "CHEESE"
DD(77) = "BUTTER"
DD(78) = "DYNODE"
DD(79) = "EATING"
DD(80) = "REWARD"
DD(81) = "EFFIGY"
DD(82) = "EFFETE"
DD(83) = "ELEVEN"
DD(84) = "TWELVE"
DD(85) = "ELAPSE"
DD(89) = "ELAPID"
DD(90) = "INDIAN"
DD(91) = "EMERGE"
DD(92) = "EMPIRE"
DD(93) = "ENERGY"
DD(94) = "FEEBLE"
DD(95) = "ENGLUT"
DD(96) = "ENOUGH"
DD(97) = "SHROUD"
DD(98) = "ENSIGN"
DD(99) = "ENTAIL"
DD(100) = "EOLITH"
DD(101) = "ENZYME"
DD(102) = "FLOWER"
DD(103) = "BASKET"
DD(104) = "EXCEPT"
DD(105) = "EXPEND"
DD(106) = "POWDER"
DD(107) = "FETTLE"
DD(108) = "KETTLE"
DD(109) = "CATTLE"
DD(110) = "FIDDLE"
DD(111) = "FIDGET"
DD(112) = "WIDGET"
DD(113) = "FINITE"
DD(114) = "CLAUSE"
DD(115) = "ESCAPE"
DD(116) = "WAIVER"
DD(117) = "LEADER"
DD(118) = "FLYING"
DD(119) = "BRIDGE"
DD(120) = "FONDUE"
DD(121) = "COURSE"
DD(122) = "FOSTER"
DD(123) = "FOSSIL"
DD(124) = "FREEZE"
DD(125) = "FRUITY"
DD(126) = "LAUNCH"
DD(127) = "DINNER"
DD(128) = "SUPPER"
DD(129) = "TONGUE"
DD(130) = "GIBLET"
DD(131) = "GIMLET"
DD(132) = "COOKER"
DD(133) = "GANDER"
DD(134) = "GLANCE"
DD(135) = "GINGER"
DD(136) = "GRANGE"
DD(137) = "RUSSIA"
DD(138) = "GREECE"
DD(139) = "GREASE"
DD(140) = "QUAVER"
DD(141) = "BRIEVE"
DD(142) = "TREBLE"
DD(143) = "BINDER"
DD(144) = "FINDER"
DD(145) = "HINDER"
DD(146) = "SINDER"
DD(147) = "TINDER"
DD(148) = "HOLDER"
DD(149) = "HOLDEN"
DD(150) = "HOLILY"
DD(151) = "CAMPUS"
DD(152) = "CORNER"
DD(153) = "CLOSED"
DD(154) = "BEAKER"
DD(155) = "HEADER"
DD(156) = "BUCKLE"
DD(157) = "HUCKLE"
DD(158) = "HUNTER"
DD(159) = "HUNTED"
DD(160) = "HUMBUG"
DD(161) = "HYBRID"
DD(162) = "HYBRIS"
DD(163) = "PEPPER"
DD(164) = "STATIC"
DD(165) = "IMMUNE"
DD(166) = "PROPER"
DD(167) = "ENCASE"
DD(168) = "INDENT"
DD(169) = "DIRECT"
DD(170) = "INDIGO"
DD(171) = "INDUCE"
DD(172) = "SCRIBE"
DD(173) = "INSECT"
DD(174) = "DISECT"
DD(175) = "SCRAPE"
DD(176) = "INVERT"
DD(177) = "ISLAND"
DD(178) = "JERKIN"
DD(179) = "JIGSAW"
DD(180) = "JIGGER"
DD(181) = "JETTON"
DD(182) = "JUNKET"
DD(183) = "JUNGLE"
DD(184) = "KENNEL"
DD(185) = "KNIGHT"
DD(186) = "STREAM"
DD(187) = "SCREAM"
DD(188) = "MANGER"
DD(189) = "MANGLE"
DD(190) = "MARINE"
DD(191) = "MARINA"
DD(192) = "ENGINE"
DD(193) = "PRINCE"
DD(194) = "MEDIUM"
DD(195) = "STRIKE"
DD(196) = "MELLOW"
DD(197) = "MEMORY"
DD(198) = "METHOD"
DD(199) = "HEAVEN"
DD(200) = "MIDDLE"
DD(201) = "MIDDEN"
DD(202) = "OBLIGE"
DD(203) = "OBLONG"
DD(204) = "OBTECT"
DD(205) = "OCTAVE"
DD(206) = "OCTAVO"
DD(207) = "OPAQUE"
DD(208) = "ORDAIN"
DD(209) = "ORDEAL"
DD(210) = "ORIENT"
DD(211) = "ORRERY"
DD(212) = "OTIOSE"
DD(213) = "OUTLET"
DD(214) = "SHADOW"
DD(215) = "PALACE"
DD(216) = "PAPACY"
DD(217) = "PARCEL"
DD(218) = "PARSON"
DD(219) = "PATENT"
DD(222) = "PATROL"
DD(223) = "OYSTER"
DD(224) = "PEBBLE"
DD(225) = "PERIOD"
DD(226) = "PERRON"
DD(227) = "PERSON"
DD(228) = "PIERCE"
DD(229) = "PILLAR"
DD(230) = "PILLOW"
DD(231) = "PIRATE"
DD(232) = "PIQUET"
DD(233) = "PIRACY"
DD(234) = "PITSAW"
DD(235) = "PLAQUE"
DD(236) = "PLUNGE"
DD(237) = "POCKET"
DD(238) = "PLURAL"
DD(239) = "RETURN"
DD(240) = "POISED"
DD(241) = "POLISH"
DD(242) = "HELMET"
DD(243) = "POTION"
DD(244) = "MOTION"
DD(245) = "POTTLE"
DD(246) = "PRAYER"
DD(247) = "BURSAR"
DD(248) = "PRIMAL"
DD(249) = "PRIMER"
DD(250) = "PROFIT"
DD(251) = "QUASAR"
DD(252) = "QUARTZ"
DD(253) = "RATTLE"
DD(254) = "RASTER"
DD(255) = "RECKON"
DD(256) = "CHROME"
DD(257) = "REGNAL"
DD(258) = "REGLET"
DD(259) = "RELISH"
DD(260) = "RENDER"
DD(261) = "REPEAT"
DD(262) = "REPORT"
DD(263) = "RESIST"
DD(264) = "RESILE"
DD(265) = "RETARD"
DD(266) = "RITUAL"
DD(267) = "ROCKER"
DD(268) = "ROCKET"
DD(269) = "FENDER"
DD(270) = "SHRINK"
DD(271) = "SHRIVE"
DD(272) = "SICKLE"
DD(273) = "SIGNAL"
DD(274) = "SIMILE"
DD(275) = "SIPHON"
DD(276) = "SINGLE"
DD(277) = "SISTER"
DD(278) = "SKETCH"
DD(279) = "SMUDGE"
DD(280) = "SODIUM"
DD(281) = "SOCKET"
DD(282) = "SOCIAL"
DD(283) = "CREDIT"
DD(284) = "SALMON"
DD(285) = "SYSTEM"
DD(286) = "NAPKIN"
DD(287) = "TENNIS"
DD(288) = "TABLET"
DD(289) = "TACKLE"
DD(290) = "COVERT"
DD(291) = "TALENT"
DD(292) = "UNICAL"
DD(293) = "CHURCH"
DD(294) = "FACTOR"
DD(295) = "MUPPET"
DD(296) = "UNITED"
DD(297) = "LOCKED"
DD(298) = "VACUUM"
DD(299) = "VANITY"
DD(300) = "VELVET"
DD(301) = "VESTRY"
DE(1) = "HABITAT"
DE(2) = "BIOLOGY"
DE(3) = "TRUMPET"
DE(4) = "CRICKET"
DE(5) = "SEASIDE"
DE(6) = "BUILDER"
DE(7) = "DELIGHT"
DE(8) = "FEATHER"
DE(9) = "BASSOON"
DE(10) = "ORANGES"
DE(11) = "PRINTER"
DE(12) = "MONITOR"
DE(13) = "PREFECT"
DE(14) = "PERFECT"
DE(15) = "TEACHER"
DE(16) = "BAPTIZE"
DE(17) = "BAPTIST"
DE(18) = "MANAGER"
DE(19) = "BALLOON"
DE(20) = "CONCERT"
DE(21) = "SPEAKER"
DE(22) = "BALANCE"
DE(23) = "PAYMENT"
DE(24) = "BRIGADE"
DE(25) = "BAGGAGE"
DE(26) = "TUESDAY"
DE(27) = "OCTOBER"
DE(28) = "LIBRARY"
DE(29) = "AUREATE"
DE(30) = "ATTAINT"
DE(31) = "AUCTION"
DE(32) = "ASPHALT"
DE(33) = "TRACTOR"
DE(34) = "TRAILER"
DE(35) = "ASPERSE"
DE(36) = "DESCENT"
DE(37) = "DESCEND"
DE(38) = "ARTWORK"
DE(39) = "ARTICLE"
DE(40) = "SEEKERS"
DE(41) = "SINGERS"
DE(42) = "DANCING"
DE(43) = "MILLION"
DE(44) = "HOARDER"
DE(45) = "CURTAIN"
DE(46) = "GARAGES"
DE(47) = "TRUCKER"
DE(48) = "SHOPPER"
DE(49) = "PUMPKIN"
DE(50) = "AUGMENT"
DE(51) = "COLLEGE"
DE(52) = "ARDUOUS"
DE(53) = "BLADDER"
DE(54) = "CANVASS"
DE(55) = "CANTEEN"
DE(56) = "CAPELIN"
DE(57) = "CAPTAIN"
DE(58) = "CAPTIVE"
DE(59) = "CAPITAL"
DE(60) = "CAPTURE"
DE(61) = "AMERICA"
DE(62) = "ENGLAND"
DE(63) = "GERMANY"
DE(64) = "CARLING"
DE(65) = "CARTOON"
DE(66) = "CARRIER"
DE(67) = "CASCADE"
DE(68) = "CASUIST"
DE(69) = "CATHODE"
DE(70) = "CROWDED"
DE(71) = "FISSURE"
DE(72) = "DAGLOCK"
DE(73) = "CAUDATE"
DE(74) = "CAUTION"
DE(75) = "CENTRAL"
DE(76) = "RESERVE"
DE(77) = "HEATING"
DE(78) = "HEATHER"
DE(79) = "CENTURY"
DE(80) = "CERTAIN"
DE(81) = "CHAMFER"
DE(82) = "CHAMBER"
DE(83) = "CHAMOIS"
DE(84) = "CHANTRY"
DE(85) = "CHANNEL"
DE(86) = "CHAPPAL"
DE(87) = "CHAPLET"
DE(88) = "DENSITY"
DE(89) = "CHARITY"
DE(90) = "CHATTER"
DE(91) = "CHATTEL"
DE(92) = "CHARTER"
DE(93) = "CHELATE"
DE(94) = "CHEESED"
DE(95) = "CHEMIST"
DE(96) = "DENTIST"
DE(97) = "EFFABLE"
DE(98) = "FIFTEEN"
DE(99) = "ELASTIC"
DE(100) = "PLASTIC"
DE(101) = "ELECTOR"
DE(102) = "EMERALD"
DE(103) = "DIAMOND"
DE(104) = "EMOTION"
DE(105) = "COUNTRY"
DE(106) = "ENGRAFT"
DE(107) = "MANSION"
DE(108) = "ENTASIS"
DE(109) = "ENVIRON"
DE(110) = "ENTOPIC"
DE(111) = "EPIDOTE"
DE(112) = "EVERTOR"
DE(113) = "EVENING"
DE(114) = "EXCLAIM"
DE(115) = "EXCLUDE"
DE(116) = "INHIBIT"
DE(117) = "EXHIBIT"
DE(118) = "FERTILE"
DE(119) = "FIDDLER"
DE(120) = "FINLAND"
DE(121) = "RORGUAL"
DE(122) = "MARSHAL"
DE(123) = "LANDING"
DE(124) = "FISSION"
DE(125) = "FLAUNCH"
DE(126) = "LEATHER"
DE(127) = "BLANKET"
DE(128) = "FOLIAGE"
DE(129) = "FOOTING"
DE(130) = "FORLORN"
DE(131) = "HUNDRED"
DE(132) = "THOUGHT"
DE(133) = "FREEZER"
DE(134) = "TURBINE"
DE(135) = "GIMBALS"
DE(136) = "PLATERS"
DE(137) = "GLARING"
DE(138) = "GLENOID"
DE(139) = "GNARLED"
DE(140) = "GLUCOSE"
DE(141) = "GLYPTIC"
DE(142) = "GRANITE"
DE(143) = "GRAPHIC"
DE(144) = "GRATING"
DE(145) = "GRINDER"
DE(146) = "BUTCHER"
DE(147) = "HEMIOLA"
DE(148) = "HERRING"
DE(149) = "HICKORY"
DE(150) = "HISTORY"
DE(151) = "SCIENCE"
DE(152) = "ENGLISH"
DE(153) = "HOLIDAY"
DE(154) = "HOLDING"
DE(155) = "BARRIER"
DE(156) = "FARRIER"
DE(157) = "HULKING"
DE(158) = "VINEGAR"
DE(159) = "IDOLIZE"
DE(160) = "IMAGINE"
DE(161) = "IMPETUS"
DE(162) = "IMPROVE"
DE(163) = "INCENSE"
DE(164) = "INBUILT"
DE(165) = "CONTEST"
DE(166) = "CONSUME"
DE(167) = "INFERNO"
DE(168) = "GENIOUS"
DE(169) = "INFLICT"
DE(170) = "INKLING"
DE(171) = "SOLVENT"
DE(172) = "FORWARD"
DE(173) = "INTERIM"
DE(174) = "INVERSE"
DE(175) = "ISOTOPE"
DE(176) = "JUGGLER"
DE(177) = "JUGULAR"
DE(178) = "JUSTIFY"
DE(179) = "HEARING"
DE(180) = "CERAMIC"
DE(181) = "KERATIN"
DE(182) = "KINETIC"
DE(183) = "LINSEED"
DE(184) = "PICTURE"
DE(185) = "SULPHUR"
DE(186) = "CUSTARD"
DE(187) = "LUNETTE"
DE(188) = "MAMMOTH"
DE(189) = "MANAGER"
DE(190) = "MARBLES"
DE(191) = "MIRACLE"
DE(192) = "MARQUEE"
DE(193) = "MARRIED"
DE(194) = "KNOCKER"
DE(195) = "MYSTERY"
DE(196) = "MACHINE"
DE(197) = "WASHING"
DE(198) = "MEDICAL"
DE(199) = "MEERKAT"
DE(200) = "MINIMUM"
DE(201) = "OBELISK"
DE(202) = "OBSERVE"
DE(203) = "OCELLUS"
DE(204) = "ODOACER"
DE(206) = "ODYSSEY"
DE(207) = "OFFICER"
DE(208) = "OPERATE"
DE(209) = "OPPIDAN"
DE(210) = "OPTICAL"
DE(211) = "OPTIMAL"
DE(212) = "OSCULAR"
DE(213) = "OTOLOGY"
DE(214) = "PALETTE"
DE(215) = "PARTIAL"
DE(216) = "PARTNER"
DE(217) = "PASTIME"
DE(218) = "PASTURE"
DE(219) = "PASSIVE"
DE(220) = "PATIENT"
DE(221) = "PATRIAL"
DE(222) = "PEACOCK"
DE(223) = "PEASANT"
DE(224) = "PENGUIN"
DE(225) = "PENSION"
DE(226) = "WHISTLE"
DE(227) = "PERFORM"
DE(228) = "PERFUME"
DE(229) = "PERFUSE"
DE(230) = "PERJURY"
DE(231) = "PICADOR"
DE(232) = "WRITING"
DE(233) = "PINBALL"
DE(234) = "PIRANHA"
DE(235) = "PLACOID"
DE(236) = "PLASTIC"
DE(237) = "PLASTER"
DE(238) = "PLUMMET"
DE(239) = "PLUNDER"
DE(240) = "POTICHE"
DE(241) = "TRAMMEL"
DE(242) = "PRECEDE"
DE(243) = "PRESENT"
DE(244) = "PRETEND"
DE(245) = "PRIMATE"
DE(246) = "PROBLEM"
DE(247) = "PROCEED"
DE(248) = "PROBATE"
DE(249) = "PROFILE"
DE(250) = "PROLONG"
DE(251) = "PROPHET"
DE(252) = "QUARTET"
DE(253) = "QUARTER"
DE(254) = "CRYSTAL"
DE(255) = "QUINTET"
DE(256) = "SHACKLE"
DE(257) = "RAPTURE"
DE(258) = "REACTOR"
DE(259) = "REALISM"
DE(260) = "RECEIVE"
DE(261) = "RECEIPT"
DE(262) = "CLASSIC"
DE(263) = "RECITAL"
DE(264) = "RECRUIT"
DE(265) = "REFRACT"
DE(266) = "REGNANT"
DE(267) = "RELAPSE"
DE(268) = "REGULAR"
DE(269) = "RADDISH"
DE(270) = "REMAINS"
DE(271) = "REMNANT"
DE(272) = "REPLICA"
DE(273) = "REPLETE"
DE(274) = "REPULSE"
DE(275) = "RESTIVE"
DE(276) = "RESTORE"
DE(277) = "RETRACT"
DE(278) = "RINGLET"
DE(280) = "ROLLBAR"
DE(281) = "ROMANCE"
DE(282) = "RONDURE"
DE(283) = "ROSETTE"
DE(284) = "PREMIUM"
DE(285) = "SHINGLE"
DE(286) = "SHINKIN"
DE(287) = "SUBJECT"
DE(288) = "SHUTTLE"
DE(289) = "SIMILAR"
DE(290) = "SKITTLE"
DE(291) = "SLUMBER"
DE(292) = "SOCIETY"
DE(293) = "SOLANUM"
DE(294) = "SYRINGE"
DE(295) = "TADPOLE"
DE(296) = "TANKARD"
DE(297) = "TANTRUM"
DE(298) = "UNICORN"
DE(299) = "NOURISH"
DE(300) = "VALUATE"
DE(301) = "VAGRANT"
CI(1) = "No way Jose!"
CI(2) = "NA  NA  N'NA  NA  NA"
CI(3) = "Nice try!"
CI(4) = "Good try!"
CI(5) = "Never mind!"
CI(6) = "You blew it!"
CI(7) = "No way siree"
CI(8) = "Yer plonker!"
CI(9) = "We all make misteaks"
CI(10) = "0BO 110X !"
CI(11) = "Try again!"
CI(12) = "Unlucky - Try again!"
CI(13) = "Micro brain!"
CI(14) = "No way!"
CI(15) = "Not a prayer!"
CI(16) = "Not a chance!"
CI(17) = "Not a hope!"
CI(18) = "Naff effort!"
CI(19) = "Close but not quite!"
CI(20) = "Had you there - Tee Hee Hee!"
CI(21) = "It's a toughie - isn't it?"
CI(22) = "@h!*t!"
CI(23) = "It's no use just guessing!"
CI(24) = "Fluffed up there!!"
CI(25) = "Hopeless!"
CI(26) = "HAR HAR HAR!"
CI(27) = "Load 'a cobblers!"
CI(28) = "Keep going...Keep going!"
CI(29) = "We might be here all night!"
CI(30) = "Knickers!"
CI(31) = "OOPS!!"
CI(32) = "Not on your life!"
CI(33) = "Guffaw...Guffaw!"
CI(34) = "Er, not quite!"
CI(35) = "Geek!"
CI(36) = "You must be kidding!"
CI(37) = "You must be joking!"
CJ(1) = "Correctamundo - "
CJ(2) = "Congrats' - "
CJ(3) = "Congratulations - "
CJ(4) = "Your lucky day - "
CJ(5) = "Well done - "
CJ(6) = "Well played - "
CJ(7) = "Eureka! - "
CJ(8) = "Your'e the winner, "
CJ(9) = "Hurray - "
CJ(10) = "Brilliant job - "
CJ(11) = "You did it! - "
CJ(12) = "You made it! - "
CJ(13) = "Hip Hip Hurray! - "
CJ(14) = "Brillo! - "
CK(1) = "Correctamundo"
CK(2) = "Congrats'"
CK(3) = "Congratulations"
CK(4) = "Your lucky day"
CK(5) = "Well done"
CK(6) = "Well played"
CK(7) = "Eureka!"
CK(8) = "Your'e the winner"
CK(9) = "Hurray"
CK(10) = "Brilliant job"
CK(11) = "You did it!"
CK(12) = "You made it!"
CK(13) = "Hip Hip Hurray!"
CK(14) = "Brillo!"
End Function

Public Function GetUserWords()
Dim ABC As String
Dim ABD
Dim AA, AD As Integer
Dim AB, AC As String
UW = 1
mnuUserWord.Checked = True
mnuComputerWord.Checked = False
10 If R = 1 Then
ABC = InputBox(Player1 + "  Enter your word.  -  (Max. 5 characters)   ", Player2 + "  -   PLEASE LOOK AWAY")
End If
If R = 2 Then
ABC = InputBox(Player1 + "  Enter your word.  -  (Max. 6 characters)   ", Player2 + "  -   PLEASE LOOK AWAY")
End If
If R = 3 Then
ABC = InputBox(Player1 + "  Enter your word.  -  (Max. 7 characters)   ", Player2 + "  -   PLEASE LOOK AWAY")
End If
If ABC = "" Then
AA = vbYesNo + vbQuestion
AB = "FLIP!"
AC = "Use computer generated words?"
AD = MsgBox(AC, AA, AB)
If AD = vbNo Then GoTo 10
If AD = vbYes Then GoTo 15
End If
11 If R = 1 Then
ABD = InputBox(Player2 + "  Enter your word.  -  (Max. 5 characters)   ", Player1 + "  -   PLEASE LOOK AWAY")
End If
If R = 2 Then
ABD = InputBox(Player2 + "  Enter your word.  -  (Max. 6 characters)   ", Player1 + "  -   PLEASE LOOK AWAY")
End If
If R = 3 Then
ABD = InputBox(Player2 + "  Enter your word.  -  (Max. 7 characters)   ", Player1 + "  -   PLEASE LOOK AWAY")
End If
If ABD = "" Then
AD = MsgBox(AC, AA, AB)
If AD = vbNo Then GoTo 11
If AD = vbYes Then GoTo 15
End If
Word1 = UCase(ABC)
Word2 = UCase(ABD)
GoTo 20
15 UW = 0
mnuUserWord.Checked = False
mnuComputerWord.Checked = True
20
End Function
Public Function Setup1Player()
If PX = 0 Then GoTo 11
PX = 0
UW = 0
Pic1player.Picture = ImageList2.ListImages(22).Picture
pic2player.Picture = ImageList2.ListImages(23).Picture
mnu2Player.Checked = False
mnu1Player.Checked = True
mnuUserWord.Enabled = False
mnuUserWord.Checked = False
mnuComputerWord.Checked = True
lblplayer1.Visible = False
lblplayer2.Visible = False
Shpplay1.Visible = False
shpplay2.Visible = False
lblTries.Visible = True
ChangeLevels
11
End Function

Public Function GetPlayersNames()
Dim AA, AD As Integer
Dim AB, AC As String
Player1 = InputBox("Please enter your name;", "PLAYER 1")
If Player1 = "" Then
Setup1Player
GoTo 100
End If
Player1 = StrConv(Player1, vbProperCase)
lblplayer1.Caption = Player1 + "  " + Str(PL1) + "     " + Str(GA1)
Player2 = InputBox("Please enter your name;", "PLAYER 2")
If Player2 = "" Then
PX = 0
Setup1Player
GoTo 100
End If
Player2 = StrConv(Player2, vbProperCase)
lblplayer2.Caption = Player2 + "  " + Str(PL2) + "     " + Str(GA2)
AA = vbYesNo + vbQuestion
AB = "FLIP!"
AC = "Use computer generated words?"
AD = MsgBox(AC, AA, AB)
If AD = vbYes Then GoTo 100
UW = 1
100
End Function
Public Function Switch()
SW = SW + 1
If SW = 2 Then
SW = 0
End If
If SW = 1 Then
shpplay2.Visible = True
Shpplay1.Visible = False
txtWord = Word1
Text3.Text = Word1
End If
If SW = 0 Then
shpplay2.Visible = False
Shpplay1.Visible = True
txtWord = Word2
Text3.Text = Word2
End If
End Function

Public Function Upgrade()
mnuValidate.Enabled = False
Flip.Caption = "Flip! for Windows V2.1"
Surround.FillColor = RGB(0, 255, 0)
lbl1.BackColor = RGB(0, 255, 0)
lbl2.BackColor = RGB(0, 255, 0)
lbl3.BackColor = RGB(0, 255, 0)
lbl4.BackColor = RGB(0, 255, 0)
lbl5.BackColor = RGB(0, 255, 0)
lbl6.BackColor = RGB(0, 255, 0)
lbl7.BackColor = RGB(0, 255, 0)
lbl8.BackColor = RGB(0, 255, 0)
lbl9.BackColor = RGB(0, 255, 0)
lbl10.BackColor = RGB(0, 255, 0)
lbl11.BackColor = RGB(0, 255, 0)
lbl12.BackColor = RGB(0, 255, 0)
lblA.BackColor = RGB(0, 255, 0)
lblB.BackColor = RGB(0, 255, 0)
lblC.BackColor = RGB(0, 255, 0)
lblD.BackColor = RGB(0, 255, 0)
lblE.BackColor = RGB(0, 255, 0)
lblF.BackColor = RGB(0, 255, 0)
Surround.BorderColor = RGB(0, 255, 0)
Shpplay1.BorderColor = RGB(255, 0, 0)
shpplay2.BorderColor = RGB(255, 0, 0)
End Function

Public Function Setup2Players()
If PX = 1 Then GoTo 11
PX = 1
ST = 1
LT = 0
GA1 = 0
GA2 = 0
CT = 0
If OPT = 1 Then
AX = 0
W = 1
End If
Pic1player.Picture = ImageList2.ListImages(21).Picture
pic2player.Picture = ImageList2.ListImages(24).Picture
GetPlayersNames
mnu2Player.Checked = True
mnu1Player.Checked = False
Mnucontinuous.Checked = False
mnuShortTunes.Checked = True
mnuNoTunes.Checked = False
DL = 1820
DM = 2325
mnuUserWord.Enabled = True
lblplayer1.Visible = True
lblplayer2.Visible = True
Shpplay1.Visible = True
shpplay2.Visible = False
lblTries.Visible = False
ChangeLevels
Timer5.Enabled = False
11
End Function

Private Sub Timer5_Timer()
Timer5.Interval = 1000
lblNana.Caption = "Time - " + Str(SE) + "  Seconds,   " + Str(CX) + " Cleared"
SE = SE + 1
If SE >= 32000 Then SE = 32000
If CX >= (Z - 1) Then
Timer5.Enabled = False
PlayCorrect
CheckScores
End If
End Sub



Public Function TimedGame()
ST = 0: CT = 1: LT = 0: SA = 0: AX = 1
SE = 0
PX = 0
UW = 0
lblTries.Visible = False
With lblNana
.Visible = True
.Caption = "Clear the Board:"
.ForeColor = RGB(255, 255, 0)
End With
DL = 500
DM = 1000
Timer5.Enabled = True
With frmOpen.MMControl1
.Shareable = False
.DeviceType = "Sequencer"
.Command = "close"
.filename = "37.MID"
.Command = "Open"
.Command = "play"
End With
End Function

Public Function Animated()
For DW = 1 To Z
If Q(DW) >= 1 Then
G(DW).Picture = ImageList1.ListImages(S(DW)).Picture
S(DW) = S(DW) + 1
If S(DW) > E(DW) Then
S(DW) = D(DW)
End If
End If
Next
End Function
