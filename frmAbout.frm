VERSION 4.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About....."
   ClientHeight    =   2565
   ClientLeft      =   1515
   ClientTop       =   1620
   ClientWidth     =   6690
   Height          =   2970
   Left            =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6690
   Top             =   1275
   Width           =   6810
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
frmAbout.Hide
End Sub


