Attribute VB_Name = "Module1"
Public FA As WordInfo
Type WordInfo '  <---- This is for writing the score table to the Hard drive
    Name As String * 7
    Begin1 As String * 12
    Inter1 As String * 12
    Expert1 As String * 12
    BScore1 As Integer
    BScore2 As Integer
    BScore3 As Integer
    Begin2 As String * 12
    Inter2 As String * 12
    Expert2 As String * 12
    WScore1 As Integer
    WScore2 As Integer
    WScore3 As Integer
    BTime1 As Integer
    BTime2 As Integer
    BTime3 As Integer
    TName1 As String * 12
    Tname2 As String * 12
    Tname3 As String * 12
    Validate As String * 8
End Type

Public GWD, LA, LB, LC, LD, LH, LI, LJ, LN, LO, LP, LQ As String
Public FB, LE, LF, LG, LK, LL, LM, JJ, JK, JL, JM, JN, JU As Integer
Public Fd, FC As Long
Public SWP(1 To 2) As Integer

