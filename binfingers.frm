VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   Caption         =   "Digital Finger Counting"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleMode       =   0  'User
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin Project1.DigitalFingers ctlDigitalFingers1 
      Height          =   4515
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   7964
      Palm            =   0   'False
      PDisp           =   3
   End
   Begin VB.Frame fraScore 
      Caption         =   "Score"
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   3735
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   725
         Index           =   2
         Left            =   100
         ScaleHeight     =   720
         ScaleWidth      =   3540
         TabIndex        =   16
         Top             =   175
         Width           =   3535
         Begin VB.CommandButton cmdNew 
            Caption         =   "Clear"
            Height          =   495
            Left            =   20
            TabIndex        =   17
            Top             =   40
            Width           =   615
         End
         Begin VB.Label lblResult 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2180
            TabIndex        =   19
            Top             =   40
            Width           =   1335
         End
         Begin VB.Label lblScore 
            Caption         =   "Label2"
            Height          =   615
            Left            =   860
            TabIndex        =   18
            Top             =   40
            Width           =   1215
         End
      End
   End
   Begin VB.TextBox txtHelp 
      Height          =   2295
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "binfingers.frx":0000
      Top             =   4920
      Width           =   6135
   End
   Begin VB.Frame fraMakeItOdd 
      Caption         =   "Make it Odd"
      Height          =   1095
      Left            =   2040
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   845
         Index           =   1
         Left            =   100
         ScaleHeight     =   840
         ScaleWidth      =   1620
         TabIndex        =   8
         Top             =   175
         Width           =   1620
         Begin VB.CheckBox chkComplex 
            Alignment       =   1  'Right Justify
            Caption         =   "Complex"
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdOddEven 
            Caption         =   "5"
            Height          =   375
            Index           =   5
            Left            =   1295
            TabIndex        =   14
            Top             =   160
            Width           =   255
         End
         Begin VB.CommandButton cmdOddEven 
            Caption         =   "4"
            Height          =   375
            Index           =   4
            Left            =   1040
            TabIndex        =   13
            Top             =   160
            Width           =   255
         End
         Begin VB.CommandButton cmdOddEven 
            Caption         =   "3"
            Height          =   375
            Index           =   3
            Left            =   785
            TabIndex        =   12
            Top             =   160
            Width           =   255
         End
         Begin VB.CommandButton cmdOddEven 
            Caption         =   "2"
            Height          =   375
            Index           =   2
            Left            =   530
            TabIndex        =   11
            Top             =   160
            Width           =   255
         End
         Begin VB.CommandButton cmdOddEven 
            Caption         =   "1"
            Height          =   375
            Index           =   1
            Left            =   275
            TabIndex        =   10
            Top             =   160
            Width           =   255
         End
         Begin VB.CommandButton cmdOddEven 
            Caption         =   "0"
            Height          =   375
            Index           =   0
            Left            =   20
            TabIndex        =   9
            Top             =   160
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraRockPaperScissors 
      Caption         =   "Rock, Paper, Scissors"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   0
         Left            =   100
         ScaleHeight     =   840
         ScaleWidth      =   1620
         TabIndex        =   2
         Top             =   175
         Width           =   1620
         Begin VB.CommandButton cmdRPS 
            Caption         =   "Rock"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   5
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdRPS 
            Caption         =   "Paper"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdRPS 
            Caption         =   "Scissors"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   3
            Top             =   600
            Width           =   735
         End
      End
   End
   Begin VB.ListBox lstNumbers 
      Height          =   4740
      Left            =   9120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin Project1.DigitalFingers ctlDigitalFingers1 
      Height          =   4515
      Index           =   1
      Left            =   4500
      TabIndex        =   21
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   7964
      PDisp           =   2
   End
   Begin VB.Label lblKeyboardFinger 
      Caption         =   $"binfingers.frx":0006
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   4560
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strLongBin  As String
Private RPSUser     As Long
Private RPSComp     As Long
Private RPSDraw     As Long
Private Playing     As Boolean

Private Function Base2To10(strBase2 As String) As Long

  Dim I As Long

  For I = 1 To 10
    Base2To10 = Base2To10 + IIf(Mid$(strBase2, I, 1) = "1", 2 ^ (10 - I), 0)
  Next I

End Function

Private Function binaryString(ByVal X As Long) As String

'MODIFIED FROM VB HELP FILE called
''How to Convert a Decimal Number to a Binary Number in a String'
'Check it out if you need bigger numbers

  Dim I          As Long

  For I = 9 To 0 Step -1
    binaryString = binaryString & IIf(X And (2 ^ I), "1", "0")
  Next I
' The binaryString string contains the binaryString number.
  binaryString = Right$(binaryString, 10)
  

End Function

Private Sub cmdNew_Click()

  lblResult.Caption = ""
  RPSUser = 0
  RPSComp = 0
  RPSDraw = 0
  ScoreBoard

End Sub

Private Sub cmdOddEven_Click(Index As Integer)

  Dim MyNum     As Long
  Dim UserCount As Long
  Dim CompNum   As Long
  Dim CompCount As Long

  CompCount = Int(Rnd * 6) + 1
  UserCount = Index + 1
  MyNum = Choose(UserCount, 0, 512, 768, 896, 960, 992)
  CompNum = Choose(CompCount, 0, 1, 3, 7, 15, 31)
' the results produced by this are
'        0     1     2     3     4     5 (CompCount)
'   0    D     W     L     W     L     D
'   1    W     D     W     L     D     L
'   2    L     W     D     D     L     W
'   3    W     L     D     D     W     L
'   4    L     D     L     W     D     W
'   5    D     L     W     L     W     D
'(UserCount)
'     Simple       Complex
'Win   18            12
'Lose  18            12
'Draw  --            12
  strLongBin = binaryString(MyNum + CompNum)
'     Equal values                         Total = 5 (but XXXCount are 1-6 so deduct 1 from each)
  If chkComplex.Value = vbChecked And ((CompCount = UserCount) Or (CompCount - 1 + UserCount - 1 = 5)) Then
    RPSDraw = RPSDraw + 1
    lblResult.Caption = "DRAW"
   ElseIf (CompCount + UserCount) Mod 2 = 0 Then
    RPSComp = RPSComp + 1
    lblResult.Caption = "LOSE"
   Else
    RPSUser = RPSUser + 1
    lblResult.Caption = "WIN"
  End If
  ScoreBoard

End Sub

Private Sub cmdOddEven_MouseDown(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

  GameFist

End Sub

Private Sub cmdRPS_Click(Index As Integer)

  Dim MyNum   As Long
  Dim CompNum As Long

  MyNum = Choose(Index + 1, 0, 960, 192)
'if your paranoid that the machine is cheating get the comp number first ;)
  CompNum = Choose(Int(Rnd * 3) + 1, 0, 12, 15)
' the numbers produced by this are
'         R      P       S (Comp)
'   R    D 0    L 15    W 12
'   P    W 960  D 975   L 972
'   S    L 192  W 207   D 204
'(User)
  strLongBin = binaryString(MyNum + CompNum)
'hand
  Select Case MyNum + CompNum
   Case 12, 960, 207
    lblResult.Caption = "WIN"
    RPSUser = RPSUser + 1
   Case 0, 975, 204
    lblResult.Caption = "DRAW"
    RPSDraw = RPSDraw + 1
   Case 15, 972, 192
    lblResult.Caption = "LOSE"
    RPSComp = RPSComp + 1
  End Select
  ScoreBoard

End Sub

Private Sub cmdRPS_MouseDown(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

  GameFist

End Sub

Private Sub ctlDigitalFingers1_Change(Index As Integer)

'only test the second control;
'setting the list from 1st control would always reset value of the second control to 0

  If Index = 1 Or Playing = False Then
    lstNumbers.ListIndex = Base2To10(ctlDigitalFingers1(0).BinaryValue & ctlDigitalFingers1(1).BinaryValue)
  End If

End Sub

Private Sub ctlDigitalFingers1_GotFocus(Index As Integer)

  GameOff

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

' not very useful but fun
'allows you to use keyboard to change fingers
'feel free to change the keys (this set is just comfortable for my hands & keyboard)
'uses a simple toggle effect to determine whether to raise/lower the finger
'Let me know if you think of a way to integrate this with the Control
  Select Case UCase$(Chr$(KeyAscii))
   Case "S"
    Mid$(strLongBin, 1, 1) = IIf(Mid$(strLongBin, 1, 1) = "0", "1", "0")
    GameOff
   Case "E"
    Mid$(strLongBin, 2, 1) = IIf(Mid$(strLongBin, 2, 1) = "0", "1", "0")
    GameOff
   Case "R"
    Mid$(strLongBin, 3, 1) = IIf(Mid$(strLongBin, 3, 1) = "0", "1", "0")
    GameOff
   Case "T"
    Mid$(strLongBin, 4, 1) = IIf(Mid$(strLongBin, 4, 1) = "0", "1", "0")
    GameOff
   Case "B"
    Mid$(strLongBin, 5, 1) = IIf(Mid$(strLongBin, 5, 1) = "0", "1", "0")
    GameOff
   Case "N"
    Mid$(strLongBin, 6, 1) = IIf(Mid$(strLongBin, 6, 1) = "0", "1", "0")
    GameOff
   Case "I"
    Mid$(strLongBin, 7, 1) = IIf(Mid$(strLongBin, 7, 1) = "0", "1", "0")
    GameOff
   Case "O"
    Mid$(strLongBin, 8, 1) = IIf(Mid$(strLongBin, 8, 1) = "0", "1", "0")
    GameOff
   Case "P"
    Mid$(strLongBin, 9, 1) = IIf(Mid$(strLongBin, 9, 1) = "0", "1", "0")
    GameOff
   Case "'"
    Mid$(strLongBin, 10, 1) = IIf(Mid$(strLongBin, 10, 1) = "0", "1", "0")
    GameOff
  End Select
  hand

End Sub

Private Sub Form_Load()

  Dim I As Long

  For I = 0 To 1023
    lstNumbers.AddItem I
  Next I
  Randomize Timer
  strLongBin = "0000000000"
  hand
  ScoreBoard
  lstNumbers.ListIndex = Int(Rnd * 1023)
  txtHelp.Text = "Digital Fingers" & vbNewLine & "This is a simple tool to teach how to count in binary on your fingers." & vbNewLine & _
   "A finger extended is 1; folded is 0. For the full range the program assumes you are holding your hands palm down." & vbNewLine & _
   "It is unlikely you will ever want to go beyond 31 ( one hand). If you are Left-handed (like me) just hold your hand palm up to copy the right hand image in the program (the original picture was a left hand palm up (31))." & vbNewLine & _
   "You can select a list value, click on fingers or use the keyboard (see keys listed under the images) to set the list value." & vbNewLine & _
   "NOTE the keyboard input is not integrated into the UserControl and can easily be changed for smaller/larger hands; see code)" & vbNewLine & _
   "Games" & vbNewLine & "I have included a couple of games which exploit the UserControl to let you play. If your culture has similar hand games you might look at these to develop them." & vbNewLine & _
   "General Rule for games: Click on a button to fold hands into fists, release the button and the fingers will display." & vbNewLine & _
   "Left image is user; Right image is machine." & vbNewLine & "(NOTE Right image inverts; using any of the number input methods returns it to upright)" & vbNewLine & _
   "Rock, Paper, Scissors." & vbNewLine & "DRAW: hands match. " & vbNewLine & "WIN: Rock breaks Scissors, Scissors cut Paper and Paper covers Rock. " & vbNewLine & _
   "Make it Odd" & vbNewLine & "Select a number of fingers (0-5)." & vbNewLine & "WIN: Your number plus the computer's is odd you win." & vbNewLine & _
   "Complex version includes potential for Draw result." & vbNewLine & "DRAW: Total is 5 fingers OR you both produce same number of fingers)." & vbNewLine & _
   "Apologies if any of the hand forms are rude in your culture. (see 4, 128 and for emphasis 132 ;)"




End Sub

Private Sub GameFist()

'Call from MouseDown Event of game buttons
'Clears the game result label and forms fists

  GameOn
  lblResult.Caption = ""
  lstNumbers.ListIndex = 0
  Playing = True

End Sub

Private Sub GameOff()

  ctlDigitalFingers1(1).Orientation = orUp
  Me.Caption = "Digital Finger Counting "

End Sub

Private Sub GameOn()
Me.Caption = "Digital Finger Counting  GAME MODE"
  ctlDigitalFingers1(1).Orientation = orDown

End Sub

Private Sub hand()

  ctlDigitalFingers1(0).BinaryValue = Left$(strLongBin, 5)
  ctlDigitalFingers1(1).BinaryValue = Right$(strLongBin, 5)

End Sub

Private Sub lstNumbers_Click()

  strLongBin = binaryString(lstNumbers.ListIndex)
  hand
  Playing = True

End Sub

Private Sub lstNumbers_GotFocus()

  GameOff
':)Code Fixer V2.0.0 (12/04/2004 1:05:47 PM) 6 + 221 = 227 Lines Thanks Ulli for inspiration and lots of code.

End Sub

Private Sub ScoreBoard()

  hand
  lblScore.Caption = "  User: " & RPSUser & vbNewLine & "Comp: " & RPSComp & vbNewLine & "Draw: " & RPSDraw

End Sub


':)Code Fixer V2.0.0 (14/04/2004 1:09:46 PM) 6 + 255 = 261 Lines Thanks Ulli for inspiration and lots of code.

