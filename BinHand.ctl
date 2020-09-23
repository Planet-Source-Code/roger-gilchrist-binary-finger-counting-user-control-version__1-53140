VERSION 5.00
Begin VB.UserControl DigitalFingers 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   MaskColor       =   &H00000000&
   ScaleHeight     =   5685
   ScaleWidth      =   5970
   Begin VB.PictureBox picFingerSource 
      BorderStyle     =   0  'None
      Height          =   1545
      Index           =   1
      Left            =   5520
      Picture         =   "BinHand.ctx":0000
      ScaleHeight     =   1545
      ScaleWidth      =   660
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picFingerSource 
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   2
      Left            =   4800
      Picture         =   "BinHand.ctx":0982
      ScaleHeight     =   1980
      ScaleWidth      =   525
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picFingerSource 
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   3
      Left            =   5640
      Picture         =   "BinHand.ctx":138F
      ScaleHeight     =   2130
      ScaleWidth      =   510
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picFingerSource 
      BorderStyle     =   0  'None
      Height          =   2010
      Index           =   4
      Left            =   4800
      Picture         =   "BinHand.ctx":1EB5
      ScaleHeight     =   2010
      ScaleWidth      =   705
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picThumbSource 
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   1
      Left            =   120
      Picture         =   "BinHand.ctx":29AF
      ScaleHeight     =   4500
      ScaleWidth      =   4500
      TabIndex        =   6
      Top             =   4800
      Width           =   4500
   End
   Begin VB.PictureBox picThumbSource 
      BorderStyle     =   0  'None
      Height          =   4515
      Index           =   0
      Left            =   4800
      Picture         =   "BinHand.ctx":5256
      ScaleHeight     =   4515
      ScaleWidth      =   4500
      TabIndex        =   5
      Top             =   4680
      Width           =   4500
   End
   Begin VB.PictureBox picThumb 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4560
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.PictureBox picFingers 
         AutoRedraw      =   -1  'True
         Height          =   1545
         Index           =   1
         Left            =   3240
         ScaleHeight     =   1485
         ScaleWidth      =   600
         TabIndex        =   4
         Top             =   1000
         Width           =   660
      End
      Begin VB.PictureBox picFingers 
         AutoRedraw      =   -1  'True
         Height          =   1980
         Index           =   2
         Left            =   2760
         ScaleHeight     =   1920
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   275
         Width           =   525
      End
      Begin VB.PictureBox picFingers 
         AutoRedraw      =   -1  'True
         Height          =   2130
         Index           =   3
         Left            =   2175
         ScaleHeight     =   2070
         ScaleWidth      =   450
         TabIndex        =   2
         Top             =   25
         Width           =   510
      End
      Begin VB.PictureBox picFingers 
         AutoRedraw      =   -1  'True
         Height          =   2010
         Index           =   4
         Left            =   1320
         ScaleHeight     =   1950
         ScaleWidth      =   645
         TabIndex        =   1
         Top             =   83
         Width           =   705
      End
      Begin VB.Label lblDisplayValues 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11111"
         ForeColor       =   &H0080FF80&
         Height          =   195
         Left            =   3735
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
   End
End
Attribute VB_Name = "DigitalFingers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Just a simple UserControl to take care of the images needed for the graphics
'NOTE because the images may be blanked and/or flipped they are stored seperately from the
'displaying controls
'The UC uses the BackColor Property to detect whether a finger is extended (1) or not (0).
'The eye can't detect the difference between 0 and 1.
'You could contruct an internal logic to do the same thing
'but it was an idea I had while roughing it out and I liked it.
'for simplicity sake each DigitalFingers control only recognizes values between 0 and 31
'the demo form does its own math to convert them to 10 digit values
'Thumb hot zone is the quarter of the image containing the thumb
'Finger hot zones are the bounding rectangle of the pictureboxes
'to save fiddling I have made the control maintain a single size
'If you want to allow resizing you will also have to deal with relocating the fingers in the mirror option
'PROPERTIES
'BinaryValue:     Accepts a string 5 characters made of '1's and '0's.
'                 Longer strings are left cut. Shorter strings are right padded with '0'.
'                 Unaceptable characters are set to '0'
'     NOTE Control keeps the binary and decimal inputs in sync, so only the BinaryValue is stored in the PropBag.
'DecimalValue:    Allows you to set fingers with decimal numbers.
'                 Numbers outside the 0-31 range are reduced to 0 if < 0 and 31 if > 31.
'                 Letters are rejected.
'DisplayColour:   Set the colour of the numeric readout
'DisplayPosition: Set the position (corners of control) for numeric readout
'DisplayShow:     Turn numeric readout on/off
'DisplayStyle:    Set numeric readout style Decimal/Binary
'LeftPalm :       Determine which hand is shown (Default True; thumb to left)
'Orientation :    I have started to develop this but at present only fingers pointing Up/Down are working
Public Enum Orient
  orUp
  orDown
  orRight
  orLeft
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private orUp, orDown, orRight, orLeft
#End If
Public Enum DisplayMode
  dmBinary
  dmDecimal
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private dmBinary, dmDecimal
#End If
Public Enum DisplayPos
  TopLeft
  TopRight
  BottomLeft
  BottomRight
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TopLeft, TopRight, BottomLeft, BottomRight
#End If
Private strBin            As String ' internal store for finger value
Private DMode             As DisplayMode
Private DPos              As DisplayPos
Private dcolor            As Long
Private bDisplayShow      As Boolean
Private bleftPalm         As Boolean
Private m_Orientation     As Orient
Private RealWidth         As Long    ' used to prevent resizing
Private RealHeight        As Long
Private Type Desc
  dLeft                   As Long
  dTop                    As Long
  dWidth                  As Long
  dHeight                 As Long
End Type
Private fingPos(4)        As Desc
'stores the initial position of the controls on the form used for orientating controls
Public Event Change()

Private Function Base10to2(ByVal X As Long, _
                           Optional Range As Long = 5) As String

'MODIFIED FROM VB HELP FILE called
''How to Convert a Decimal Number to a Binary Number in a String'
'Check it out if you need bigger numbers
'Range is number of binary digits to deal with

  Dim I          As Long

  For I = Range - 1 To 0 Step -1
    Base10to2 = Base10to2 & IIf(X And (2 ^ I), "1", "0")
  Next I
' The binaryString string contains the binaryString number.

End Function

Private Function Base2To10(strBase2 As String, _
                           Optional Range As Long = 5) As Long

'logical reversal of above
'Range is number of binary digits to deal with

  Dim I As Long

  For I = 1 To Range
    Base2To10 = Base2To10 + IIf(Mid$(strBase2, I, 1) = "1", 2 ^ (Range - I), 0)
  Next I

End Function

Public Property Get BinaryValue() As String

  BinaryValue = strBin

End Property

Public Property Let BinaryValue(strInput As String)

  Dim I As Long

  If Len(strInput) > 5 Then
    strInput = Left$(strInput, 5)
   ElseIf Len(strInput) < 5 Then
    Do While Len(strInput) < 5
      strInput = "0" & strInput
    Loop
  End If
  For I = 1 To Len(strInput)
    If Mid$(strInput, I, 1) <> "1" Then
      If Mid$(strInput, I, 1) <> "0" Then
        Mid$(strInput, I, 1) = "0"
      End If
    End If
  Next I
  strBin = strInput
  DrawHand

End Property

Public Property Get DecimalValue() As Long

  DecimalValue = Base2To10(strBin)

End Property

Public Property Let DecimalValue(lngVal As Long)

  If lngVal < 0 Then
    lngVal = 0
   ElseIf lngVal > 21 Then
    lngVal = 31
  End If
  BinaryValue = Base10to2(lngVal)
  DrawHand

End Property

Public Property Get DisplayColor() As OLE_COLOR

  DisplayColor = dcolor

End Property

Public Property Let DisplayColor(ByVal DCol As OLE_COLOR)

' using OLE_COLOR rather than Long allows the VB Property Window to display the full colour selector

  dcolor = DCol
  DrawHand

End Property

Public Property Get DisplayPosition() As DisplayPos

  DisplayPosition = DPos

End Property

Public Property Let DisplayPosition(DMod As DisplayPos)

'Using Enums allows the VB Property Window to display more informative lists of possible values

  DPos = DMod
  DrawHand

End Property

Public Property Get DisplayShow() As Boolean

  DisplayShow = bDisplayShow

End Property

Public Property Let DisplayShow(ByVal bDisp As Boolean)

  bDisplayShow = bDisp
  lblDisplayValues.Visible = bDisplayShow

End Property

Public Property Get DisplayStyle() As DisplayMode

  DisplayStyle = DMode

End Property

Public Property Let DisplayStyle(DMod As DisplayMode)

'Using Enums allows the VB Property Window to display more informative lists of possible values

  DMode = DMod
  DrawHand

End Property

Private Sub DoReadOut()

  If bDisplayShow Then
    lblDisplayValues.ForeColor = dcolor
    lblDisplayValues.Caption = IIf(DMode = dmBinary, strBin, Base2To10(strBin))
    Select Case DPos
     Case TopLeft
      lblDisplayValues.Move 0, 0
     Case TopRight
      lblDisplayValues.Move UserControl.Width - lblDisplayValues.Width, 0
     Case BottomLeft
      lblDisplayValues.Move 0, UserControl.Height - lblDisplayValues.Height
     Case BottomRight
      lblDisplayValues.Move UserControl.Width - lblDisplayValues.Width, UserControl.Height - lblDisplayValues.Height
    End Select
  End If

End Sub

Private Sub DrawHand()

  FingerShow IIf(bleftPalm, 1, 1), Mid$(strBin, IIf(bleftPalm, 5, 1), 1) = "1"
  FingerShow IIf(bleftPalm, 2, 2), Mid$(strBin, IIf(bleftPalm, 4, 2), 1) = "1"
  FingerShow IIf(bleftPalm, 3, 3), Mid$(strBin, IIf(bleftPalm, 3, 3), 1) = "1"
  FingerShow IIf(bleftPalm, 4, 4), Mid$(strBin, IIf(bleftPalm, 2, 4), 1) = "1"
  Flip picThumb, picThumbSource(IIf(Mid$(strBin, IIf(bleftPalm, 1, 5), 1) = "1", 1, 0)), IIf(Mid$(strBin, IIf(bleftPalm, 1, 5), 1) = "1", 1, 0)
  DoReadOut
'feedback (no effect if set from list but finger clicking updates list)
  RaiseEvent Change

End Sub

Private Sub FingerLocation()

  Dim I As Long

  Select Case Me.Orientation
   Case orUp
    If bleftPalm Then
      For I = 1 To 4
        With picFingers(I)
          .Left = fingPos(I).dLeft
          .Top = fingPos(I).dTop
          .Width = fingPos(I).dWidth
          .Height = fingPos(I).dHeight
        End With 'picFingers(I)
      Next I
     Else
      picFingers(4).Left = fingPos(1).dLeft - 800
      picFingers(3).Left = fingPos(2).dLeft - 950
      picFingers(2).Left = fingPos(3).dLeft - 950
      picFingers(1).Left = fingPos(4).dLeft - 700
    End If
   Case orDown
    picFingers(4).Top = fingPos(4).dTop + 2300
    picFingers(3).Top = fingPos(3).dTop + 2350
    picFingers(2).Top = fingPos(2).dTop + 1900
    picFingers(1).Top = fingPos(1).dTop + 950
    If bleftPalm Then
      For I = 1 To 4
        picFingers(I).Left = fingPos(I).dLeft
      Next I
     Else
      picFingers(4).Left = fingPos(1).dLeft - 800
      picFingers(3).Left = fingPos(2).dLeft - 950
      picFingers(2).Left = fingPos(3).dLeft - 900
      picFingers(1).Left = fingPos(4).dLeft - 700
      picFingers(4).Top = fingPos(4).dTop + 2300
      picFingers(3).Top = fingPos(3).dTop + 2350
      picFingers(2).Top = fingPos(2).dTop + 1900
      picFingers(1).Top = fingPos(1).dTop + 950
    End If
   Case orLeft
    MsgBox "Not available yet"
    Orientation = orUp
   Case orRight
    MsgBox "Not available yet"
    Orientation = orUp
  End Select
  DrawHand

End Sub

Private Sub FingerShow(ByVal FNumber As Long, _
                       ByVal bUp As Boolean)

  With picFingers(FNumber)
    If bUp Then
      Flip picFingers(FNumber), picFingerSource(FNumber), 1
     Else
      .Picture = LoadPicture("")
      .BackColor = vbBlack
    End If
  End With

End Sub

Private Sub Flip(dest As PictureBox, _
                 src As PictureBox, _
                 ByVal BCol As Long)

  dest.BackColor = BCol
  dest.Picture = LoadPicture("") 'stops a nasty flicker when changing thumb positions
  Select Case m_Orientation
   Case orUp
    If bleftPalm Then
      dest.Picture = src.Picture
     Else
      dest.PaintPicture src.Picture, 0, 0, src.ScaleWidth, src.ScaleHeight, src.ScaleWidth, 0, -src.ScaleWidth, src.ScaleHeight, &HCC0020
    End If
   Case orDown
    If bleftPalm Then
      dest.PaintPicture src.Picture, 0, 0, src.ScaleWidth, src.ScaleHeight, 0, src.ScaleHeight, src.ScaleWidth, -src.ScaleHeight, &HCC0020
     Else
      dest.PaintPicture src.Picture, 0, 0, src.ScaleWidth, src.ScaleHeight, src.ScaleWidth, src.ScaleHeight, -src.ScaleWidth, -src.ScaleHeight, &HCC0020
    End If
   Case orLeft
   Case orRight
  End Select

End Sub

Public Property Get LeftPalm() As Boolean

  LeftPalm = bleftPalm

End Property

Public Property Let LeftPalm(ByVal TFPalm As Boolean)

  bleftPalm = TFPalm
  FingerLocation

End Property

Public Property Get Orientation() As Orient

  Orientation = m_Orientation

End Property

Public Property Let Orientation(Ornt As Orient)

  m_Orientation = Ornt
  FingerLocation

End Property

Private Sub picFingers_Click(Index As Integer)

  Mid$(strBin, IIf(bleftPalm, 6 - Index, Index), 1) = IIf(picFingers(Index).BackColor = vbBlack, "1", "0")
  DrawHand

End Sub

Private Sub picThumb_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

  If ThumbHotSpot(X, Y) Then
    Mid$(strBin, IIf(bleftPalm, 1, 5), 1) = IIf(picThumb.BackColor = vbBlack, "1", "0")
    DrawHand
  End If

End Sub

Private Function ThumbHotSpot(ByVal X As Single, _
                              ByVal Y As Single) As Boolean

  With picThumb
    If IIf(bleftPalm, X < .Width / 2, X > .Width / 2) Then
      If IIf(m_Orientation = orUp, Y > .Height / 2, Y < .Height / 2) Then
        ThumbHotSpot = True
      End If
    End If
  End With

End Function

Private Sub UserControl_Initialize()

  Dim I As Long

'store the base position for the fingers
'This allows programatic hand reversal
  For I = 1 To 4
      With fingPos(I)
      .dTop = picFingers(I).Top
      .dLeft = picFingers(I).Left
      .dWidth = picFingers(I).Width
      .dHeight = picFingers(I).Height
    End With 'fingPos(I)
  Next I
  dcolor = &H80FF80
  bleftPalm = True
  BinaryValue = "11111"
'values to force resizing to keep for the same size
  RealWidth = picThumbSource(0).Width
  RealHeight = picThumbSource(0).Height
  UserControl.Width = picThumb.Width
  UserControl.Height = picThumb.Height
'turn off the borders (only on for ease of construction)
  For I = 1 To 4
    picFingers(I).BorderStyle = 0
  Next I
  DrawHand

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  With PropBag
    LeftPalm = .ReadProperty("Palm", True)
    BinaryValue = .ReadProperty("Bin", "11111")
    DisplayShow = .ReadProperty("VDisp", True)
    DisplayStyle = .ReadProperty("SDisp", dmBinary)
    DisplayPosition = .ReadProperty("PDisp", TopLeft)
    DisplayColor = .ReadProperty("CDisp", &H80FF80)
    Orientation = .ReadProperty("Orient", orUp)
  End With 'PropBag
  DrawHand

End Sub

Private Sub UserControl_Resize()

  UserControl.Width = RealWidth
  UserControl.Height = RealHeight

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  With PropBag
    .WriteProperty "Palm", bleftPalm, True
    .WriteProperty "Bin", BinaryValue, "11111"
    .WriteProperty "VDisp", bDisplayShow, True
    .WriteProperty "SDisp", DisplayStyle, dmBinary
    .WriteProperty "PDisp", DisplayPosition, TopLeft
    .WriteProperty "CDisp", DisplayColor, &H80FF80
    .WriteProperty "Orient", m_Orientation, orUp

  End With 'PropBag

End Sub

':)Code Fixer V2.0.0 (14/04/2004 1:09:52 PM) 71 + 396 = 467 Lines Thanks Ulli for inspiration and lots of code.

