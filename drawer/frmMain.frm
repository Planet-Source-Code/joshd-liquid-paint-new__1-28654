VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.OptionButton optCircle 
      Caption         =   "Circle"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton optSquare 
      Caption         =   "Square"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.FileListBox filScheme 
      Height          =   1455
      Left            =   3600
      Pattern         =   "*.col"
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.VScrollBar scrSize 
      Height          =   3015
      Left            =   3240
      Max             =   40
      Min             =   1
      TabIndex        =   1
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.Timer tmrDraw 
      Interval        =   50
      Left            =   3120
      Top             =   720
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   120
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
   Begin VB.Label Label1 
      Caption         =   "Colour Scheme"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pixel(0 To 199, 0 To 199) As Integer
Dim Colours(1 To 256) As Long
Dim currentCol
Dim LastX As Single, LastY As Single
Dim MaxCols As Integer

Private Sub filScheme_Click()
    LoadColours (filScheme.path & "/" & filScheme.filename)
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer
    For i = 0 To picDraw.Width
        For j = 0 To picDraw.Height
            Pixel(i, j) = 0
        Next j
    Next i
    currentCol = 1
    filScheme.path = App.path
    LoadColours (App.path & "/default.col")
End Sub
Public Sub increment()
    Dim i As Integer
    Dim tempCol As Long
    tempCol = Colours(1)
    For i = 1 To MaxCols - 1
        Colours(i) = Colours(i + 1)
    Next i
    Colours(MaxCols) = tempCol
End Sub
Public Sub Draw()
    Dim x As Integer, y As Integer
    For x = 0 To picDraw.Width
        For y = 0 To picDraw.Height
            If Pixel(x, y) <> 0 Then
                Call SetPixel(picDraw.hdc, x, y, Colours(Pixel(x, y)))
            End If
        Next y
    Next x
   picDraw.Refresh
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call DrawPoint(x, y)
    LastX = x
    LastY = y
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call FillLine(x, y, LastX, LastY)
        LastX = x
        LastY = y
    End If
End Sub

Private Sub tmrDraw_Timer()
    Draw
    increment
    DoEvents
End Sub
Private Sub DrawPoint(x As Single, y As Single)
    If optDelete.Value = True Then
        Call DeleteSquare(x, y)
    ElseIf scrSize.Value = 1 Then
        Call DrawPixel(x, y)
    ElseIf optSquare.Value = True Then
        Call DrawSquare(x, y)
    ElseIf optCircle.Value = True Then
        Call DrawCircle(x, y)
    End If
End Sub
Public Sub DrawPixel(x As Single, y As Single)
    If x <= picDraw.Width And x >= 0 And y <= picDraw.Height And y >= 0 Then
        Pixel(x, y) = currentCol
    End If
    currentCol = currentCol + 1
    If currentCol > MaxCols Then currentCol = 1
End Sub
Public Sub DrawSquare(x As Single, y As Single)
    Dim xMax As Integer, yMax As Integer
    Dim thisX As Integer, thisY As Integer
    x = x - (scrSize.Value / 2)
    y = y - (scrSize.Value / 2)
    xMax = x + scrSize.Value - 1
    yMax = y + scrSize.Value - 1
    For thisX = x To xMax
        For thisY = y To yMax
            If thisX <= picDraw.Width And thisX >= 0 And thisY <= picDraw.Height And thisY >= 0 Then
                Pixel(thisX, thisY) = currentCol
            End If
        Next thisY
    Next thisX
    currentCol = currentCol + 1
    If currentCol > MaxCols Then currentCol = 1
End Sub
Public Sub DeleteSquare(x As Single, y As Single)
    Dim xMax As Integer, yMax As Integer
    Dim thisX As Integer, thisY As Integer
    x = x - (scrSize.Value / 2)
    y = y - (scrSize.Value / 2)
    xMax = x + scrSize.Value - 1
    yMax = y + scrSize.Value - 1
    For thisX = x To xMax
        For thisY = y To yMax
            If thisX <= picDraw.Width And thisX >= 0 And thisY <= picDraw.Height And thisY >= 0 Then
                Pixel(thisX, thisY) = 0
                Call SetPixel(picDraw.hdc, thisX, thisY, 0)
            End If
        Next thisY
    Next thisX
    picDraw.Refresh
End Sub
Public Sub DrawCircle(x As Single, y As Single)
    Dim xMax As Integer, yMax As Integer
    Dim thisX As Integer, thisY As Integer
    Dim changeX As Integer, changeY As Integer
    x = x
    y = y
    xMax = x + (scrSize.Value / 2)
    yMax = y + (scrSize.Value / 2)
    For thisX = x - (scrSize.Value / 2) To xMax
        For thisY = y - (scrSize.Value / 2) To yMax
            If thisX <= picDraw.Width And thisX >= 0 And thisY <= picDraw.Height And thisY >= 0 Then
                changeX = thisX - x
                changeY = thisY - y
                If Sqr(changeX * changeX + changeY * changeY) < (scrSize.Value / 2) Then Pixel(thisX, thisY) = currentCol
            End If
        Next thisY
    Next thisX
    currentCol = currentCol + 1
    If currentCol > MaxCols Then currentCol = 1
End Sub
Public Sub FillLine(x As Single, y As Single, LastX As Single, LastY As Single)
    Dim MaxChange As Integer, XChange As Integer, yChange As Integer
    Dim thisX As Single, thisY As Single
    XChange = LastX - x
    yChange = LastY - y
    If Abs(yChange) < Abs(XChange) Then
        MaxChange = Abs(XChange)
    Else
        MaxChange = Abs(yChange)
    End If
    
    For i = 1 To MaxChange
        thisX = (XChange / i) + x
        thisY = (yChange / i) + y
        Call DrawPoint(thisX, thisY)
    Next i
End Sub

Public Sub LoadColours(path As String)
    Dim rCol As Integer, gCol As Integer, bColour As Integer
    MaxCols = 0
    Open path For Input As #1
        Do While EOF(1) = False
            Input #1, rCol, gCol, bCol
            MaxCols = MaxCols + 1
            Colours(MaxCols) = RGB(rCol, gCol, bCol)
        Loop
    Close #1
End Sub
