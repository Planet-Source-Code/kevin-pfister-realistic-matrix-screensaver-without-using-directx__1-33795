VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Matrix By Kevin Pfister"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Matrix"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   ScaleHeight     =   40.25
   ScaleMode       =   4  'Character
   ScaleWidth      =   84.25
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   90
      Top             =   90
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LengthOfDrop(1 To 130, 1 To 52) As Byte    'Length of Drop
Dim Leading(1 To 130, 1 To 52) As Byte    'Is it a leading one
Dim Letter(1 To 130, 1 To 52) As Byte    'Letter
Dim Colour(1 To 130, 1 To 52) As Integer    'Colour of the letter /symbol
Dim WaitBeforeClear(1 To 130, 1 To 52) As Byte        'Wait before it dissappears
Dim F   'Max Length
Dim G   'Max Wait
Dim H   'No of Drops
Dim M   'Fade Speed
Dim O   'Fall From Top
Dim Q   'Fade?

Sub StartUp(I, J, K, L, N, P)
    ShowCursor (0)
    F = I
    G = J
    H = K
    M = L
    O = N
    Q = P
    Randomize Timer
    For DoRand = 1 To H
        XR = Int(Rnd * 130) + 1
        YR = Int(Rnd * 52) + 1
        LengthOfDrop(XR, YR) = Int(Rnd * F)
        Leading(XR, YR) = 1
        Letter(XR, YR) = Int(Rnd * 43) + 65
        Colour(XR, YR) = Int(Rnd * 200) + 55
    Next
End Sub

Private Sub Form_Click()
    ShowCursor (1)
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ShowCursor (1)
    End
End Sub

Private Sub Timer1_Timer()
    FrmMain.Cls
    For X = 1 To 130
        For Y = 1 To 52
            If Leading(X, Y) = 1 Then 'Is it leading
                If Y + 1 <= 52 Then 'Is it smaller than the screen height
                    If LengthOfDrop(X, Y) > 0 Then 'Is there still drops in this column
                        LengthOfDrop(X, Y + 1) = LengthOfDrop(X, Y) - 1
                        Leading(X, Y + 1) = 2
                        Letter(X, Y + 1) = Int(Rnd * 43) + 65
                        Colour(X, Y + 1) = Int(Rnd * 200) + 55
                        Leading(X, Y) = 0
                        WaitBeforeClear(X, Y) = G
                    Else    'End of Drop(Kill Letter/Symbol)
                        Leading(X, Y) = 0
                        WaitBeforeClear(X, Y) = G
                    End If
                Else    'End of Drop(Kill Letter/Symbol)
                    Leading(X, Y) = 0
                    WaitBeforeClear(X, Y) = G
                End If
            ElseIf WaitBeforeClear(X, Y) > 0 Then 'Is the Letter/Symbol dieing?
                WaitBeforeClear(X, Y) = WaitBeforeClear(X, Y) - 1
                If Q = 1 Then   'Is fading ativated
                    Colour(X, Y) = Colour(X, Y) - M
                End If
                If WaitBeforeClear(X, Y) = 0 Or Colour(X, Y) < 0 Then
                    Letter(X, Y) = 0
                End If
            End If
            If Leading(X, Y) = 1 Or Leading(X, Y) = 2 Then
                Leading(X, Y) = 1
                Drops = Drops + 1
            End If
            If Letter(X, Y) > 0 Then
                FrmMain.CurrentX = X
                FrmMain.CurrentY = Y - 5
                If Leading(X, Y) = 0 Then
                    FrmMain.ForeColor = RGB(0, Colour(X, Y), 0)
                Else
                    FrmMain.ForeColor = vbWhite
                End If
                FrmMain.Print Chr(Letter(X, Y))
            End If
        Next
    Next
    If Drops < H Then
        For MakeNew = Drops To H
            XR = Int(Rnd * 130) + 1
            If O = 1 Then
                YR = Int(Rnd * 5) + 1
            Else
                YR = Int(Rnd * 52) + 1
            End If
            LengthOfDrop(XR, YR) = Int(Rnd * F)
            Leading(XR, YR) = 1
            Letter(XR, YR) = 64 + Int(Rnd * 26)
            Colour(XR, YR) = Int(Rnd * 200) + 55
        Next
    End If
End Sub
