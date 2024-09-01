VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   360
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   360
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   360
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   1245
      Index           =   0
      Left            =   4320
      Picture         =   "random drop.frx":0000
      Top             =   240
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   0
      Left            =   2400
      Picture         =   "random drop.frx":236A
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim kount As Double

Private Sub Form_Load()
    For i = 1 To 3
        Load Image1(i)
        Load Timer1(i)
        Load Image2(i)
        Load Timer2(i)
    Next
End Sub

Private Sub Command1_Click()
    Randomize
    x = Rnd * 5
    Print x
    Print Int((Rnd * 9720) + 2400)
    For v = 0 To 2
        If x = v Then
            Image1(v).Left = Int((Rnd * 7320) + 2400)
            Image1(v).Top = 360
            Image1(v).Visible = True
            Timer1(v).Enabled = True
        ElseIf x - 3 = v Then
            Image2(v).Left = Int((Rnd * 7320) + 2400)
            Image2(v).Top = 360
            Image2(v).Visible = True
            Timer2(v).Enabled = True
        End If
    Next
End Sub

Private Sub Timer1_Timer(Index As Integer)
    Image1(Index).Move Image1(Index).Left + 0, Image1(Index).Top + 100
End Sub
Private Sub Timer2_Timer(Index As Integer)
    Image2(Index).Move Image2(Index).Left + 0, Image2(Index).Top + 100
End Sub

Private Sub Timer3_Timer()
    kount = kount + 0.5
    s = Rnd * 3
    Print s
    If s = 1 Then
        r = Rnd * 2
        Image2(r).Left = Int((Rnd * 7320) + 2400)
        Image2(r).Top = 360
        Image2(r).Visible = True
        Timer2(r).Enabled = True
    Else
        r = Rnd * 2
        Image1(r).Left = Int((Rnd * 7320) + 2400)
        Image1(r).Top = 360
        Image1(r).Visible = True
        Timer1(r).Enabled = True
End Sub
