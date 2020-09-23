VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   855
   ClientTop       =   735
   ClientWidth     =   10215
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10215
   Begin MSComctlLib.Slider grav 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Max             =   30
      SelStart        =   1
      Value           =   1
   End
   Begin VB.PictureBox can 
      Height          =   495
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag n Drop The Ball On The Arrow"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gravity"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image ball 
      Appearance      =   0  'Flat
      DragIcon        =   "Form1.frx":0442
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1200
      Picture         =   "Form1.frx":0884
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gravity Effect By George Papadopoulos
'Please Visit Http://www.g-soft.cjb.net
Dim more As Integer
Dim lef As Integer, rit As Integer
Dim dis As Integer
Dim stopit As Boolean
Private Sub can_DragDrop(Source As Control, X As Single, Y As Single)
ball.Visible = True
ball.Top = 0
ball.Left = 0
ball.Tag = "right"
a = Int(Rnd * 100) + 30
more = grav.Value
lef = a
rit = a
dis = 50
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Source.Visible = True
Source.Top = Y - Source.Height / 2
Source.Left = X - Source.Width / 2
more = grav.Value
dis = 50
End Sub

Private Sub ball_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
ball.Visible = False
End Sub

Private Sub Form_Load()
Randomize
more = 50
dis = 50
Me.Show
stopit = False
Do
DoEvents
e = e + 0.05
If e >= 100 Then
    movex
    movey
    e = 0
End If
Loop Until stopit = True
MsgBox "Gravity Effect By George Papadopoulos"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
stopit = True
End Sub

Private Sub grav_Scroll()
grav.Tag = grav.Value
End Sub

Private Sub movey()
On Error Resume Next
If ball.Visible = True Then
    ball.Visible = False
    more = more + Val(grav.Tag)
    ball.Top = ball.Top + more
    If ball.Top + ball.Height > Me.Height - 350 Then
        lef = lef - 3
        If lef < 0 Then lef = 0
        rit = rit - 3
        If rit < 0 Then rit = 0
        
        ball.Top = Me.Height - ball.Height - 351
        dis = dis + grav.Value
        more = -more + dis * 1.5
    End If
    ball.Visible = True
End If
End Sub

Private Sub movex()
On Error Resume Next
Select Case ball.Tag
    Case "left"
        ball.Left = ball.Left - lef
        If ball.Left < 0 Then
            ball.Tag = "right"
            lef = lef - 20
            If lef < 0 Then lef = 0
        End If
    Case "right"
        ball.Left = ball.Left + rit
        If ball.Left + ball.Width > Me.Width Then
            ball.Tag = "left"
            rit = rit - 20
            If rit < 0 Then rit = 0
        End If
        
End Select
End Sub
