VERSION 5.00
Begin VB.Form BigCForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Wow! What a big 'RED' Clock"
   ClientHeight    =   2835
   ClientLeft      =   4110
   ClientTop       =   4635
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "BigClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3840
      Top             =   585
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   165
   End
   Begin BIGCLOCK.Big7Clock Big7Clock 
      Height          =   780
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   1376
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "John (the newcomer, old but still coding)"
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   45
      TabIndex        =   5
      Top             =   2295
      Visible         =   0   'False
      Width           =   3870
   End
   Begin VB.Label BriansBrain 
      BackStyle       =   0  'Transparent
      Caption         =   "briansbrainmail@yahoo.com"
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   495
      TabIndex        =   4
      Top             =   2535
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Greets 
      BackStyle       =   0  'Transparent
      Caption         =   $"BigClock.frx":0442
      ForeColor       =   &H00FFFF80&
      Height          =   1065
      Left            =   105
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   3525
      WordWrap        =   -1  'True
   End
   Begin VB.Label CloseX 
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3435
      TabIndex        =   2
      Top             =   -15
      Width           =   180
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   1155
      Left            =   15
      Top             =   15
      Width           =   3630
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1110
      Left            =   30
      Top             =   30
      Width           =   3585
   End
   Begin VB.Label TheDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "The date will show here."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   75
      TabIndex        =   1
      ToolTipText     =   "Right click for Options"
      Top             =   75
      Width           =   3510
   End
End
Attribute VB_Name = "BigCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==== NOTE =====
'>> NO OCXs Needed <<<
' >ME> briansbrainmail@yahoo.com
Option Explicit
Dim Moving As Boolean 'Pannels Moving Yes/No
Dim NL As Integer '(X) Left, TipText & Headings Move
Dim NT As Integer '(Y) Top, TipText & Headings Move


Sub CloseXOFF()
If CloseX.ForeColor <> &HC0& Then
    CloseX.ForeColor = &HC0&
    CloseX.FontBold = False
End If
End Sub
Sub PutBrain()
TheDate.Caption = "briansbrainmail@yahoo.com"
TheDate.ForeColor = &HFF00&

End Sub

Private Sub Big7Clock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    MousePointer = 5
    Moving = True
    NL = X
    NT = Y
End If
If Button = 2 Then PutBrain
End Sub
Private Sub Big7Clock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next 'ON
CloseXOFF
If Moving = True Then
Dim MT As Integer
Dim ML As Integer
If Y < NT Then Y = NT
TryTop:
If Top + Y - NT > 0 - MT - NT And Top + Y - NT < 50 + MT - NT Then
    If MT - NT < 0 Then Top = 0: GoTo TryLft
    If MT - NT > Screen.Height - Height Then Top = Screen.Height - Height: GoTo TryLft
    If MT - NT > 0 Then Top = MT - NT: GoTo TryLft
End If
MT = MT + 45
GoTo TryTop

TryLft:
If Left + X - NL > 0 - ML - NL And Left + X - NL < 50 + ML - NL Then
    If ML - NL < 100 Then Left = 0: GoTo OutOut
    If ML - NL > (Screen.Width - Width) Then Left = (Screen.Width - Width): GoTo OutOut
    If ML - NL > 0 Then Left = ML - NL: GoTo OutOut
End If
ML = ML + 45
GoTo TryLft

OutOut:

End If

End Sub


Private Sub Big7Clock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Moving = False
    MousePointer = 0
End If
If Button = 2 Then TheDate.ForeColor = &HFF& 'Red
End Sub





Private Sub CloseX_Click()
On Error Resume Next
DoEvents
Me.Hide
DoEvents
Unload Me
End Sub

Private Sub CloseX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CloseX.FontBold = True
CloseX.ForeColor = &HFFFFFF

End Sub

Private Sub Form_Load()
Big7Clock.TimeIN = Format(Time, "hh:nn:ss") & "." & Right(Format(Timer, "#.0"), 1)
DoEvents
Width = 3610 + 30
Height = 1125 + 45
Left = Screen.Width / 2 - Width / 2
Top = 0

TheDate.Caption = Format(Date, "dddd  d  mmmm  yyyy")


End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseXOFF
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Enabled = False
Timer2.Enabled = False
End Sub







Private Sub TheDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    MousePointer = 5
    Moving = True
    NL = X
    NT = Y
End If
If Button = 2 Then PutBrain

End Sub



Private Sub TheDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Big7Clock_MouseMove Button, Shift, X, Y
CloseXOFF
End Sub

Private Sub TheDate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Big7Clock_MouseUp Button, Shift, X, Y

If Button = 2 Then TheDate.ForeColor = &HFF& 'Red

End Sub

Private Sub Timer1_Timer()
Big7Clock.TimeIN = Format(Time, "hh:nn:ss") & "." & Right(Format(Timer, "#.0"), 1)

End Sub


Private Sub Timer2_Timer()
On Error Resume Next
If TheDate.ForeColor = &HFF& Then 'RED
    If TheDate.Caption <> Format(Date, "dddd  d  mmmm  yyyy") Then TheDate.Caption = Format(Date, "dddd  d  mmmm  yyyy")
End If
End Sub


