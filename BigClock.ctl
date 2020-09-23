VERSION 5.00
Begin VB.UserControl Big7Clock 
   BackColor       =   &H00000000&
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LockControls    =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   5235
   ToolboxBitmap   =   "BigClock.ctx":0000
   Begin VB.Label BriansBrain 
      BackStyle       =   0  'Transparent
      Caption         =   "(c) - John - 2002 - briansbrainmail@yahoo.com"
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
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   1050
      Width           =   4425
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   0
      Left            =   150
      Picture         =   "BigClock.ctx":0534
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   1
      Left            =   585
      Picture         =   "BigClock.ctx":0DAC
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   2
      Left            =   1020
      Picture         =   "BigClock.ctx":1624
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   3
      Left            =   1455
      Picture         =   "BigClock.ctx":1E9C
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   4
      Left            =   1920
      Picture         =   "BigClock.ctx":2714
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   5
      Left            =   2445
      Picture         =   "BigClock.ctx":2F8C
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   6
      Left            =   2955
      Picture         =   "BigClock.ctx":3804
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   7
      Left            =   3420
      Picture         =   "BigClock.ctx":407C
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   8
      Left            =   3960
      Picture         =   "BigClock.ctx":48F4
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image BR 
      Height          =   525
      Index           =   9
      Left            =   4455
      Picture         =   "BigClock.ctx":516C
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Tenths 
      Enabled         =   0   'False
      Height          =   525
      Left            =   3150
      Picture         =   "BigClock.ctx":59E4
      Top             =   225
      Width           =   300
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   90
      Left            =   3015
      Picture         =   "BigClock.ctx":625C
      Top             =   645
      Width           =   90
   End
   Begin VB.Image Image5 
      Enabled         =   0   'False
      Height          =   90
      Left            =   945
      Picture         =   "BigClock.ctx":62C4
      Top             =   345
      Width           =   90
   End
   Begin VB.Image Hr1 
      Enabled         =   0   'False
      Height          =   690
      Left            =   45
      Picture         =   "BigClock.ctx":632C
      Top             =   45
      Width           =   390
   End
   Begin VB.Image Hr2 
      Enabled         =   0   'False
      Height          =   690
      Left            =   510
      Picture         =   "BigClock.ctx":6664
      Top             =   45
      Width           =   390
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   90
      Left            =   945
      Picture         =   "BigClock.ctx":699C
      Top             =   645
      Width           =   90
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   90
      Left            =   1980
      Picture         =   "BigClock.ctx":6A04
      Top             =   645
      Width           =   90
   End
   Begin VB.Image S2 
      Enabled         =   0   'False
      Height          =   690
      Left            =   2580
      Picture         =   "BigClock.ctx":6A6C
      Top             =   45
      Width           =   390
   End
   Begin VB.Image S1 
      Enabled         =   0   'False
      Height          =   690
      Left            =   2115
      Picture         =   "BigClock.ctx":6DA4
      Top             =   45
      Width           =   390
   End
   Begin VB.Image M2 
      Enabled         =   0   'False
      Height          =   690
      Left            =   1545
      Picture         =   "BigClock.ctx":70DC
      Top             =   45
      Width           =   390
   End
   Begin VB.Image M1 
      Enabled         =   0   'False
      Height          =   690
      Left            =   1080
      Picture         =   "BigClock.ctx":7414
      Top             =   45
      Width           =   390
   End
   Begin VB.Image S_DoubleDot 
      Enabled         =   0   'False
      Height          =   90
      Left            =   1980
      Picture         =   "BigClock.ctx":774C
      Top             =   345
      Width           =   90
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   9
      Left            =   4470
      Picture         =   "BigClock.ctx":77B4
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   8
      Left            =   3975
      Picture         =   "BigClock.ctx":7AEC
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   7
      Left            =   3435
      Picture         =   "BigClock.ctx":7E20
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   6
      Left            =   2970
      Picture         =   "BigClock.ctx":8158
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   5
      Left            =   2460
      Picture         =   "BigClock.ctx":8490
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   4
      Left            =   1935
      Picture         =   "BigClock.ctx":87C8
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   3
      Left            =   1470
      Picture         =   "BigClock.ctx":8B00
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   2
      Left            =   1035
      Picture         =   "BigClock.ctx":8E38
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   1
      Left            =   600
      Picture         =   "BigClock.ctx":9170
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image B 
      Height          =   690
      Index           =   0
      Left            =   165
      Picture         =   "BigClock.ctx":94A8
      Top             =   2985
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "Big7Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==== NOTE =====
'>>> OCXs Needed <<<
'Threed32.ocx
'Msflxgrd.ocx
' >>> briansbrainmail@yahoo.com
Option Explicit
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Property Let TimeIN(ByVal TheTime As String)
On Error Resume Next
If Hr1.Picture <> B(CInt(Mid(TheTime, 1, 1))) Then Hr1.Picture = B(CInt(Mid(TheTime, 1, 1)))
If Hr2.Picture <> B(CInt(Mid(TheTime, 2, 1))) Then Hr2.Picture = B(CInt(Mid(TheTime, 2, 1)))
If M1.Picture <> B(CInt(Mid(TheTime, 4, 1))) Then M1.Picture = B(CInt(Mid(TheTime, 4, 1)))
If M2.Picture <> B(CInt(Mid(TheTime, 5, 1))) Then M2.Picture = B(CInt(Mid(TheTime, 5, 1)))
If S1.Picture <> B(CInt(Mid(TheTime, 7, 1))) Then S1.Picture = B(CInt(Mid(TheTime, 7, 1)))
If S2.Picture <> B(CInt(Mid(TheTime, 8, 1))) Then S2.Picture = B(CInt(Mid(TheTime, 8, 1)))
If Tenths.Picture <> BR(CInt(Mid(TheTime, 10, 1))) Then Tenths.Picture = BR(CInt(Mid(TheTime, 10, 1)))
End Property










Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_Resize()
Height = 780
Width = 3525
End Sub


