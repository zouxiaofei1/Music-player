VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "音乐播放"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3015
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始!"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "："
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Command1_Click()
Form1.Hide
Timer1.Enabled = True


'MsgBox Time()

'MsgBox minu
'If b > 2 Then hour=
End Sub

Private Sub Form_Load()

Dim a As String
a = Time()
Dim b As String
b = Left(a, 2)
Dim c As String
If Val(b) >= 10 Then c = Mid(a, 4, 2) Else c = Mid(a, 3, 2)
Text1.Text = b
Text2.Text = c

End Sub

Private Sub Timer1_Timer()
Dim a As String
a = Time()
Dim b As String
Dim hour As Integer
b = Left(a, 2)
hour = Val(b)
Dim c As String
Dim minu As Integer
If hour >= 10 Then c = Mid(a, 4, 2) Else c = Mid(a, 3, 2)
minu = Val(c)
If Val(Text1.Text) = hour And Val(Text2.Text) = minu Then

d = sndPlaySound("C:\SAtemp\1.wav", 1)
Timer1.Enabled = False
End If


End Sub
