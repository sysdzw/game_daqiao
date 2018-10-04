VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "搭桥游戏 V2.0"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   16035
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Interval        =   800
      Left            =   1440
      Top             =   1800
   End
   Begin VB.Frame fraScore 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFD088&
      BorderStyle     =   0  'None
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   10560
      TabIndex        =   4
      Top             =   480
      Width           =   5295
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Height          =   585
         Left            =   0
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label labLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   600
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label labProc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   0
         TabIndex        =   6
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label lblFen 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "已通过墙块数: 0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   1320
         TabIndex        =   5
         Tag             =   "0"
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   840
      Top             =   1800
   End
   Begin VB.PictureBox picKey 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H008080FF&
      FillStyle       =   3  'Vertical Line
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picStone 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   2895
      Index           =   0
      Left            =   360
      ScaleHeight     =   2865
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "背景音效"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   120
      TabIndex        =   10
      Top             =   2768
      Width           =   1200
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   1440
      Tag             =   "1"
      Top             =   2775
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   1440
      Tag             =   "1"
      Top             =   2295
      Width           =   1230
   End
   Begin VB.Image src2 
      Height          =   390
      Left            =   120
      Picture         =   "frmMain.frx":08CA
      Top             =   3840
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image src1 
      Height          =   390
      Left            =   120
      Picture         =   "frmMain.frx":0B13
      Top             =   3360
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "背景音乐"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   120
      TabIndex        =   9
      Top             =   2288
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作者：sysdzw && Chen8013"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   2070
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   795
      Left            =   360
      Picture         =   "frmMain.frx":0DC7
      Stretch         =   -1  'True
      Top             =   5565
      Width           =   315
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "【游戏说明】点击小人所在的墙块，鼠标左键摁住时间长短可控制桥的长度，规则就是这么简单，小人的小命就掌握在你的手中了！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   315
      Width           =   4935
   End
   Begin VB.Image imgGlass 
      Enabled         =   0   'False
      Height          =   2430
      Left            =   120
      Picture         =   "frmMain.frx":1074
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   16500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "提示信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1125
      Left            =   5100
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image imgSky 
      Enabled         =   0   'False
      Height          =   4020
      Left            =   360
      Picture         =   "frmMain.frx":4791
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'名    称：vb搭桥小游戏
'描    述：这是网页上比较常见的小游戏，今天闲来用vb写了下，还是有点不足
'使用方法：点击小人所在的墙块，鼠标摁住的时间长短控制桥的长度，以保证小人顺利通过
'编    程：sysdzw 原创开发，Chen8013做了增强改进基本重写了，如果有需要对模块扩充或更新的话请邮箱发我一份
'发布日期：2017-03-06
'博    客：http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0   sysdzw初版                                                        2017-03-07
'版    本：V2.0   感谢网友Chen8013的加入做了大幅度改进，增加了卷轴闯关等功能        2017-03-09
'                 sysdzw对部分资源文件进行了调整，加入了一些音效
'版    本：V2.1   修复了音乐不能循环播放的问题                                      2017-03-12
'==============================================================================================
Option Explicit

Private Type STONE      ' “墙块”的关键信息参数
   cw    As Long     ' 宽度
   cxC   As Long     ' 中间坐标
   cxL   As Long     ' 左边界坐标
   cxR   As Long     ' 右边界坐标
End Type

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundW" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySoundStop Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long

Private Const SCORE_TEXT      As String = "已通过墙块数："        ' 记分牌提示文本

Private Const SND_ASYNC       As Long = &H1           ' 异步播放，否则就独占播放
Private Const SND_NODEFAULT   As Long = &H2           ' 不使用缺省声音
Private Const SND_MEMORY      As Long = &H4           ' 指向一个内存文件
Private Const SND_FILENAME    As Long = &H20000       ' 指向一个实际文件
Private Const SND_LOOP        As Long = &H8           ' 循环播放
Private Const SND_ALIAS_START As Long = 0             ' 结束播放
Private Const SND_SYNC        As Long = &H0

Private Const NUM_LEVEL       As Long = 8             ' 预设关卡数
Private Const ADD_SPEED       As Long = 300           ' 移动步长
Private Const DISP_FIX        As Long = 210           ' 角色显示修正量
Private Const START_CX        As Long = 450           ' 第1块“墙” 的左边界坐标
Private Const SPACE_END       As Long = 570           ' 场景右端预留宽度
Private Const BASE_LINE       As Long = 6360          ' 基准线 y坐标值
Private Const MAX_SIZE        As Long = 6300          ' 最大“桥长”
Private Const PROC_WIDTH      As Long = 5295          ' “记分牌”宽度

Private arrStone()   As STONE ' 加载的“墙”的信息参数
Private arrLevStep() As Long  ' “墙块”数量（步数）
Private arrLevDomW() As Long  ' 动态宽度值
Private arrLevMinW() As Long  ' 最小宽度值
Private arrSound()   As Byte  ' 声音资源数据
Private arrSound2()   As Byte  ' 声音资源数据
Private arrSound3()   As Byte  ' 声音资源数据

Private mlOpeLock As Long     ' 操作逻辑锁
Private mlCFXH    As Long     ' 窗口高度修正量
Private mlCFXW    As Long     ' 窗口宽度修正量
Private mlClientH As Long     ' 窗口客户区高度
Private mlClientW As Long     ' 窗口客户区宽度
Private mlSceneW  As Long     ' 卷轴场景宽度
Private mlSceneX  As Long     ' 卷轴场景位置
Private mlDispX   As Long     ' 角色显示位置
Private mlRoleX   As Long     ' 角色在场景中的位置
Private mlLevel   As Long     ' 关卡索引值
Private mlCritiL  As Long     ' 左临界点
Private mlCritiR  As Long     ' 右临界点
Private mlStoneN  As Long     ' “墙块”数据信息下标的上界值
Private mlObjPnt  As Long     ' 加载的对象索引上界
Private mlDispPnt As Long     ' 显示的对象上界值
Private mlCurrPnt As Long     ' 当前所在“位置”
Private mlCurrObj As Long     ' 当前在哪个“墙块”上
Private mlCurrLen As Long     ' 当前“桥长”的值
Private mlKeyLeft As Long     ' “桥”的左端坐标
Private mlMoveToX As Long     ' 角色将“走到”的位置
Private mlFunFlag As Long     ' 操作功能标识

Dim Buffer As String * 128
Dim Ret As Long
Dim strAppPath As String  '应用程序目录

Private Function DispScene() As Boolean
   Dim i&, k&, w&, u As Long

   For i = 0& To mlStoneN
      If (mlSceneX < arrStone(i).cxR) Then Exit For
   Next
   w = i
   k = -1&
   Do
      k = 1& + k
      If (k > mlObjPnt) Then              ' 需要更多的“墙块”
         mlObjPnt = 4& + mlObjPnt         ' 增加 4个
         For u = k To mlObjPnt
            Call Load(picStone(u))
         Next
      End If
      u = arrStone(i).cxL - mlSceneX      ' 左边界显示坐标值
      If (u > mlClientW) Then k = k - 1&: Exit Do
      'picStone(k).Left = u
      'picStone(k).Width = arrStone(i).cw
      Call picStone(k).Move(u, BASE_LINE, arrStone(i).cw)
      If (i = mlStoneN) Then Exit Do      ' 显示到最后一个了
      i = 1& + i                          ' 处理“下一个”
   Loop
   If (mlDispPnt < k) Then
      u = 30& + mlClientH - BASE_LINE
      For i = mlDispPnt To k
         picStone(i).Visible = True
         picStone(i).Height = u
      Next
   Else
      For u = 1& + k To mlDispPnt
         picStone(u).Visible = False
      Next
   End If
   mlDispPnt = k
   mlCurrObj = mlCurrPnt - w
   
   mlDispX = mlRoleX - mlSceneX           ' 角色显示位置
   Image1.Left = mlDispX - DISP_FIX       ' 在相应位置显示角色

   DispScene = (SPACE_END < mlClientW + mlSceneX - arrStone(mlStoneN).cxR)
End Function

Private Sub InitData()        ' 关卡预设参数初始化过程
   Dim u  As Long

   u = NUM_LEVEL - 1&         ' 预设8关数据
   ReDim arrLevStep(u)
   ReDim arrLevDomW(u)
   ReDim arrLevMinW(u)
   ' “墙块”数量（步数）
   arrLevStep(0&) = 20&
   arrLevStep(1&) = 35&       ' +15
   arrLevStep(2&) = 60&       ' +25
   arrLevStep(3&) = 100&      ' +40
   arrLevStep(4&) = 155&      ' +60
   arrLevStep(5&) = 240&      ' +85
   arrLevStep(6&) = 360&      ' +120
   arrLevStep(7&) = 525&      ' +165
    ' 动态宽度值
   arrLevDomW(0&) = 1650&
   arrLevDomW(1&) = 1420&
   arrLevDomW(2&) = 1200&
   arrLevDomW(3&) = 1050&
   arrLevDomW(4&) = 900&
   arrLevDomW(5&) = 750&
   arrLevDomW(6&) = 630&
   arrLevDomW(7&) = 525&
   ' 最小宽度值
   arrLevMinW(0&) = 270&
   arrLevMinW(1&) = 240&
   arrLevMinW(2&) = 210&
   arrLevMinW(3&) = 180&
   arrLevMinW(4&) = 150&
   arrLevMinW(5&) = 120&
   arrLevMinW(6&) = 105&
   arrLevMinW(7&) = 90&
End Sub

Private Sub LoadLevel()
   Dim u&, i As Long
   Dim w&, n As Long
   Dim k&, v As Long
   Dim cx&, dw As Long

   u = arrLevStep(mlLevel)    ' 加载数量
   w = arrLevMinW(mlLevel)    ' 最小宽度
   n = arrLevDomW(mlLevel)    ' 动态宽度
   k = 270&                   ' 最小间距（18像素）
   v = 900&                   ' 动态间距（60像素）
   mlStoneN = u
   ReDim arrStone(u)
   dw = 0.7 * n               ' 第1块的特定设置：宽度 70%动态宽度
   mlRoleX = START_CX + dw \ 2&
   cx = START_CX + dw
   arrStone(0&).cw = dw
   arrStone(0&).cxC = mlRoleX
   arrStone(0&).cxL = START_CX
   arrStone(0&).cxR = cx
   Call Randomize             ' 随机化
   For i = 1& To u
      dw = k + v * Rnd()      ' 随机“间距”
      cx = cx + dw
      arrStone(i).cxL = cx    ' 左边界
      dw = w + n * Rnd()      ' 随机“宽度”
      arrStone(i).cxC = cx + dw \ 2&      ' 中间点
      cx = cx + dw
      arrStone(i).cw = dw
      arrStone(i).cxR = cx    ' 右边界
   Next
   v = n \ 2&
   If (v > dw) Then           ' 最后一块太窄就重新调整
      arrStone(u).cw = v
      arrStone(u).cxR = cx - dw + v
   End If
   mlSceneW = SPACE_END + cx  ' 卷轴场景总宽度
   mlSceneX = 0&              ' 卷轴场景初始位置
   mlCurrPnt = 0&             ' 起始原点
   mlCurrObj = 0&             ' 位于“第1块”
   labLevel.Caption = "第" & (1& + mlLevel) & "關"
   labProc.Width = 0&
End Sub

Private Sub ReStart()
   If (mlOpeLock) Then Exit Sub
   lblFen.Caption = SCORE_TEXT & "0"
   labProc.Width = 0          ' “进度”清0
   Label1.Visible = False
   Image1.Top = BASE_LINE - Image1.Height
   Image1.Visible = True
   picKey.Visible = False
   mlCurrLen = 0&             ' “桥长”置０
   mlCurrPnt = 0&             ' 起始原点
   mlSceneX = 0&              ' 卷轴场景位置
   mlRoleX = arrStone(0&).cxC ' 角色起点
   Call DispScene             ' 刷新场景
   mlOpeLock = vbFalse
End Sub

Private Sub Form_Click()
   If (mlOpeLock) Then Exit Sub           ' 操作“锁定”
   If Label1.Visible Then Call ReStart    ' 失败后“重新开始”
End Sub

Private Sub Form_Load()
   Dim i As Long

   arrSound = LoadResData(101, "CUSTOM")
   arrSound2 = LoadResData(102, "CUSTOM")
   arrSound3 = LoadResData(103, "CUSTOM")
   mlClientH = Me.ScaleHeight
   mlCFXW = Me.Width - Me.ScaleWidth
   mlCFXH = Me.Height - mlClientH
   Label1.Left = 0
   mlLevel = 0&
   mlObjPnt = 23&             ' 设置“预加载对象”数量
   For i = 1& To mlObjPnt:    Load picStone(i):    Next
   Call InitData              ' 初始化“关卡数据”
   Call LoadLevel             ' 加载数据
   mlOpeLock = vbFalse
   mlFunFlag = vbTrue
   
   Set Image2.Picture = src1.Picture
   Set Image3.Picture = src1.Picture
   
    strAppPath = App.Path
    If Right(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    Call playBackSound
End Sub

Private Sub Form_Resize()
   Const MIN_CLIENT_H   As Long = 9000    ' 600像素
   Const MIN_CLIENT_W   As Long = 13800   ' 920像素
   Dim lFlag   As Long
   Dim cw&, ch As Long

   If (-2& = mlOpeLock) Then Exit Sub
   mlOpeLock = -2&            ' 防止重入
   lFlag = vbFalse
   cw = Me.ScaleWidth
   ch = Me.ScaleHeight
   If (MIN_CLIENT_W > cw) Then
      Me.Enabled = False
      Me.Width = MIN_CLIENT_W + mlCFXW
      lFlag = vbTrue
      cw = MIN_CLIENT_W
   End If
   If (MIN_CLIENT_H > ch) Then
      Me.Enabled = False
      Me.Height = MIN_CLIENT_H + mlCFXH
      lFlag = vbTrue
      ch = MIN_CLIENT_H
   End If
   If (lFlag) Then            ' 如果“重新设置”过窗口大小
      DoEvents
      Me.Enabled = True
   End If
   fraScore.Left = cw - PROC_WIDTH - 120
   imgGlass.Top = ch - imgGlass.Height
   imgGlass.Width = cw
   Label1.Width = cw
   imgSky.Width = cw
   lFlag = mlClientH          ' 保存之前的窗口客户区高度
   mlClientW = cw             ' 记录窗口客户区宽度和高度
   mlClientH = ch
   mlCritiL = cw \ 4&         ' 计算“场景卷轴临界点”
   mlCritiR = cw - cw \ 5&
   Call DispScene             ' 刷新显示
   
   If (ch > lFlag) Then       ' 如果窗口“高度增加”时，调整“墙块高度”
      ch = ch - picStone(0&).Top + 15
      For cw = 0& To mlDispPnt
          picStone(cw).Height = ch
      Next
   End If
   mlOpeLock = vbFalse        ' 解除锁定
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Dir(strAppPath & "bg.mid") <> "" Then mciSendString "stop mp3", Buffer, Ret, 0
End Sub

Private Sub Image2_Click()
    If Image2.Tag = "" Or Image2.Tag = 0 Then
        Call playBackSound
        Image2.Tag = 1
        Set Image2.Picture = src1.Picture
        Timer3.Enabled = True
    ElseIf Image2.Tag = 1 Then '当前为开，需要关掉
        Timer3.Enabled = False
        If Dir(strAppPath & "bg.mid") <> "" Then mciSendString "stop mp3", Buffer, Ret, 0
        Image2.Tag = 0
        Set Image2.Picture = src2.Picture
    End If
End Sub

Private Sub Image3_Click()
    If Image3.Tag = "" Or Image3.Tag = 0 Then
        Image3.Tag = 1
        Set Image3.Picture = src1.Picture
    ElseIf Image3.Tag = 1 Then '当前为开，需要关掉
        sndPlaySoundStop 0, SND_SYNC
        Image3.Tag = 0
        Set Image3.Picture = src2.Picture
    End If
End Sub

Private Sub Label1_Click()
    Call Form_Click
End Sub

Private Sub picStone_MouseDown(Index As Integer, Button As Integer, _
                           Shift As Integer, X As Single, Y As Single)
   If (mlOpeLock) Then Exit Sub
   If Label1.Visible Then     ' 游戏失败重置
       Call ReStart
       Exit Sub
   End If
   If (mlCurrObj = Index) Then
      mlCurrLen = 0&
      mlFunFlag = 0&
      mlKeyLeft = arrStone(mlCurrPnt).cxC - mlSceneX - 60&
      Image1.Top = BASE_LINE - Image1.Height
      Call picKey.Move(mlKeyLeft, BASE_LINE, 135, 0)
      picKey.Visible = True
      Timer1.Enabled = True
   Else
      MsgBox "点错墙块啦！你得点当前小人所在的墙块呀", vbExclamation
   End If
End Sub

Private Sub picStone_MouseUp(Index As Integer, Button As Integer, _
                           Shift As Integer, X As Single, Y As Single)
   If (mlOpeLock) Then Exit Sub
   If (ADD_SPEED < mlCurrLen) Then
      Call picKey.Move(mlKeyLeft, BASE_LINE - 120&, mlCurrLen, 135&)
      Image1.Top = picKey.Top - Image1.Height         '小人上桥
      mlMoveToX = mlKeyLeft + mlCurrLen
      If (mlSceneX + mlMoveToX > arrStone(mlStoneN).cxR) Then
         mlMoveToX = 120& + arrStone(mlStoneN).cxR - mlSceneX
         picKey.Width = mlMoveToX - mlKeyLeft
      End If
      mlFunFlag = 1&
      mlOpeLock = vbTrue
      Timer1.Enabled = True
      If Image3.Tag = 1 Then sndPlaySound arrSound2(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
   Else
      mlFunFlag = vbTrue
      Timer1.Enabled = False
      picKey.Visible = False
      Image1.Top = BASE_LINE - Image1.Height
   End If
End Sub

Private Sub Timer1_Timer()
   Dim w As Long

   Select Case mlFunFlag
      Case 0&: w = ADD_SPEED + mlCurrLen
               mlCurrLen = w
               picKey.Top = BASE_LINE - w
               picKey.Height = w
               If (MAX_SIZE = w) Then Timer1.Enabled = False
      Case 1&: w = ADD_SPEED + mlDispX
               If (w < mlMoveToX) Then
                  mlDispX = w
                  Image1.Left = w - DISP_FIX
               Else
                  mlCurrLen = 0&
                  Timer1.Enabled = False
                  w = Screen.TwipsPerPixelX
                  mlRoleX = (mlSceneX + mlMoveToX) \ w            ' 把“角色终点”圆整到整数像素处
                  mlRoleX = mlRoleX * w
                  Image1.Left = mlMoveToX - DISP_FIX
                  If (mlRoleX > arrStone(mlStoneN).cxR) Then      ' 已经超过“尽头”
                     w = vbTrue           ' 标识“失败”
                  Else
                     For w = mlCurrPnt To mlStoneN                ' 检测走到哪一块上
                        If (mlRoleX < arrStone(w).cxR) Then       ' 位置在某块“右边界之左”
                           If (mlRoleX > arrStone(w).cxL) Then    ' 并且在某块“左边界之右”
                              lblFen.Caption = SCORE_TEXT & w
                              labProc.Width = PROC_WIDTH * w / mlStoneN
                              mlCurrPnt = w
                              w = vbFalse ' 标识“成功”
                           Else
                              w = vbTrue  ' 标识“失败”
                           End If
                           Exit For
                        End If
                     Next
                  End If
                  If (w) Then          ' 失败！
                     If Image3.Tag = 1 Then sndPlaySound arrSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
                     mlFunFlag = 2&    ' 进行“掉落”动画处理
                     Timer1.Enabled = True
                     Exit Sub
                  Else
                    If Image3.Tag = 1 Then sndPlaySound arrSound3(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
                  End If
                  If (mlStoneN = mlCurrPnt) Then
                     labProc.Width = PROC_WIDTH
                     MsgBox "恭喜过关！　点“确定”进入下一关继续挑战。", 64&
                     mlOpeLock = vbFalse
                     mlLevel = 1& + mlLevel
                     If (NUM_LEVEL = mlLevel) Then
                        MsgBox "您真牛！已经通过了所有关卡！", 64&
                     Else
                        Call LoadLevel
                        Call ReStart
                     End If
                     Exit Sub
                  Else
                     For w = 0& To mlStoneN
                        If (mlSceneX < arrStone(w).cxR) Then
                           mlCurrObj = mlCurrPnt - w
                           Exit For
                        End If
                     Next
                  End If
                  If (mlDispX > mlCritiR) Then
                     Timer2.Enabled = True            ' 开启“卷轴处理”
                  Else
                     mlOpeLock = vbFalse
                  End If
               End If
      Case 2&: Image1.Top = Image1.Top * 1.15
               If (mlClientH < Image1.Top) Then
                  Label1.Caption = "失败了！点击任意地方重来~"
                  Label1.Visible = True
                  Image1.Visible = False
                  Timer1.Enabled = False
                  mlOpeLock = vbFalse
               End If
      
      Case Else:  mlFunFlag = vbTrue
                  Timer1.Enabled = False

   End Select

End Sub

Private Sub Timer2_Timer()    ' “卷轴”处理
   If (mlDispX > mlCritiL) Then
      picKey.Left = picKey.Left - ADD_SPEED
      mlSceneX = ADD_SPEED + mlSceneX
      If (DispScene()) Then
         mlOpeLock = vbFalse
         Timer2.Enabled = False
      End If
   Else
      mlOpeLock = vbFalse
      Timer2.Enabled = False
   End If
End Sub
Public Function GetPlayMode() As String '得到播放状态
    Dim Buffer As String * 128
    Dim pos As Integer
    mciSendString "status mp3 mode", Buffer, 128, 0&
    pos = InStr(Buffer, Chr(0))
    GetPlayMode = Left(Buffer, pos - 1)
End Function
'播放背景音乐
Private Sub playBackSound()
    If Dir(strAppPath & "bg.mid") <> "" Then
        mciSendString "close mp3", Buffer, Ret, 0
        mciSendString "open bg.mid alias mp3", Buffer, Ret, 0
        mciSendString "play mp3", Buffer, Ret, 0
    End If
End Sub
Private Sub Timer3_Timer()
    If GetPlayMode = "stopped" Then Call playBackSound
End Sub
