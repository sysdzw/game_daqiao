VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "������Ϸ V2.0"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   16035
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "��ͨ��ǽ����: 0"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "������Ч"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "���ߣ�sysdzw && Chen8013"
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
      Caption         =   "����Ϸ˵�������С�����ڵ�ǽ�飬��������סʱ�䳤�̿ɿ����ŵĳ��ȣ����������ô�򵥣�С�˵�С������������������ˣ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��ʾ��Ϣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
'��    �ƣ�vb����С��Ϸ
'��    ����������ҳ�ϱȽϳ�����С��Ϸ������������vbд���£������е㲻��
'ʹ�÷��������С�����ڵ�ǽ�飬�����ס��ʱ�䳤�̿����ŵĳ��ȣ��Ա�֤С��˳��ͨ��
'��    �̣�sysdzw ԭ��������Chen8013������ǿ�Ľ�������д�ˣ��������Ҫ��ģ���������µĻ������䷢��һ��
'�������ڣ�2017-03-06
'��    �ͣ�http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ��sysdzw@163.com
'QQ      ��171977759
'��    ����V1.0   sysdzw����                                                        2017-03-07
'��    ����V2.0   ��л����Chen8013�ļ������˴���ȸĽ��������˾��ᴳ�صȹ���        2017-03-09
'                 sysdzw�Բ�����Դ�ļ������˵�����������һЩ��Ч
'��    ����V2.1   �޸������ֲ���ѭ�����ŵ�����                                      2017-03-12
'==============================================================================================
Option Explicit

Private Type STONE      ' ��ǽ�顱�Ĺؼ���Ϣ����
   cw    As Long     ' ���
   cxC   As Long     ' �м�����
   cxL   As Long     ' ��߽�����
   cxR   As Long     ' �ұ߽�����
End Type

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundW" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySoundStop Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long

Private Const SCORE_TEXT      As String = "��ͨ��ǽ������"        ' �Ƿ�����ʾ�ı�

Private Const SND_ASYNC       As Long = &H1           ' �첽���ţ�����Ͷ�ռ����
Private Const SND_NODEFAULT   As Long = &H2           ' ��ʹ��ȱʡ����
Private Const SND_MEMORY      As Long = &H4           ' ָ��һ���ڴ��ļ�
Private Const SND_FILENAME    As Long = &H20000       ' ָ��һ��ʵ���ļ�
Private Const SND_LOOP        As Long = &H8           ' ѭ������
Private Const SND_ALIAS_START As Long = 0             ' ��������
Private Const SND_SYNC        As Long = &H0

Private Const NUM_LEVEL       As Long = 8             ' Ԥ��ؿ���
Private Const ADD_SPEED       As Long = 300           ' �ƶ�����
Private Const DISP_FIX        As Long = 210           ' ��ɫ��ʾ������
Private Const START_CX        As Long = 450           ' ��1�顰ǽ�� ����߽�����
Private Const SPACE_END       As Long = 570           ' �����Ҷ�Ԥ�����
Private Const BASE_LINE       As Long = 6360          ' ��׼�� y����ֵ
Private Const MAX_SIZE        As Long = 6300          ' ����ų���
Private Const PROC_WIDTH      As Long = 5295          ' ���Ƿ��ơ����

Private arrStone()   As STONE ' ���صġ�ǽ������Ϣ����
Private arrLevStep() As Long  ' ��ǽ�顱������������
Private arrLevDomW() As Long  ' ��̬���ֵ
Private arrLevMinW() As Long  ' ��С���ֵ
Private arrSound()   As Byte  ' ������Դ����
Private arrSound2()   As Byte  ' ������Դ����
Private arrSound3()   As Byte  ' ������Դ����

Private mlOpeLock As Long     ' �����߼���
Private mlCFXH    As Long     ' ���ڸ߶�������
Private mlCFXW    As Long     ' ���ڿ��������
Private mlClientH As Long     ' ���ڿͻ����߶�
Private mlClientW As Long     ' ���ڿͻ������
Private mlSceneW  As Long     ' ���᳡�����
Private mlSceneX  As Long     ' ���᳡��λ��
Private mlDispX   As Long     ' ��ɫ��ʾλ��
Private mlRoleX   As Long     ' ��ɫ�ڳ����е�λ��
Private mlLevel   As Long     ' �ؿ�����ֵ
Private mlCritiL  As Long     ' ���ٽ��
Private mlCritiR  As Long     ' ���ٽ��
Private mlStoneN  As Long     ' ��ǽ�顱������Ϣ�±���Ͻ�ֵ
Private mlObjPnt  As Long     ' ���صĶ��������Ͻ�
Private mlDispPnt As Long     ' ��ʾ�Ķ����Ͻ�ֵ
Private mlCurrPnt As Long     ' ��ǰ���ڡ�λ�á�
Private mlCurrObj As Long     ' ��ǰ���ĸ���ǽ�顱��
Private mlCurrLen As Long     ' ��ǰ���ų�����ֵ
Private mlKeyLeft As Long     ' ���š����������
Private mlMoveToX As Long     ' ��ɫ�����ߵ�����λ��
Private mlFunFlag As Long     ' �������ܱ�ʶ

Dim Buffer As String * 128
Dim Ret As Long
Dim strAppPath As String  'Ӧ�ó���Ŀ¼

Private Function DispScene() As Boolean
   Dim i&, k&, w&, u As Long

   For i = 0& To mlStoneN
      If (mlSceneX < arrStone(i).cxR) Then Exit For
   Next
   w = i
   k = -1&
   Do
      k = 1& + k
      If (k > mlObjPnt) Then              ' ��Ҫ����ġ�ǽ�顱
         mlObjPnt = 4& + mlObjPnt         ' ���� 4��
         For u = k To mlObjPnt
            Call Load(picStone(u))
         Next
      End If
      u = arrStone(i).cxL - mlSceneX      ' ��߽���ʾ����ֵ
      If (u > mlClientW) Then k = k - 1&: Exit Do
      'picStone(k).Left = u
      'picStone(k).Width = arrStone(i).cw
      Call picStone(k).Move(u, BASE_LINE, arrStone(i).cw)
      If (i = mlStoneN) Then Exit Do      ' ��ʾ�����һ����
      i = 1& + i                          ' ������һ����
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
   
   mlDispX = mlRoleX - mlSceneX           ' ��ɫ��ʾλ��
   Image1.Left = mlDispX - DISP_FIX       ' ����Ӧλ����ʾ��ɫ

   DispScene = (SPACE_END < mlClientW + mlSceneX - arrStone(mlStoneN).cxR)
End Function

Private Sub InitData()        ' �ؿ�Ԥ�������ʼ������
   Dim u  As Long

   u = NUM_LEVEL - 1&         ' Ԥ��8������
   ReDim arrLevStep(u)
   ReDim arrLevDomW(u)
   ReDim arrLevMinW(u)
   ' ��ǽ�顱������������
   arrLevStep(0&) = 20&
   arrLevStep(1&) = 35&       ' +15
   arrLevStep(2&) = 60&       ' +25
   arrLevStep(3&) = 100&      ' +40
   arrLevStep(4&) = 155&      ' +60
   arrLevStep(5&) = 240&      ' +85
   arrLevStep(6&) = 360&      ' +120
   arrLevStep(7&) = 525&      ' +165
    ' ��̬���ֵ
   arrLevDomW(0&) = 1650&
   arrLevDomW(1&) = 1420&
   arrLevDomW(2&) = 1200&
   arrLevDomW(3&) = 1050&
   arrLevDomW(4&) = 900&
   arrLevDomW(5&) = 750&
   arrLevDomW(6&) = 630&
   arrLevDomW(7&) = 525&
   ' ��С���ֵ
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

   u = arrLevStep(mlLevel)    ' ��������
   w = arrLevMinW(mlLevel)    ' ��С���
   n = arrLevDomW(mlLevel)    ' ��̬���
   k = 270&                   ' ��С��ࣨ18���أ�
   v = 900&                   ' ��̬��ࣨ60���أ�
   mlStoneN = u
   ReDim arrStone(u)
   dw = 0.7 * n               ' ��1����ض����ã���� 70%��̬���
   mlRoleX = START_CX + dw \ 2&
   cx = START_CX + dw
   arrStone(0&).cw = dw
   arrStone(0&).cxC = mlRoleX
   arrStone(0&).cxL = START_CX
   arrStone(0&).cxR = cx
   Call Randomize             ' �����
   For i = 1& To u
      dw = k + v * Rnd()      ' �������ࡱ
      cx = cx + dw
      arrStone(i).cxL = cx    ' ��߽�
      dw = w + n * Rnd()      ' �������ȡ�
      arrStone(i).cxC = cx + dw \ 2&      ' �м��
      cx = cx + dw
      arrStone(i).cw = dw
      arrStone(i).cxR = cx    ' �ұ߽�
   Next
   v = n \ 2&
   If (v > dw) Then           ' ���һ��̫խ�����µ���
      arrStone(u).cw = v
      arrStone(u).cxR = cx - dw + v
   End If
   mlSceneW = SPACE_END + cx  ' ���᳡���ܿ��
   mlSceneX = 0&              ' ���᳡����ʼλ��
   mlCurrPnt = 0&             ' ��ʼԭ��
   mlCurrObj = 0&             ' λ�ڡ���1�顱
   labLevel.Caption = "��" & (1& + mlLevel) & "�P"
   labProc.Width = 0&
End Sub

Private Sub ReStart()
   If (mlOpeLock) Then Exit Sub
   lblFen.Caption = SCORE_TEXT & "0"
   labProc.Width = 0          ' �����ȡ���0
   Label1.Visible = False
   Image1.Top = BASE_LINE - Image1.Height
   Image1.Visible = True
   picKey.Visible = False
   mlCurrLen = 0&             ' ���ų����ã�
   mlCurrPnt = 0&             ' ��ʼԭ��
   mlSceneX = 0&              ' ���᳡��λ��
   mlRoleX = arrStone(0&).cxC ' ��ɫ���
   Call DispScene             ' ˢ�³���
   mlOpeLock = vbFalse
End Sub

Private Sub Form_Click()
   If (mlOpeLock) Then Exit Sub           ' ������������
   If Label1.Visible Then Call ReStart    ' ʧ�ܺ����¿�ʼ��
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
   mlObjPnt = 23&             ' ���á�Ԥ���ض�������
   For i = 1& To mlObjPnt:    Load picStone(i):    Next
   Call InitData              ' ��ʼ�����ؿ����ݡ�
   Call LoadLevel             ' ��������
   mlOpeLock = vbFalse
   mlFunFlag = vbTrue
   
   Set Image2.Picture = src1.Picture
   Set Image3.Picture = src1.Picture
   
    strAppPath = App.Path
    If Right(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    Call playBackSound
End Sub

Private Sub Form_Resize()
   Const MIN_CLIENT_H   As Long = 9000    ' 600����
   Const MIN_CLIENT_W   As Long = 13800   ' 920����
   Dim lFlag   As Long
   Dim cw&, ch As Long

   If (-2& = mlOpeLock) Then Exit Sub
   mlOpeLock = -2&            ' ��ֹ����
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
   If (lFlag) Then            ' ������������á������ڴ�С
      DoEvents
      Me.Enabled = True
   End If
   fraScore.Left = cw - PROC_WIDTH - 120
   imgGlass.Top = ch - imgGlass.Height
   imgGlass.Width = cw
   Label1.Width = cw
   imgSky.Width = cw
   lFlag = mlClientH          ' ����֮ǰ�Ĵ��ڿͻ����߶�
   mlClientW = cw             ' ��¼���ڿͻ�����Ⱥ͸߶�
   mlClientH = ch
   mlCritiL = cw \ 4&         ' ���㡰���������ٽ�㡱
   mlCritiR = cw - cw \ 5&
   Call DispScene             ' ˢ����ʾ
   
   If (ch > lFlag) Then       ' ������ڡ��߶����ӡ�ʱ��������ǽ��߶ȡ�
      ch = ch - picStone(0&).Top + 15
      For cw = 0& To mlDispPnt
          picStone(cw).Height = ch
      Next
   End If
   mlOpeLock = vbFalse        ' �������
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
    ElseIf Image2.Tag = 1 Then '��ǰΪ������Ҫ�ص�
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
    ElseIf Image3.Tag = 1 Then '��ǰΪ������Ҫ�ص�
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
   If Label1.Visible Then     ' ��Ϸʧ������
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
      MsgBox "���ǽ��������õ㵱ǰС�����ڵ�ǽ��ѽ", vbExclamation
   End If
End Sub

Private Sub picStone_MouseUp(Index As Integer, Button As Integer, _
                           Shift As Integer, X As Single, Y As Single)
   If (mlOpeLock) Then Exit Sub
   If (ADD_SPEED < mlCurrLen) Then
      Call picKey.Move(mlKeyLeft, BASE_LINE - 120&, mlCurrLen, 135&)
      Image1.Top = picKey.Top - Image1.Height         'С������
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
                  mlRoleX = (mlSceneX + mlMoveToX) \ w            ' �ѡ���ɫ�յ㡱Բ�����������ش�
                  mlRoleX = mlRoleX * w
                  Image1.Left = mlMoveToX - DISP_FIX
                  If (mlRoleX > arrStone(mlStoneN).cxR) Then      ' �Ѿ���������ͷ��
                     w = vbTrue           ' ��ʶ��ʧ�ܡ�
                  Else
                     For w = mlCurrPnt To mlStoneN                ' ����ߵ���һ����
                        If (mlRoleX < arrStone(w).cxR) Then       ' λ����ĳ�顰�ұ߽�֮��
                           If (mlRoleX > arrStone(w).cxL) Then    ' ������ĳ�顰��߽�֮�ҡ�
                              lblFen.Caption = SCORE_TEXT & w
                              labProc.Width = PROC_WIDTH * w / mlStoneN
                              mlCurrPnt = w
                              w = vbFalse ' ��ʶ���ɹ���
                           Else
                              w = vbTrue  ' ��ʶ��ʧ�ܡ�
                           End If
                           Exit For
                        End If
                     Next
                  End If
                  If (w) Then          ' ʧ�ܣ�
                     If Image3.Tag = 1 Then sndPlaySound arrSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
                     mlFunFlag = 2&    ' ���С����䡱��������
                     Timer1.Enabled = True
                     Exit Sub
                  Else
                    If Image3.Tag = 1 Then sndPlaySound arrSound3(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
                  End If
                  If (mlStoneN = mlCurrPnt) Then
                     labProc.Width = PROC_WIDTH
                     MsgBox "��ϲ���أ����㡰ȷ����������һ�ؼ�����ս��", 64&
                     mlOpeLock = vbFalse
                     mlLevel = 1& + mlLevel
                     If (NUM_LEVEL = mlLevel) Then
                        MsgBox "����ţ���Ѿ�ͨ�������йؿ���", 64&
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
                     Timer2.Enabled = True            ' ���������ᴦ��
                  Else
                     mlOpeLock = vbFalse
                  End If
               End If
      Case 2&: Image1.Top = Image1.Top * 1.15
               If (mlClientH < Image1.Top) Then
                  Label1.Caption = "ʧ���ˣ��������ط�����~"
                  Label1.Visible = True
                  Image1.Visible = False
                  Timer1.Enabled = False
                  mlOpeLock = vbFalse
               End If
      
      Case Else:  mlFunFlag = vbTrue
                  Timer1.Enabled = False

   End Select

End Sub

Private Sub Timer2_Timer()    ' �����ᡱ����
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
Public Function GetPlayMode() As String '�õ�����״̬
    Dim Buffer As String * 128
    Dim pos As Integer
    mciSendString "status mp3 mode", Buffer, 128, 0&
    pos = InStr(Buffer, Chr(0))
    GetPlayMode = Left(Buffer, pos - 1)
End Function
'���ű�������
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
