VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "�ׂ�܂� Ver. 0.8"
   ClientHeight    =   4500
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8880
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2640
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   2880
   End
   Begin VB.Frame Frame2 
      Caption         =   "�o�ߎ���"
      Height          =   2415
      Left            =   600
      TabIndex        =   10
      Top             =   240
      Width           =   5415
      Begin VB.Label Label20 
         BackStyle       =   0  '����
         Caption         =   "���ݎ����F"
         BeginProperty Font 
            Name            =   "System"
            Size            =   13.5
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   2030
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "System"
            Size            =   13.5
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   2400
         TabIndex        =   11
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   2175
      Begin VB.Label Label19 
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "���\�I���\�莞��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "���O"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "���^����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "�\��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "���\����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "���^�I���\�莞��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "���\�J�n����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Menu files 
      Caption         =   "�t�@�C��"
      Begin VB.Menu reading 
         Caption         =   "�ݒ�ǂݍ���"
         Enabled         =   0   'False
      End
      Begin VB.Menu saving 
         Caption         =   "�ݒ�ۑ�"
         Enabled         =   0   'False
      End
      Begin VB.Menu sysend 
         Caption         =   "�I��"
      End
   End
   Begin VB.Menu view 
      Caption         =   "�\��"
      Begin VB.Menu view1 
         Caption         =   "�o�ߎ��ԕ\��"
         Checked         =   -1  'True
      End
      Begin VB.Menu view2 
         Caption         =   "�c�莞�ԕ\��"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu option 
      Caption         =   "�ݒ�"
      Begin VB.Menu setting 
         Caption         =   "����"
      End
   End
   Begin VB.Menu help 
      Caption         =   "�w���v"
      Begin VB.Menu usage 
         Caption         =   "�g����"
      End
      Begin VB.Menu version 
         Caption         =   "�o�[�W�������"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�\�����ԁi�b�j
Dim ss As Integer
'�\�����ԁi���j
Dim mm As Integer
'���\����
Private presen  As Integer
'���^����
Private qanda As Integer
'�\��
Private check As Integer
Private preTime As Integer
'�I������
Private endTime As Integer

Private qandaFlag As Boolean



Private pro As Integer

Private wRet As Long
Private preSound() As Byte         '�T�E���h�f�[�^
Private presenSound() As Byte         '�T�E���h�f�[�^
Private endSound() As Byte         '�T�E���h�f�[�^




Private Sub Command1_Click()

    If presen < 1 Then
        MsgBox "�ݒ胁�j���[�Ŏ��Ԃ�ݒ肵�Ă�������"
    Else
        Timer1.Enabled = True
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = True
        'Label5.Caption = Format$(TimeSerial(0, 0, Timer - start), "hh:mm:ss")
        'Label5.Caption = Format(Timer, "hh:mm:ss")
        Label5.Caption = Time()
         Label19.Caption = Time() + (0.000011574 * 60 * (presen))
        Label4.Caption = Time() + (0.000011574 * 60 * (presen + qanda))
        ProgressBar1.Min = 0
        ProgressBar1.Max = 60 * (presen + qanda)
        pro = 0
    End If

End Sub

Private Sub Command2_Click()

    ss = 0
    mm = 0
    pro = 0
    DrawMTime
    DrawSTime
    
    

''���\�[�X��荞��
'    pHogeSound = LoadResData("TEST1", "WAVE")
'
''�Đ�
'Dim wRet As Long
 '   wRet = sndPlaySound(preSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)

    
    
End Sub

Private Sub Command3_Click()
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = False
    Timer1.Enabled = False
End Sub

Public Sub showSetting()

    Label12.Caption = presen
    Label13.Caption = check
    Label6.Caption = qanda

End Sub

Private Sub Form_Load()

ss = 0
mm = 0

''���\�[�X��荞��
preSound = LoadResData("bell1", "WAVE")
presenSound = LoadResData("bell2", "WAVE")
endSound = LoadResData("bell3", "WAVE")
'With StatusBar1
'
'         .SimpleText = Time   ' ������\�����܂��B
'         .Style = sbrSimple   ' �p�l�� 1 �̃X�^�C���ɂ��܂��B
'
'   End With

End Sub

Private Sub setting_Click()

    frm_setting.Show 1


End Sub

Private Sub sysend_Click()
    End
    
End Sub

Private Sub Timer1_Timer()
 ss = ss + 1
 DrawSTime
End Sub

Private Sub DrawSTime()
    If ss < 10 And ss >= 0 Then
        Label2.Caption = "0" & ss
    ElseIf ss = 60 Then
        Label2.Caption = "00"
        ss = 0
        mm = mm + 1
        DrawMTime
    Else
        Label2.Caption = ss
    
    End If
    pro = pro + 1
    If pro <= ProgressBar1.Max Then
    ProgressBar1.Value = pro
    End If
End Sub

Private Sub DrawMTime()
    '�\���̏���
    If mm < 10 And mm >= 0 Then
        Label1.Caption = "0" & mm
    Else
        Label1.Caption = mm
    
    End If
        
    If mm = preTime Then
        Debug.Print "�\��"
        
        wRet = sndPlaySound(preSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    ElseIf mm = presen Then
        If qandaFlag = True Then
            wRet = sndPlaySound(presenSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
        Else
              wRet = sndPlaySound(endSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
        End If
    ElseIf mm = endTime And qandaFlag = True Then
        Debug.Print "�I��"
        wRet = sndPlaySound(endSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
        


End Sub



Public Sub setTime(PresenTime As Integer, CheckTime As Integer, qandaTime As Integer)

    presen = PresenTime
    check = CheckTime
    qanda = qandaTime
    If qanda < 1 Then
        qandaFlag = False
    Else
        qandaFlag = True
    End If
    preTime = PresenTime - CheckTime
    endTime = PresenTime + qanda
    showSetting

End Sub

Private Sub Timer2_Timer()

'With StatusBar1
'
'         .SimpleText = Time   ' ������\�����܂��B
'         .Style = sbrSimple   ' �p�l�� 1 �̃X�^�C���ɂ��܂��B
'
'   End With
'
    Label17.Caption = Time
End Sub
