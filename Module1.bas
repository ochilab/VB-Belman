Attribute VB_Name = "Module1"
'sndPlaySound�p�̒萔
Public Const SND_SYNC = &H0        '�����Đ�
Public Const SND_ASYNC = &H1       '�񓯊��Đ�
Public Const SND_NODEFAULT = &H2   '�Đ��Ɏ��s���Ă��ʂ̃T�E���h���Đ����Ȃ�
Public Const SND_MEMORY = &H4      '�������t�@�C��
Public Const SND_LOOP = &H8        '���[�v�Đ�
Public Const SND_NOSTOP = &H10     '���݂̍Đ����̃T�E���h������΍Đ����Ȃ�



Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
    (lpszSoundName As Any, _
    ByVal uFlags As Long) As Long


Sub main()

Debug.Print "aaa"


End Sub
