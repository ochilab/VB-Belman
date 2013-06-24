Attribute VB_Name = "Module1"
'sndPlaySound用の定数
Public Const SND_SYNC = &H0        '同期再生
Public Const SND_ASYNC = &H1       '非同期再生
Public Const SND_NODEFAULT = &H2   '再生に失敗しても別のサウンドを再生しない
Public Const SND_MEMORY = &H4      'メモリファイル
Public Const SND_LOOP = &H8        'ループ再生
Public Const SND_NOSTOP = &H10     '現在の再生中のサウンドがあれば再生しない



Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
    (lpszSoundName As Any, _
    ByVal uFlags As Long) As Long


Sub main()

Debug.Print "aaa"


End Sub
