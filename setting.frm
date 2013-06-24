VERSION 5.00
Begin VB.Form frm_setting 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾀﾞｲｱﾛｸﾞ ｷｬﾌﾟｼｮﾝ"
   ClientHeight    =   3810
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Alignment       =   1  '右揃え
      Height          =   270
      Left            =   240
      TabIndex        =   9
      Text            =   "0"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '右揃え
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Text            =   "0"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "（発表時間終了の時を基準に）"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "分前"
      Height          =   180
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label6 
      Caption         =   "予鈴時間"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   1200
      TabIndex        =   8
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      Caption         =   "質疑時間"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "（質疑時間は含めないでください）"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "発表時間"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

    Unload Me
    

End Sub

Private Sub OKButton_Click()
    If Text1.Text < 1 Then
        MsgBox "発表時間を1分以上設定してください"
    Else
        If Text1.Text - Text3.Text < 1 Then
            MsgBox "発表時間と予鈴のバランスがおかしいです"
        Else
            If Text2.Text < 1 Then
                MsgBox "質疑時間のない発表になります"
            End If
            Form1.setTime Text1.Text, Text3.Text, Text2.Text
            Unload Me
        End If
    End If
End Sub
