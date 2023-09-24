VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "小鸟快跑"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   Picture         =   "birdfrm.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   9795
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "游戏规则"
      Height          =   690
      Left            =   3465
      TabIndex        =   2
      Top             =   3375
      Width           =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "结束游戏"
      Height          =   825
      Left            =   3465
      TabIndex        =   1
      Top             =   2160
      Width           =   2160
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "开始游戏"
      Height          =   825
      Left            =   3465
      TabIndex        =   0
      Top             =   945
      Width           =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.Visible = False
    Form2.Visible = True
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Dim rule As String
    rule = MsgBox("通过按键盘空格、鼠标单击或双击控制小鸟。如果触碰到障碍则游戏结束！", vbQuestion, "游戏规则说明：")
End Sub

Private Sub Form_Load()

End Sub
