VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "С�����"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   Picture         =   "birdfrm.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   9795
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "��Ϸ����"
      Height          =   690
      Left            =   3465
      TabIndex        =   2
      Top             =   3375
      Width           =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������Ϸ"
      Height          =   825
      Left            =   3465
      TabIndex        =   1
      Top             =   2160
      Width           =   2160
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ʼ��Ϸ"
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
    rule = MsgBox("ͨ�������̿ո���굥����˫������С������������ϰ�����Ϸ������", vbQuestion, "��Ϸ����˵����")
End Sub

Private Sub Form_Load()

End Sub
