VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   Caption         =   "С����ܡ���ף����Ŀ��ģ�"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11640
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   330
      Top             =   270
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1230
      Left            =   10230
      TabIndex        =   3
      Top             =   6615
      Visible         =   0   'False
      Width           =   840
      URL             =   "G:\Game\BirdClass\game.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1482
      _cy             =   2170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����ҿ�ʼ��Ϸ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   825
      Left            =   3630
      TabIndex        =   2
      Top             =   3780
      Width           =   4140
   End
   Begin VB.Image xia4 
      Height          =   3465
      Left            =   7755
      Picture         =   "Form2.frx":871A
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   945
   End
   Begin VB.Image shang4 
      Height          =   3135
      Left            =   7755
      Picture         =   "Form2.frx":A1E5
      Stretch         =   -1  'True
      Top             =   -135
      Width           =   945
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   420
      Left            =   11220
      TabIndex        =   1
      Top             =   270
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   9075
      Picture         =   "Form2.frx":BBCF
      Stretch         =   -1  'True
      Top             =   -135
      Width           =   1440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   420
      Left            =   10560
      TabIndex        =   0
      Top             =   270
      Width           =   510
   End
   Begin VB.Image xia3 
      Height          =   3465
      Left            =   5775
      Picture         =   "Form2.frx":C96C
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   945
   End
   Begin VB.Image xia2 
      Height          =   3465
      Left            =   3630
      Picture         =   "Form2.frx":E437
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   945
   End
   Begin VB.Image shang3 
      Height          =   3135
      Left            =   5775
      Picture         =   "Form2.frx":FF02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   945
   End
   Begin VB.Image shang2 
      Height          =   3135
      Left            =   3630
      Picture         =   "Form2.frx":118EC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   945
   End
   Begin VB.Image xia1 
      Height          =   3465
      Left            =   1485
      Picture         =   "Form2.frx":132D6
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   1005
   End
   Begin VB.Image shang1 
      Height          =   3195
      Left            =   1500
      Picture         =   "Form2.frx":14DA1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   945
   End
   Begin VB.Image bird 
      Height          =   795
      Left            =   495
      Picture         =   "Form2.frx":1678B
      Stretch         =   -1  'True
      Top             =   3375
      Width           =   840
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim score As Integer


Private Sub Form_Click()
    bird.Top = bird.Top - 800   '������
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc(" ") Then bird.Top = bird.Top - 800  '�����񣬼��̿ո�߶�����
End Sub

Private Sub Image3_Click()

End Sub


Private Sub Form_Load()
    'G:\Game\BirdClass\game.mp3
    'WindowsMediaPlayer1.URL = App.Path & "/G:/Game/BirdClass/game.mp3"
End Sub

Private Sub Label1_Click()
    Label1.Visible = False
    Timer1.Enabled = True
    
End Sub

Private Sub Label1_DblClick()
    bird.Top = bird.Top - 800   '������
End Sub

Private Sub Timer1_Timer()
    '���λ��
    bird.Top = bird.Top + 400
       
    '�����ƶ�
    shang1.Left = shang1.Left - 80
    shang2.Left = shang2.Left - 80
    shang3.Left = shang3.Left - 80
    shang4.Left = shang4.Left - 80
    score = score + 1
    Label2.Caption = score
    
    
    '�����ƶ�
    xia1.Left = shang1.Left
    xia2.Left = shang2.Left
    xia3.Left = shang3.Left
    xia4.Left = shang4.Left
    
    '����䶯
        If shang1.Left + shang1.Width < 0 Then  '������
            shang1.Left = shang4.Left + shang4.Width + 800  '��һ��
                Randomize
                shang1.Height = Int(Rnd * 3000 + 800)   '�ı�߶�
        End If
        
        If shang2.Left + shang2.Width < 0 Then
            shang2.Left = shang1.Left + shang1.Width + 800
                Randomize
                shang2.Height = Int(Rnd * 3000 + 800)
        End If
        
        If shang3.Left + shang3.Width < 0 Then
            shang3.Left = shang2.Left + shang2.Width + 800
                Randomize
                shang3.Height = Int(Rnd * 3000 + 800)
        End If
        
        If shang4.Left + shang4.Width < 0 Then
            shang4.Left = shang3.Left + shang3.Width + 800
                Randomize
                shang4.Height = Int(Rnd * 3000 + 800)
        End If
        
    '����䶯
        xia1.Top = shang1.Height + 3500
        xia2.Top = shang2.Height + 3500
        xia3.Top = shang3.Height + 3500
        xia4.Top = shang4.Height + 3500
        
    '��ײ
        '���泬��
            If bird.Left + bird.Width > shang1.Left And bird.Top < shang1.Height Then
                Timer1.Enabled = False
                
                Dim a1 As String
                a1 = MsgBox("�ǳ��ź���������Ϸ ������Ŷ��" + "���ĵ÷�Ϊ��" + Str(score) + "��", vbExclamation, "�𾴵�������ã�")
                    Form1.Show
                    Form2.Hide
                    
            End If
            
            If bird.Left + bird.Width > shang2.Left And bird.Top < shang2.Height Then
                Timer1.Enabled = False
                
                Dim a2 As String
                a2 = MsgBox("�ǳ��ź���������Ϸ ������Ŷ��" + "���ĵ÷�Ϊ��" + Str(score) + "��", vbExclamation, "�𾴵�������ã�")
                    Form1.Show
                    Form2.Hide
            End If
            
            If bird.Left + bird.Width > shang3.Left And bird.Top < shang3.Height Then
                Timer1.Enabled = False
                
                Dim a3 As String
                a3 = MsgBox("�ǳ��ź���������Ϸ ������Ŷ��" + "���ĵ÷�Ϊ��" + Str(score) + "��", vbExclamation, "�𾴵�������ã�")
                    Form1.Show
                    Form2.Hide
            End If
        '���泬��
            If bird.Left + bird.Width > xia1.Left And bird.Top > xia1.Top Then
            
                Timer1.Enabled = False
                
                Dim b1 As String
                b1 = MsgBox("�ǳ��ź���������Ϸ ������Ŷ��", vbExclamation, "�𾴵�������ã�")
                    Form1.Show
                    Form2.Hide
            End If
            
            If bird.Left + bird.Width > xia2.Left And bird.Top > xia2.Top Then
            
                Timer1.Enabled = False
                
                Dim b2 As String
                b2 = MsgBox("�ǳ��ź���������Ϸ ������Ŷ��", vbExclamation, "�𾴵�������ã�")
                    Form1.Show
                    Form2.Hide
            End If
            
            If bird.Left + bird.Width > xia3.Left And bird.Top > xia3.Top Then
            
                Timer1.Enabled = False
                
                Dim b3 As String
                b3 = MsgBox("�ǳ��ź���������Ϸ ������Ŷ��", vbExclamation, "�𾴵�������ã�")
                    Form1.Show
                    Form2.Hide
            End If
    
End Sub

