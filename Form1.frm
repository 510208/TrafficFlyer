VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TrafficFlyer [���椤]"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "����(&C)"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "���b���椤�I�Ф������s��������"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
    Form1.Hide
    Main.Show
End Sub

Private Sub Form_Load()
    Shell "cmd.exe /c start " & Main.main_urls(Int((100 * Rnd) + 1))
End Sub
