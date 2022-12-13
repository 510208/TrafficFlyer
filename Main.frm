VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "TrafficFlyer"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   5730
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame3 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "'動作"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2400
      TabIndex        =   18
      Top             =   5400
      Width           =   3135
      Begin VB.CommandButton Command2 
         Appearance      =   0  '平面
         Cancel          =   -1  'True
         Caption         =   "離開(&E)"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  '平面
         Caption         =   "開始(&S)"
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "流量與次數設定(&T)"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   5535
      Begin VB.TextBox main_waitsec 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  '平面
         Caption         =   "+"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  '平面
         Caption         =   "-"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton main_turnmin 
         Appearance      =   0  '平面
         Caption         =   "-"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton main_turnadd 
         Appearance      =   0  '平面
         Caption         =   "+"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox main_turn 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "秒"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label4 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "等待時間(&W)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "請注意！數值太高可能導致被判定為惡意刷流量！請謹慎設定"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "次數(&R)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "網站設定(&W)"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5535
      Begin VB.CommandButton main_clearLIST 
         Appearance      =   0  '平面
         Caption         =   "全部刪除(&C)"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton main_delthis 
         Appearance      =   0  '平面
         Caption         =   "刪除此項(&D)"
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ListBox main_urls 
         Appearance      =   0  '平面
         Height          =   750
         ItemData        =   "Main.frx":0000
         Left            =   960
         List            =   "Main.frx":0007
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton main_add 
         Appearance      =   0  '平面
         Caption         =   "添加(&A)"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox main_url 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "網址(&U)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  '平面
      Height          =   1920
      Left            =   0
      Picture         =   "Main.frx":0016
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim turnmax As Boolean

Private Sub Command1_Click()
    Form1.Show
    Main.Hide
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    main_waitsec.Text = Val(main_waitsec.Text) - 1
End Sub

Private Sub Command4_Click()
    main_waitsec.Text = Val(main_waitsec.Text) + 1
End Sub

Private Sub main_delthis_Click()
On Error GoTo err
    main_urls.RemoveItem (main_urls.ListIndex)
    Exit Sub
err:
    MsgBox "請選擇選項", 16
End Sub

Private Sub main_clearLIST_Click()
    main_urls.Clear
End Sub

Private Sub Form_Load()
    turnmax = False
    main_urls.Clear
    main_turn.Text = 0
    main_waitsec.Text = 0
End Sub


Private Sub main_add_Click()
    If main_url <> "" Then
        main_urls.AddItem main_url.Text
        main_url.Text = ""
    Else
        MsgBox "請輸入文字", 16
    End If
End Sub

Private Sub main_turnadd_Click()
    main_turn.Text = Val(main_turn.Text) + 1
End Sub

Private Sub main_turnmin_Click()
    main_turn.Text = Val(main_turn.Text) - 1
End Sub

Private Sub main_url_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call main_add_Click
    End If
End Sub
