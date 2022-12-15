VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main 
   Appearance      =   0  '平面
   BackColor       =   &H00C0E0FF&
   Caption         =   "TrafficFlyer"
   ClientHeight    =   4785
   ClientLeft      =   11745
   ClientTop       =   1815
   ClientWidth     =   5730
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5730
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command3 
      Caption         =   "i"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   4320
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      Caption         =   "雜項設定(&O)"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   4095
      Begin VB.ComboBox Combo2 
         Appearance      =   0  '平面
         Height          =   300
         ItemData        =   "Main.frx":10CA
         Left            =   1440
         List            =   "Main.frx":10DD
         TabIndex        =   16
         Text            =   "Chrome"
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "預設瀏覽器(&A)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      Cancel          =   -1  'True
      Caption         =   "離開(&E)"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      Caption         =   "執行(&R)"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BackColor       =   &H00FFC0C0&
      Caption         =   "網站設定(&W)"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '平面
         Height          =   300
         ItemData        =   "Main.frx":1131
         Left            =   960
         List            =   "Main.frx":113E
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox main_url 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   2280
         TabIndex        =   8
         Text            =   "sam0616.pixnet.net"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "網址(&U)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFC0&
      Caption         =   "流量與次數設定(&T)"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5535
      Begin VB.TextBox main_waitsec 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox main_turn 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Text            =   "10"
         Top             =   240
         Width           =   1335
      End
      Begin MSForms.SpinButton SpinButton2 
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   255
         Size            =   "450;450"
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   255
         Size            =   "450;450"
      End
      Begin VB.Label Label5 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "秒"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "等待時間(&W)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "請注意！數值太高可能導致被判定為惡意刷流量！請謹慎設定"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "次數(&R)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  '平面
      Height          =   1920
      Left            =   0
      Picture         =   "Main.frx":115D
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    frmAbout.Show
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

Private Sub Command1_Click()
    If Not ((Combo1.Text = "") Or (main_url.Text = "")) Then
        If Not (main_turn.Text) = "0" And Not (Val(main_turn) < 1) Then
            If Combo1.Text = "ftp://" Then
                J = MsgBox("確定要使用FTP格式開啟嗎？", vbInformation + vbYesNo)
                If J = vbNo Then
                    Exit Sub
                End If
            End If
            Form1.Show
            Me.Hide
        Else
            MsgBox "請給定次數！", 16, "給定次數"
        End If
    Else
        MsgBox "找不到給定的網址，請給定網址！", 16, "給定網址"
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    turnmax = False
    SpinButton2.Value = 10
    SpinButton1.Value = 5
    main_turn.Text = SpinButton2.Value
    main_waitsec = SpinButton1.Value
End Sub


Private Sub main_add_Click()
    If main_url <> "" Then
        main_urls.AddItem main_url.Text
        main_url.Text = ""
    Else
        MsgBox "請輸入文字", 16
    End If
End Sub


Private Sub main_url_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call main_add_Click
    End If
End Sub

Private Sub SpinButton1_Change()
    main_waitsec.Text = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    main_turn.Text = SpinButton2.Value
End Sub
