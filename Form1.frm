VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '平面
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  '沒有框線
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   15855
   ClientTop       =   9645
   ClientWidth     =   3270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  '平面
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "執行中"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Me.Hide
    Main.Show
End Sub

Private Sub Form_Load()
start:
    Dim turn As Long
    Dim maxturn As Long
    Dim urlwz
    Dim msgbox_num
    turn = 0
    maxturn = 5
    urlwz = Main.Combo1.Text & Main.main_url.Text
    For i = 1 To Val(Main.main_turn.Text)
        Sleep Val(Main.main_waitsec.Text) * 1000
        Call ShellExecute(Me.hwnd, "open", urlwz, "", "", vbNormalFocus)
        If turn > maxturn Then
            Sleep 5000
            Select Case Main.Combo2.ListIndex
                Case 1
                    Shell "cmd.exe /c" & "taskkill /f /t /im chrome.exe ", vbNormalFocus
                Case 2
                    Shell "cmd.exe /c" & "taskkill /f /t /im firefox.exe", vbNormalFocus
                Case 3
                    Shell "cmd.exe /c" & "taskkill /f /t /im msedge.exe", vbNormalFocus
                Case 4
                    Shell "cmd.exe /c" & "taskkill /f /t /im iexplore.exe ", "vbNormalFocus"
                End Select
            maxturn = maxturn + 5
        End If
        turn = turn + 1
    Next i
    msgbox_turn = MsgBox("執行結束，" & vbCrLf & "是否要重新開始？", vbInformation + vbYesNo)
    If msgbox_turn = vbYes Then
        GoTo start
    Else
        Call Command1_Click
    End If
End Sub
