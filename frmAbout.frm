VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  '����
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "����ڪ����ε{��"
   ClientHeight    =   2820
   ClientLeft      =   6180
   ClientTop       =   4005
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1946.414
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   7
      Text            =   "frmAbout.frx":10CA
      Top             =   1440
      Width           =   3855
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "frmAbout.frx":7E2D
      ScaleHeight     =   337.12
      ScaleMode       =   0  '�ϥΪ̦ۭq
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '����
      Cancel          =   -1  'True
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   1305
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Appearance      =   0  '����
      Caption         =   "�t�θ�T(&S)..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   1755
      Width           =   1245
   End
   Begin VB.Label Label3 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "Make By 510208�A�е�Github�@��Star�A���U���~�򰵤U�h�I"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "���v(&S)�G"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "'�Ъ`�N���n�餣�䴩��IP�A�]���нT�{�A���|�]�����n��y�����CSEO�ƦW�өǸo��@��510208�I"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '����u
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   776.495
      Y2              =   776.495
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "���ε{�����D"
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   4125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   786.848
      Y2              =   786.848
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   1605
   End
   Begin VB.Label lblDisclaimer 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "ĵ�i: ..."
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Top             =   2145
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ���U���X�w���ʿﶵ...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ���U���X ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' �H Unicode nul ���������r��
Const REG_DWORD = 4                      ' 32-�줸�ƭ�

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "���� " & App.Title & " By 510208"
    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " By 510208"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ���ձq���U�Ϩ��o�t�θ�T�{�����|\�W��...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ���ձq���U�Ϩ��o�t�θ�T�{�����|...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' �ˬd�w���� 32 �줸�ɮת����O�_�s�b
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���~ - �䤣���ɮ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���~ - �䤣����U����...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "�ثe�L�k���Ѩt�θ�T", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' �j��p�ƾ�
    Dim rc As Long                                          ' �Ǧ^�N�X
    Dim hKey As Long                                        ' �}�Ҫ����U���X������N�X
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ���U���X����ƫ��A
    Dim tmpVal As String                                    ' ���U���X�Ȫ��Ȧs�Ŷ�
    Dim KeyValSize As Long                                  ' ���U���X�ܼƪ��j�p
    '------------------------------------------------------------
    ' �}�� KeyRoot {HKEY_LOCAL_MACHINE...} ���U�����U���X (RegKey)
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' �}�ҵ��U���X
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �B�z���~...
    
    tmpVal = String$(1024, 0)                               ' �t�m�ܼƪŶ�
    KeyValSize = 1024                                       ' �Х��ܼƤj�p
    
    '------------------------------------------------------------
    ' �^�����U���X��...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���o/�إ߾��X��
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �B�z���~
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 �|�[�J�H Null ���������r��...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' ��� Null�A�q�r�ꤤ���X
    Else                                                    ' WinNT ���|�[�J�H Null ���������r��...
        tmpVal = Left(tmpVal, KeyValSize)                   ' �䤣�� Null�A���X�r��
    End If
    '------------------------------------------------------------
    ' �M�w���X�Ȫ��ഫ���A...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' �j�M��ƫ��A...
    Case REG_SZ                                             ' String ���U���X��ƫ��A
        KeyVal = tmpVal                                     ' �ƻs�r���
    Case REG_DWORD                                          ' Double Word ���U���X��ƫ��A
        For i = Len(tmpVal) To 1 Step -1                    ' �ഫ�C�@�Ӧ줸
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' �v�r�إ߭�
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' �N Double Word �ഫ�� String
    End Select
    
    GetKeyValue = True                                      ' �Ǧ^���\���T��
    rc = RegCloseKey(hKey)                                  ' �������U���X
    Exit Function                                           ' ���}
    
GetKeyError:      ' ���~�o�ͫ�M��...
    KeyVal = ""                                             ' �]�w�Ǧ^�Ȭ��Ŧr��
    GetKeyValue = False                                     ' �Ǧ^���Ѫ��T��
    rc = RegCloseKey(hKey)                                  ' �������U���X
End Function

