VERSION 5.00
Begin VB.Form frmCreateCard
   Caption         =   "�ƿ�"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7740
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3
      Caption         =   "����ԭʼ��"
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton opt2
      Caption         =   "COM2"
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton opt1
      Caption         =   "COM1"
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox lbCreateInfo
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      ItemData        =   "frmCreateCard.frx":0000
      Left            =   600
      List            =   "frmCreateCard.frx":0002
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CommandButton Command2
      Caption         =   "�ء���"
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1
      Caption         =   "�ơ���"
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1
      Caption         =   "������Ա����ϵͳ���ƿ�����"
      BeginProperty Font
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmCreateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iCount As Integer

Private Sub Command1_Click()
    Dim strContinue As Boolean
    Dim strDataGroup As String
    Dim strmsg As String

    Dim akey(6) As Byte
    Dim bkey(6) As Byte
    Dim loadmode, sector As Integer
    Dim Snr As Long
    Dim cardmode As Integer
    Dim databuff32 As String * 32
    Dim address As Integer

    strContinue = True

    MsgBox "����ϵ�һ�ſ���", vbInformation + vbOKOnly, "ϵͳ��ʾ"

    While strContinue
        icdev = -1
        strmsg = ""
        strDataGroup = "00000000000000000000000000000000"
        iCount = iCount + 1


        '��ʼ���˿�
        If icdev < 0 Then
            If opt1.value = True Then
                icdev = rf_init(3, 115200)
            End If
            If opt2.value = True Then
                icdev = rf_init(1, 115200)
            End If
        End If
        If icdev < 0 Then
            MsgBox "�豸��ʼ���˿�ʧ�ܣ�����COM1�˿����������", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
            Exit Sub
        End If

        'װ������
        akey(0) = &HFF
        akey(1) = &HFF
        akey(2) = &HFF
        akey(3) = &HFF
        akey(4) = &HFF
        akey(5) = &HFF

        loadmode = 0
        sector = 1
        st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
        If st <> 0 Then
              MsgBox "װ��A������������ԣ�", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        bkey(0) = &HFF
        bkey(1) = &HFF
        bkey(2) = &HFF
        bkey(3) = &HFF
        bkey(4) = &HFF
        bkey(5) = &HFF
        loadmode = 4
        sector = 1
        st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
        If st <> 0 Then
              MsgBox "װ��B������������ԣ�", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '��ȡ�����
        cardmode = 0
        st = rf_card(ByVal icdev, cardmode, Snr)
        If st <> 0 Then
              MsgBox "��ȡ����ų������鿨Ƭ�Ƿ�źã�", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '��֤����B
        loadmode = 4
        sector = 1
        st = rf_authentication(ByVal icdev, loadmode, sector)
        If st <> 0 Then
              MsgBox "��֤B�������", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '��֤����A
        loadmode = 0
        sector = 1
        st = rf_authentication(ByVal icdev, loadmode, sector)
        If st <> 0 Then
              MsgBox "��֤A�������", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        'д���ݣ�����
        address = 4
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "д��4���ݳ���", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        'д���ݣ�����
        address = 5
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "д��5���ݳ���", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        'д���ݣ�����
        address = 6
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "д��6���ݳ���", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '�޸�����
        bkey(0) = &H0
        bkey(1) = &H1
        bkey(2) = &H2
        bkey(3) = &H3
        bkey(4) = &H4
        bkey(5) = &H5

        akey(0) = &H0
        akey(1) = &H1
        akey(2) = &H2
        akey(3) = &H3
        akey(4) = &H4
        akey(5) = &H5

        st = rf_changeb3(ByVal icdev, 1, akey(0), 0, 0, 0, 1, 0, bkey(0))
        If st <> 0 Then
            MsgBox "�޸�A,B����ʱ����", vbCritical + vbOKOnly, "ϵͳ����"
            lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ�ʧ�ܣ�"
            Call quit
            Exit Sub
        End If

        st = rf_beep(icdev, 3)

        'ȡ���豸
        Call quit

        lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ��ƿ��ɹ���"

        strmsg = MsgBox("�������һ�ſ���", vbQuestion + vbOKCancel, "ϵͳ��ʾ")
        If strmsg = vbOK Then
            strContinue = True
        Else
            strContinue = False
        End If
    Wend
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Dim strContinue As Boolean
    Dim strDataGroup As String
    Dim strmsg As String

    Dim akey(6) As Byte
    Dim bkey(6) As Byte
    Dim loadmode, sector As Integer
    Dim Snr As Long
    Dim cardmode As Integer
    Dim databuff32 As String * 32
    Dim address As Integer

    strContinue = True

    MsgBox "����ϵ�һ�ſ���", vbInformation + vbOKOnly, "ϵͳ��ʾ"

    While strContinue
        icdev = -1
        strmsg = ""
        strDataGroup = "00000000000000000000000000000000"
        iCount = iCount + 1


        '��ʼ���˿�
        If icdev < 0 Then
            If opt1.value = True Then
                icdev = rf_init(3, 115200)
            End If
            If opt2.value = True Then
                icdev = rf_init(0, 115200)
            End If
        End If
        If icdev < 0 Then
            MsgBox "�豸��ʼ���˿�ʧ�ܣ�����COM1�˿����������", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
            Exit Sub
        End If

        'װ������

        bkey(0) = &H0
        bkey(1) = &H1
        bkey(2) = &H2
        bkey(3) = &H3
        bkey(4) = &H4
        bkey(5) = &H5

        loadmode = 4
        sector = 1
        st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
        If st <> 0 Then
              MsgBox "װ��B������������ԣ�", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '��ȡ�����
        cardmode = 0
        st = rf_card(ByVal icdev, cardmode, Snr)
        If st <> 0 Then
              MsgBox "��ȡ����ų������鿨Ƭ�Ƿ�źã�", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '��֤����B
        loadmode = 4
        sector = 1
        st = rf_authentication(ByVal icdev, loadmode, sector)
        If st <> 0 Then
              MsgBox "��֤B�������", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
              Call quit
              Exit Sub
        End If

'        '��֤����A
'        loadmode = 0
'        sector = 1
'        st = rf_authentication(ByVal icdev, loadmode, sector)
'        If st <> 0 Then
'              MsgBox "��֤A�������", vbCritical + vbOKOnly, "ϵͳ��Ϣ"
'              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
'              Call quit
'              Exit Sub
'        End If

        'д���ݣ�����
        address = 4
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "д��4���ݳ���", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        'д���ݣ�����
        address = 5
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "д��5���ݳ���", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        'д���ݣ�����
        address = 6
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "д��6���ݳ���", vbCritical + vbOKOnly, "ϵͳ����"
              lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
              Call quit
              Exit Sub
        End If

        '�޸�����
        bkey(0) = &HFF
        bkey(1) = &HFF
        bkey(2) = &HFF
        bkey(3) = &HFF
        bkey(4) = &HFF
        bkey(5) = &HFF
        akey(0) = &HFF
        akey(1) = &HFF
        akey(2) = &HFF
        akey(3) = &HFF
        akey(4) = &HFF
        akey(5) = &HFF

        st = rf_changeb3(ByVal icdev, 1, akey(0), 0, 0, 0, 1, 0, bkey(0))
        If st <> 0 Then
            MsgBox "�޸�A,B����ʱ����", vbCritical + vbOKOnly, "ϵͳ����"
            lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭʧ�ܣ�"
            Call quit
            Exit Sub
        End If

        '����
        st = rf_beep(icdev, 3)

        'ȡ���豸
        Call quit

        lbCreateInfo.AddItem vbCrLf & "�� " & Trim(Str$(iCount)) & " �ſ���ԭ�ɹ���"

        strmsg = MsgBox("�������һ�ſ���", vbQuestion + vbOKCancel, "ϵͳ��ʾ")
        If strmsg = vbOK Then
            strContinue = True
        Else
            strContinue = False
        End If
    Wend
End Sub

Private Sub Form_Load()
    opt1.value = True
    iCount = 0
End Sub
