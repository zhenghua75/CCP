VERSION 5.00
Begin VB.Form frmCreateCard
   Caption         =   "制卡"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7740
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3
      Caption         =   "返回原始卡"
      BeginProperty Font
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "关　闭"
      BeginProperty Font
         Name            =   "宋体"
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
      Caption         =   "制　卡"
      BeginProperty Font
         Name            =   "宋体"
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
      Caption         =   "福保会员管理系统的制卡程序"
      BeginProperty Font
         Name            =   "宋体"
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

    MsgBox "请放上第一张卡！", vbInformation + vbOKOnly, "系统提示"

    While strContinue
        icdev = -1
        strmsg = ""
        strDataGroup = "00000000000000000000000000000000"
        iCount = iCount + 1


        '初始化端口
        If icdev < 0 Then
            If opt1.value = True Then
                icdev = rf_init(3, 115200)
            End If
            If opt2.value = True Then
                icdev = rf_init(1, 115200)
            End If
        End If
        If icdev < 0 Then
            MsgBox "设备初始化端口失败，请检查COM1端口连接情况！", vbCritical + vbOKOnly, "系统信息"
            Exit Sub
        End If

        '装载密码
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
              MsgBox "装载A密码出错，请重试！", vbCritical + vbOKOnly, "系统信息"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
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
              MsgBox "装载B密码出错，请重试！", vbCritical + vbOKOnly, "系统信息"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '获取卡序号
        cardmode = 0
        st = rf_card(ByVal icdev, cardmode, Snr)
        If st <> 0 Then
              MsgBox "获取卡序号出错，请检查卡片是否放好！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '验证密码B
        loadmode = 4
        sector = 1
        st = rf_authentication(ByVal icdev, loadmode, sector)
        If st <> 0 Then
              MsgBox "验证B密码出错！", vbCritical + vbOKOnly, "系统信息"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '验证密码A
        loadmode = 0
        sector = 1
        st = rf_authentication(ByVal icdev, loadmode, sector)
        If st <> 0 Then
              MsgBox "验证A密码出错！", vbCritical + vbOKOnly, "系统信息"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '写数据，卡号
        address = 4
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "写块4数据出错！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '写数据，卡号
        address = 5
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "写块5数据出错！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '写数据，卡号
        address = 6
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "写块6数据出错！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
              Call quit
              Exit Sub
        End If

        '修改密码
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
            MsgBox "修改A,B密码时出错！", vbCritical + vbOKOnly, "系统错误"
            lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡失败！"
            Call quit
            Exit Sub
        End If

        st = rf_beep(icdev, 3)

        '取消设备
        Call quit

        lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡制卡成功！"

        strmsg = MsgBox("请插入下一张卡！", vbQuestion + vbOKCancel, "系统提示")
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

    MsgBox "请放上第一张卡！", vbInformation + vbOKOnly, "系统提示"

    While strContinue
        icdev = -1
        strmsg = ""
        strDataGroup = "00000000000000000000000000000000"
        iCount = iCount + 1


        '初始化端口
        If icdev < 0 Then
            If opt1.value = True Then
                icdev = rf_init(3, 115200)
            End If
            If opt2.value = True Then
                icdev = rf_init(0, 115200)
            End If
        End If
        If icdev < 0 Then
            MsgBox "设备初始化端口失败，请检查COM1端口连接情况！", vbCritical + vbOKOnly, "系统信息"
            Exit Sub
        End If

        '装载密码

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
              MsgBox "装载B密码出错，请重试！", vbCritical + vbOKOnly, "系统信息"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
              Call quit
              Exit Sub
        End If

        '获取卡序号
        cardmode = 0
        st = rf_card(ByVal icdev, cardmode, Snr)
        If st <> 0 Then
              MsgBox "获取卡序号出错，请检查卡片是否放好！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
              Call quit
              Exit Sub
        End If

        '验证密码B
        loadmode = 4
        sector = 1
        st = rf_authentication(ByVal icdev, loadmode, sector)
        If st <> 0 Then
              MsgBox "验证B密码出错！", vbCritical + vbOKOnly, "系统信息"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
              Call quit
              Exit Sub
        End If

'        '验证密码A
'        loadmode = 0
'        sector = 1
'        st = rf_authentication(ByVal icdev, loadmode, sector)
'        If st <> 0 Then
'              MsgBox "验证A密码出错！", vbCritical + vbOKOnly, "系统信息"
'              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
'              Call quit
'              Exit Sub
'        End If

        '写数据，卡号
        address = 4
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "写块4数据出错！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
              Call quit
              Exit Sub
        End If

        '写数据，卡号
        address = 5
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "写块5数据出错！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
              Call quit
              Exit Sub
        End If

        '写数据，卡号
        address = 6
        databuff32 = strDataGroup
        st = rf_write_hex(ByVal icdev, address, ByVal databuff32)
        If st <> 0 Then
              MsgBox "写块6数据出错！", vbCritical + vbOKOnly, "系统错误"
              lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
              Call quit
              Exit Sub
        End If

        '修改密码
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
            MsgBox "修改A,B密码时出错！", vbCritical + vbOKOnly, "系统错误"
            lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原失败！"
            Call quit
            Exit Sub
        End If

        '蜂鸣
        st = rf_beep(icdev, 3)

        '取消设备
        Call quit

        lbCreateInfo.AddItem vbCrLf & "第 " & Trim(Str$(iCount)) & " 张卡返原成功！"

        strmsg = MsgBox("请插入下一张卡！", vbQuestion + vbOKCancel, "系统提示")
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
