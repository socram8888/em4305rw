VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "EM4305 reader"
   ClientHeight    =   4050
   ClientLeft      =   40
   ClientTop       =   380
   ClientWidth     =   8590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   8590
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox BellOnSuccess 
      Caption         =   "Bell on success"
      Height          =   250
      Left            =   3600
      TabIndex        =   83
      Top             =   120
      Value           =   1  'Checked
      Width           =   1690
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tag configuration"
      Height          =   2770
      Left            =   5160
      TabIndex        =   67
      Top             =   720
      Width           =   3130
      Begin VB.CheckBox DisableAllowed 
         Caption         =   "Allow disable"
         Height          =   250
         Left            =   1560
         TabIndex        =   82
         Top             =   1440
         Width           =   1450
      End
      Begin VB.CommandButton CfgEncode 
         Caption         =   "Encode"
         Height          =   370
         Left            =   1680
         TabIndex        =   81
         Top             =   2280
         Width           =   1330
      End
      Begin VB.CommandButton CfgDecode 
         Caption         =   "Decode"
         Height          =   370
         Left            =   120
         TabIndex        =   80
         Top             =   2280
         Width           =   1330
      End
      Begin VB.CheckBox PigeonMode 
         Caption         =   "Pigeon mode"
         Height          =   250
         Left            =   1560
         TabIndex        =   79
         Top             =   1920
         Width           =   1450
      End
      Begin VB.CheckBox ReaderTalkFirst 
         Caption         =   "Reader talk first"
         Height          =   250
         Left            =   1560
         TabIndex        =   78
         Top             =   1680
         Width           =   1450
      End
      Begin VB.CheckBox WriteLogin 
         Caption         =   "Write login"
         Height          =   250
         Left            =   1560
         TabIndex        =   77
         Top             =   1200
         Width           =   1450
      End
      Begin VB.CheckBox ReadLogin 
         Caption         =   "Read login"
         Height          =   250
         Left            =   1560
         TabIndex        =   76
         Top             =   960
         Width           =   1450
      End
      Begin VB.ComboBox LWRCombo 
         Height          =   280
         ItemData        =   "main.frx":0000
         Left            =   1560
         List            =   "main.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   600
         Width           =   1330
      End
      Begin VB.ComboBox DelayedOnCombo 
         Height          =   280
         ItemData        =   "main.frx":0047
         Left            =   120
         List            =   "main.frx":0054
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1800
         Width           =   1330
      End
      Begin VB.ComboBox DataEncodingCombo 
         Height          =   280
         ItemData        =   "main.frx":006E
         Left            =   120
         List            =   "main.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1200
         Width           =   1330
      End
      Begin VB.ComboBox DataRateCombo 
         Height          =   280
         ItemData        =   "main.frx":0092
         Left            =   120
         List            =   "main.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   600
         Width           =   1330
      End
      Begin VB.Label Label21 
         Caption         =   "Last def. read word"
         Height          =   250
         Left            =   1560
         TabIndex        =   74
         Top             =   360
         Width           =   1450
      End
      Begin VB.Label Label20 
         Caption         =   "Delayed ON"
         Height          =   250
         Left            =   120
         TabIndex        =   72
         Top             =   1560
         Width           =   1210
      End
      Begin VB.Label Label19 
         Caption         =   "Encoding"
         Height          =   250
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   1330
      End
      Begin VB.Label Label18 
         Caption         =   "Data rate"
         Height          =   250
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   970
      End
   End
   Begin VB.TextBox Password 
      Height          =   300
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   66
      Text            =   "00000000"
      Top             =   120
      Width           =   850
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Enabled         =   0   'False
      Height          =   250
      Index           =   14
      Left            =   4560
      TabIndex        =   64
      Top             =   3000
      Width           =   250
   End
   Begin VB.CommandButton Command16 
      Caption         =   "W"
      Enabled         =   0   'False
      Height          =   250
      Left            =   4560
      TabIndex        =   63
      Top             =   3360
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   13
      Left            =   4560
      TabIndex        =   62
      Top             =   2640
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   12
      Left            =   4560
      TabIndex        =   61
      Top             =   2280
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   11
      Left            =   4560
      TabIndex        =   60
      Top             =   1920
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   10
      Left            =   4560
      TabIndex        =   59
      Top             =   1560
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   9
      Left            =   4560
      TabIndex        =   58
      Top             =   1200
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   8
      Left            =   4560
      TabIndex        =   57
      Top             =   840
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   15
      Left            =   4200
      TabIndex        =   56
      Top             =   3360
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   14
      Left            =   4200
      TabIndex        =   55
      Top             =   3000
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   13
      Left            =   4200
      TabIndex        =   54
      Top             =   2640
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   12
      Left            =   4200
      TabIndex        =   53
      Top             =   2280
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   11
      Left            =   4200
      TabIndex        =   52
      Top             =   1920
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   10
      Left            =   4200
      TabIndex        =   51
      Top             =   1560
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   9
      Left            =   4200
      TabIndex        =   50
      Top             =   1200
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   8
      Left            =   4200
      TabIndex        =   49
      Top             =   840
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   7
      Left            =   2160
      TabIndex        =   48
      Top             =   3360
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   6
      Left            =   2160
      TabIndex        =   47
      Top             =   3000
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   5
      Left            =   2160
      TabIndex        =   46
      Top             =   2640
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   4
      Left            =   2160
      TabIndex        =   45
      Top             =   2280
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   3
      Left            =   2160
      TabIndex        =   44
      Top             =   1920
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   2
      Left            =   2160
      TabIndex        =   43
      Top             =   1560
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Enabled         =   0   'False
      Height          =   250
      Index           =   1
      Left            =   2160
      TabIndex        =   42
      Top             =   1200
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   7
      Left            =   1800
      TabIndex        =   41
      Top             =   3360
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   6
      Left            =   1800
      TabIndex        =   40
      Top             =   3000
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   5
      Left            =   1800
      TabIndex        =   39
      Top             =   2640
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   4
      Left            =   1800
      TabIndex        =   38
      Top             =   2280
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   3
      Left            =   1800
      TabIndex        =   37
      Top             =   1920
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Enabled         =   0   'False
      Height          =   250
      Index           =   2
      Left            =   1800
      TabIndex        =   36
      Top             =   1560
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   1
      Left            =   1800
      TabIndex        =   35
      Top             =   1200
      Width           =   250
   End
   Begin VB.CommandButton WriteSingle 
      Caption         =   "W"
      Height          =   250
      Index           =   0
      Left            =   2160
      TabIndex        =   34
      Top             =   840
      Width           =   250
   End
   Begin VB.CommandButton ReadSingle 
      Caption         =   "R"
      Height          =   250
      Index           =   0
      Left            =   1800
      TabIndex        =   33
      Top             =   840
      Width           =   250
   End
   Begin VB.CommandButton ReadAll 
      Caption         =   "Read all"
      Height          =   370
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   15
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   31
      Top             =   3360
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   14
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   30
      Top             =   3000
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   13
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   29
      Top             =   2640
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   12
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   28
      Top             =   2280
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   11
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   27
      Top             =   1920
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   10
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   21
      Top             =   1560
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   9
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   20
      Top             =   1200
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   8
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   19
      Top             =   840
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   7
      Left            =   720
      MaxLength       =   8
      TabIndex        =   14
      Top             =   3360
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   6
      Left            =   720
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3000
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   5
      Left            =   720
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2640
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   4
      Left            =   720
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2280
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   3
      Left            =   720
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1920
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   2
      Left            =   720
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1560
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   1
      Left            =   720
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1200
      Width           =   970
   End
   Begin VB.TextBox pageData 
      Height          =   300
      Index           =   0
      Left            =   720
      MaxLength       =   8
      TabIndex        =   1
      Top             =   840
      Width           =   970
   End
   Begin VB.Label Label1 
      Caption         =   "W0"
      Height          =   250
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   370
   End
   Begin VB.Label Label5 
      Caption         =   "W4"
      Height          =   250
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   370
   End
   Begin VB.Label Label17 
      Caption         =   "Password"
      Height          =   250
      Left            =   1680
      TabIndex        =   65
      Top             =   120
      Width           =   850
   End
   Begin VB.Label Label16 
      Caption         =   "W15"
      Height          =   250
      Left            =   2640
      TabIndex        =   26
      Top             =   3360
      Width           =   490
   End
   Begin VB.Label Label15 
      Caption         =   "W14"
      Height          =   250
      Left            =   2640
      TabIndex        =   25
      Top             =   3000
      Width           =   490
   End
   Begin VB.Label Label14 
      Caption         =   "W13"
      Height          =   250
      Left            =   2640
      TabIndex        =   24
      Top             =   2640
      Width           =   490
   End
   Begin VB.Label Label13 
      Caption         =   "W12"
      Height          =   250
      Left            =   2640
      TabIndex        =   23
      Top             =   2280
      Width           =   490
   End
   Begin VB.Label Label12 
      Caption         =   "W11"
      Height          =   250
      Left            =   2640
      TabIndex        =   22
      Top             =   1920
      Width           =   490
   End
   Begin VB.Label Label11 
      Caption         =   "W10"
      Height          =   250
      Left            =   2640
      TabIndex        =   18
      Top             =   1560
      Width           =   490
   End
   Begin VB.Label Label10 
      Caption         =   "W9"
      Height          =   250
      Left            =   2640
      TabIndex        =   17
      Top             =   1200
      Width           =   490
   End
   Begin VB.Label Label9 
      Caption         =   "W8"
      Height          =   250
      Left            =   2640
      TabIndex        =   16
      Top             =   840
      Width           =   490
   End
   Begin VB.Label Label8 
      Caption         =   "W7"
      Height          =   250
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   490
   End
   Begin VB.Label Label7 
      Caption         =   "W6"
      Height          =   250
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   490
   End
   Begin VB.Label Label6 
      Caption         =   "W5"
      Height          =   250
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   370
   End
   Begin VB.Label Label4 
      Caption         =   "W3"
      Height          =   250
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   370
   End
   Begin VB.Label Label3 
      Caption         =   "W2"
      Height          =   250
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   370
   End
   Begin VB.Label Label2 
      Caption         =   "W1"
      Height          =   250
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   370
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RF_Init Lib "SRF32.dll" Alias "s_init" (ByVal Unk0 As Long, ByVal Unk1 As Long, ByVal Unk2 As Long, ByVal Unk3 As Long) As Long
Private Declare Function RF_Stop Lib "SRF32.dll" Alias "s_exit" (ByVal handle As Long)

Private Declare Function RF_EM4305_Login Lib "SRF32.dll" Alias "s_em4305_login" (ByVal handle As Long, ByVal DataRate As Long, ByRef bytes As Byte) As Integer
Private Declare Function RF_EM4305_Read Lib "SRF32.dll" Alias "s_em4305_readWord" (ByVal handle As Long, ByVal DataRate As Long, ByVal page As Long, ByRef bytes As Byte) As Integer
Private Declare Function RF_EM4305_Write Lib "SRF32.dll" Alias "s_em4305_writeWord" (ByVal handle As Long, ByVal DataRate As Long, ByVal page As Long, ByRef bytes As Byte) As Integer

Private Declare Function RF_Field_On Lib "SRF32.dll" Alias "s_el8265a_rf_start" (ByVal handle As Long) As Integer
Private Declare Function RF_Field_Off Lib "SRF32.dll" Alias "s_el8265a_rf_stop" (ByVal handle As Long) As Integer

Private Declare Function RF_Bell Lib "SRF32.dll" Alias "s_bell" (ByVal handle As Long, ByVal time As Long) As Integer

Dim rfHandle As Long

Private Sub bell(ByVal handle As Long)
    If BellOnSuccess.Value Then
        RF_Bell handle, 1
    End If
End Sub

Function BytesToHex(ByRef block() As Byte) As String
    Dim byteStr As String

    byteStr = ""
    For i = LBound(block) To UBound(block)
        If block(i) < 16 Then
            byteStr = byteStr & "0" & Hex$(block(i))
        Else
            byteStr = byteStr & Hex$(block(i))
        End If
    Next

    BytesToHex = byteStr
End Function

Function HexToBytes(ByRef str As String, ByRef block() As Byte) As Boolean
    On Error GoTo fail
    HexToBytes = False

    If Len(str) <> (UBound(block) + 1) * 2 Then
        Exit Function
    End If

    For i = 0 To UBound(block)
        block(i) = CByte("&H" & Mid$(str, i * 2 + 1, 2))
    Next

    HexToBytes = True
fail:
End Function

Private Sub CfgDecode_Click()
    Dim blockBytes(3) As Byte

    If Not HexToBytes(pageData(4).Text, blockBytes) Then
        MsgBox ("Cannot parse configuration")
        Exit Sub
    End If

    DataRate = blockBytes(0) \ 4
    If DataRate = &H30 Then
        DataRateCombo.ListIndex = 0
    ElseIf DataRate = &H38 Then
        DataRateCombo.ListIndex = 1
    ElseIf DataRate = &H3C Then
        DataRateCombo.ListIndex = 2
    ElseIf DataRate = &H32 Then
        DataRateCombo.ListIndex = 3
    ElseIf DataRate = &H3E Then
        DataRateCombo.ListIndex = 4
    Else
        MsgBox ("Unknown data rate " & CStr(DataRate))
    End If

    encoder = (blockBytes(0) And 3) * 4 Or (blockBytes(1) \ 64)
    If encoder = 8 Then
        DataEncodingCombo.ListIndex = 0
    ElseIf encoder = 4 Then
        DataEncodingCombo.ListIndex = 1
    Else
        MsgBox ("Unknown encoding " & CStr(encoder))
    End If

    DelayedOn = (blockBytes(1) \ 4) And 3
    If DelayedOn = 3 Then
        DelayedOn = 0
    End If
    DelayedOnCombo.ListIndex = DelayedOn

    ReadLogin.Value = (blockBytes(2) \ 32) And 1
    WriteLogin.Value = (blockBytes(2) \ 8) And 1
    DisableAllowed.Value = blockBytes(2) And 1
    ReaderTalkFirst.Value = blockBytes(3) \ 128
    PigeonMode.Value = (blockBytes(3) \ 32) And 1
End Sub

Private Sub CfgEncode_Click()
    Dim blockBytes(3) As Byte

    If DataRateCombo.ListIndex = 0 Then
        DataRate = &H30
    ElseIf DataRateCombo.ListIndex = 1 Then
        DataRate = &H38
    ElseIf DataRateCombo.ListIndex = 2 Then
        DataRate = &H3C
    ElseIf DataRateCombo.ListIndex = 3 Then
        DataRate = &H32
    ElseIf DataRateCombo.ListIndex = 4 Then
        DataRate = &H3E
    Else
        Debug.Assert result = False
    End If
    blockBytes(0) = blockBytes(0) Or DataRate * 4

    If DataEncodingCombo.ListIndex = 0 Then
        encoder = 8
    ElseIf DataEncodingCombo.ListIndex = 1 Then
        encoder = 4
    Else
        Debug.Assert result = False
    End If
    blockBytes(0) = blockBytes(0) Or (encoder \ 4)
    blockBytes(1) = blockBytes(1) Or ((encoder And 3) * 64)

    blockBytes(1) = blockBytes(1) Or DelayedOnCombo.ListIndex * 4

    If ReadLogin.Value Then
        blockBytes(2) = blockBytes(2) Or 32
    End If
    If WriteLogin.Value Then
        blockBytes(2) = blockBytes(2) Or 8
    End If
    If DisableAllowed.Value Then
        blockBytes(2) = blockBytes(2) Or 1
    End If
    If ReaderTalkFirst.Value Then
        blockBytes(3) = blockBytes(3) Or 128
    End If
    If PigeonMode.Value Then
        blockBytes(3) = blockBytes(3) Or 32
    End If

    pageData(4).Text = BytesToHex(blockBytes)
End Sub

Private Sub Form_Load()
    ' Change directory so it loads DLL correctly
    ChDir (App.Path)

    rfHandle = RF_Init(0, 0, 0, 0)

    DataRateCombo.ListIndex = 0
    DataEncodingCombo.ListIndex = 0
    DelayedOnCombo.ListIndex = 0
    LWRCombo.ListIndex = 0
End Sub

Private Sub ReadAll_Click()
    Dim passBytes(3) As Byte
    Dim blockBytes(3) As Byte

    If Not HexToBytes(Password.Text, passBytes) Then
        MsgBox ("Password must be an hexadecimal 32-bit string")
        Exit Sub
    End If

    loginRet = RF_EM4305_Login(rfHandle, 0, passBytes(0))
    If loginRet <> 1 Then
        MsgBox ("Login failed. Code " & CStr(loginRet))
    End If

    For i = 0 To 15
        If i = 2 Then
            If loginRet = 1 Then
                pageData(2).Text = BytesToHex(passBytes)
            Else
                pageData(2).Text = "????????"
            End If
        Else
            For try = 1 To 5
                retVal = RF_EM4305_Read(rfHandle, 0, 256 + i, blockBytes(0))
                If retVal = 4 Then
                    Exit For
                End If
            Next

            If retVal < 0 Then
                pageData(i).Text = "????????"
            Else
                pageData(i).Text = BytesToHex(blockBytes)
            End If
        End If
    Next

    bell rfHandle
End Sub

Private Sub ReadSingle_Click(Index As Integer)
    Dim blockBytes(3) As Byte

    If Not HexToBytes(Password.Text, blockBytes) Then
        MsgBox ("Password must be an hexadecimal 32-bit string")
        Exit Sub
    End If

    retVal = RF_EM4305_Login(rfHandle, 0, blockBytes(0))
    If retVal <> 1 Then
        MsgBox ("Login failed. Code " & CStr(retVal))
    End If

    retVal = RF_EM4305_Read(rfHandle, 64, Index, blockBytes(0))

    If retVal <> 4 Then
        pageData(Index).Text = "????????"
        MsgBox ("Unable to read. Code " & CStr(retVal))
    Else
        pageData(Index).Text = BytesToHex(blockBytes)
        bell rfHandle
    End If
End Sub

Private Sub WriteSingle_Click(Index As Integer)
    Dim blockBytes(3) As Byte

    If Not HexToBytes(Password.Text, blockBytes) Then
        MsgBox ("Password must be an hexadecimal 32-bit string")
        Exit Sub
    End If

    retVal = RF_EM4305_Login(rfHandle, 0, blockBytes(0))
    If retVal <> 1 Then
        MsgBox ("Login failed. Code " & CStr(retVal))
    End If

    If Not HexToBytes(pageData(Index).Text, blockBytes) Then
        MsgBox ("Invalid hex contents")
        Exit Sub
    End If

    retVal = RF_EM4305_Write(rfHandle, 0, Index, blockBytes(0))
    If retVal = 1 Then
        bell rfHandle
    Else
        MsgBox ("Unable to write. Code " & CStr(retVal))
    End If
End Sub
