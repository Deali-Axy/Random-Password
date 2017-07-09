VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Password (By Deali-Axy)"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5490
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Btn_3 
      Caption         =   "copy"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Btn_4 
      Caption         =   "copy"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Btn_2 
      Caption         =   "copy"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Btn_1 
      Caption         =   "copy"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox lbl_info 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   5295
   End
   Begin VB.TextBox Txt_4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Txt_3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Txt_2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.CommandButton Btn_Produce 
      Caption         =   "Produce"
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt_1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lbl_4 
      Caption         =   "00"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lbl_3 
      Caption         =   "00"
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lbl_2 
      Caption         =   "00"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lbl_1 
      Caption         =   "00"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_1_Click()
    Clipboard.SetText Txt_1
    lbl_info = "已经将" & Txt_1 & "复制到剪切板"
End Sub

Private Sub Btn_2_Click()
    Clipboard.SetText Txt_2
    lbl_info = "已经将" & Txt_2 & "复制到剪切板"
End Sub

Private Sub Btn_3_Click()
    Clipboard.SetText Txt_3
    lbl_info = "已经将" & Txt_3 & "复制到剪切板"
End Sub

Private Sub Btn_4_Click()
    Clipboard.SetText Txt_4
    lbl_info = "已经将" & Txt_4 & "复制到剪切板"
End Sub

Private Sub Btn_Produce_Click()
    On Error GoTo Err
    Dim iTmp As Long, strTmp As String
    Randomize
    iTmp = Int(Rnd * (1000000000))
    strTmp = Trim(Str(iTmp))
    Txt_1 = strTmp
    strTmp = Mid(strTmp, Len(strTmp) - 1, 1) & strTmp
    lbl_1 = Len(strTmp)
    strTmp = Mod_QBase64.Base64Encode(StrConv(strTmp, vbFromUnicode))
    Txt_2 = strTmp
    lbl_2 = Len(strTmp)
    strTmp = Mod_QBase64.Base64Encode(StrConv(strTmp, vbFromUnicode))
    Txt_3 = strTmp
    lbl_3 = Len(strTmp)
    strTmp = Mod_QBase64.Base64Encode(StrConv(strTmp, vbFromUnicode))
    Txt_4 = strTmp
    lbl_4 = Len(strTmp)
    Exit Sub
Err:
    MsgBox "[Error]ErrNumber=" & Err.Number & " ErrDescription=" & Err.Description
End Sub

Private Sub Form_Terminate()
    Mod_HookSkin.Detach Me.hWnd
End Sub

Private Sub Txt_1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case 2    '右键
            Clipboard.SetText Txt_1
            lbl_info = "已经将" & Txt_1 & "复制到剪切板"
    End Select
End Sub

Private Sub Txt_2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case 2    '右键
            Clipboard.SetText Txt_2
            lbl_info = "已经将" & Txt_2 & "复制到剪切板"
    End Select
End Sub

Private Sub Txt_3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case 2    '右键
            Clipboard.SetText Txt_3
            lbl_info = "已经将" & Txt_3 & "复制到剪切板"
    End Select
End Sub

Private Sub Txt_4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case 2    '右键
            Clipboard.SetText Txt_4
            lbl_info = "已经将" & Txt_4 & "复制到剪切板"
    End Select
End Sub
