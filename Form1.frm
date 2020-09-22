VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESOFT MiniFTP Manager"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   5760
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6555
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2295
      Width           =   6615
      Begin VB.Label Bilgi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Se&ttings..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2655
      TabIndex        =   20
      Top             =   1725
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Index           =   1
      Left            =   135
      TabIndex        =   13
      Top             =   2235
      Width           =   6345
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Text            =   "\db\"
         Top             =   1350
         Width           =   4650
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   975
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Top             =   225
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default <DIR>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   19
         Top             =   1410
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   17
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   16
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FTP Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   200
         TabIndex        =   15
         Top             =   285
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   885
      ScaleWidth      =   6615
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6615
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   795
      Width           =   6720
   End
   Begin VB.Frame Frame1 
      Height          =   2145
      Index           =   0
      Left            =   3270
      TabIndex        =   9
      Top             =   795
      Width           =   30
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Form1.frx":13A30
      Left            =   3465
      List            =   "Form1.frx":13A32
      TabIndex        =   1
      Top             =   1245
      Width           =   3030
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Receive file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3990
      TabIndex        =   3
      Top             =   1725
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   2
      Top             =   1725
      Width           =   1260
   End
   Begin FtpExample.FileBrowser Fbr 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   1260
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receive file (Select)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3450
      TabIndex        =   14
      Top             =   975
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send file (Selected)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   975
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Url As String
Public UName As String
Public UPass As String
Public Dizin As String



Private Sub Command1_Click()
If Trim(Fbr.Text) = "" Then
   MsgBox "Please select sending file ...", vbExclamation + vbOKOnly, " Error!"
   Exit Sub
End If

Url = Trim(Text1(0).Text)
UName = Trim(Text1(1).Text)
UPass = Trim(Text1(2).Text)
Dizin = Trim(Text1(3).Text): If Right(Dizin, 1) <> "\" Then Text1(3).Text = Text1(3).Text & "\": Dizin = Text1(3).Text

Bilgi.Caption = "(" & Trim(DosyaIsimDondur(Fbr.Text)) & ") sending... [ Wait... ]"
X = PutFtpFile(Url, UName, UPass, Dizin & DosyaIsimDondur(Fbr.Text), Fbr.Text, FTP_TRANSFER_TYPE_BINARY)
Select Case X
    Case 0
    Bilgi.Caption = "Error ! File don't sended.. Please try again."
    Case 1
    SaveSetting "MYFtp", "Rex", "Server", Trim(Text1(0).Text)
    SaveSetting "MYFtp", "Rex", "Username", Trim(Text1(1).Text)
    SaveSetting "MYFtp", "Rex", "Password", Trim(Text1(2).Text)
    SaveSetting "MYFtp", "Rex", "Dizin", Trim(Text1(3).Text)
    Bilgi.Caption = "(" & Trim(DosyaIsimDondur(Fbr.Text)) & ") file is sended ! ..."
End Select
MsgBox Bilgi.Caption, vbExclamation + vbOKOnly, " Yeah..."

End Sub

Private Sub Command2_Click()
If Trim(Combo1.Text) = "" Then
   MsgBox "Please write your receive file name ...", vbExclamation + vbOKOnly, " Error!"
   Exit Sub
End If

Url = Trim(Text1(0).Text)
UName = Trim(Text1(1).Text)
UPass = Trim(Text1(2).Text)
Dizin = Trim(Text1(3).Text): If Right(Dizin, 1) <> "\" Then Text1(3).Text = Text1(3).Text & "\": Dizin = Text1(3).Text

Bilgi.Caption = "(" & Trim(Combo1.Text) & ") çekiliyor... [Wait...]"
X = GetFtpFile(CStr(Url), CStr(UName), CStr(UPass), Dizin & Combo1.Text, Combo1.Text)
Select Case X
    Case 0
    Bilgi.Caption = "Error ! File don't receiving.. Please try again."
    Case 1
    SaveSetting "MYFtp", "Rex", "Server", Trim(Text1(0).Text)
    SaveSetting "MYFtp", "Rex", "Username", Trim(Text1(1).Text)
    SaveSetting "MYFtp", "Rex", "Password", Trim(Text1(2).Text)
    SaveSetting "MYFtp", "Rex", "Dizin", Trim(Text1(3).Text)
    Bilgi.Caption = "(" & Trim(Combo1.Text) & ") file is received ..."
End Select
MsgBox Bilgi.Caption, vbExclamation + vbOKOnly, " Yeah..."

End Sub

Private Sub Command3_Click()
If Me.Height > 3030 Then Me.Height = 3030 Else Me.Height = 4995
DoEvents
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Command4_Click()
Url = Trim(Text1(0).Text)
UName = Trim(Text1(1).Text)
UPass = Trim(Text1(2).Text)
Dizin = Trim(Text1(3).Text): If Right(Dizin, 1) <> "\" Then Text1(3).Text = Text1(3).Text & "\": Dizin = Text1(3).Text

Dim A(1024) As tFtpFile
Dim b As Integer

GetFtpDirectory Url, UName, UPass, "\db\", A(), b, ""

For i = 1 To UBound(A)
    Select Case Trim(A(i).fileName)
    Case ".", "..", ""
    Case Else
    z = z & A(i).fileName & "-" & Format(A(i).FileSize, "#,##0") & vbLf
    End Select
Next i
    MsgBox z

End Sub

Private Sub Form_Load()
Me.Caption = "CESOFT FTPManager v" & Trim(CStr(App.Major)) & "." & Trim(CStr(App.Minor)) & "." & Trim(CStr(App.Revision))
Me.Height = 3030
Fbr.InitDir = App.Path

    Text1(0).Text = GetSetting("MYFtp", "Rex", "Server")
    Text1(1).Text = GetSetting("MYFtp", "Rex", "UserName")
    Text1(2).Text = GetSetting("MYFtp", "Rex", "Password")
    Text1(3).Text = GetSetting("MYFtp", "Rex", "Dizin")
        
    If Text1(0).Text = "" Then
        Command3_Click
    End If

End Sub

Public Function DosyaIsimDondur(Gelen As String) As String
Dim X
X = Split(Fbr.Text, "\"): DosyaIsimDondur = X(UBound(X))
End Function
