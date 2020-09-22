VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "FTP Example"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   5400
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
   End
   Begin VB.ListBox lstDirectory 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox txtServerName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblDirectory 
      AutoSize        =   -1  'True
      Caption         =   "Directories"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   765
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
    lblDirectory.Caption = ""
    Call ReloadDirectories
End Sub

Private Sub Form_Load()
    txtServerName = GetSetting(App.Title, Me.Name, "ServerName", "")
    txtUsername = GetSetting(App.Title, Me.Name, "Username", "")
    txtPassword = GetSetting(App.Title, Me.Name, "Password", "")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting(App.Title, Me.Name, "ServerName", txtServerName)
    Call SaveSetting(App.Title, Me.Name, "Username", txtUsername)
    Call SaveSetting(App.Title, Me.Name, "Password", txtPassword)
End Sub

Private Sub ReloadDirectories(Optional defaultDirectory As String)
Dim nCount As Integer
Dim fileInfo() As tFtpFile
Dim fileCount As Integer

    lstDirectory.Clear
    lstFiles.Clear
    
    lblDirectory.Caption = defaultDirectory
    lblStatus.Caption = "Status: Connecting to " & txtServerName & "..."
    
    If GetFtpDirectory(txtServerName, txtUsername, txtPassword, defaultDirectory, fileInfo, fileCount) = True Then
        lblStatus.Caption = "Status: Connected"
        If lblDirectory.Caption <> "" Then
            lstDirectory.AddItem ".."
        End If
        For nCount = 1 To fileCount
            With fileInfo(nCount)
                If .isDirectory = True Then
                    lstDirectory.AddItem .fileName & vbTab & .FileSize & vbTab & .LastWriteTime
                Else
                    lstFiles.AddItem .fileName & vbTab & .FileSize & vbTab & .LastWriteTime
                End If
            End With
        Next nCount
    Else
        lblStatus.Caption = "Connection Failed!"
    End If
    
    
End Sub

Private Sub lstDirectory_DblClick()
Dim nCount As Integer

    For nCount = 0 To lstDirectory.ListCount - 1
        If lstDirectory.Selected(nCount) = True Then
            With lstDirectory
                If .List(nCount) = ".." Then
                    Call ReloadDirectories(Mid(lblDirectory.Caption, 1, InStr(1, lblDirectory.Caption, "/") - 1))
                Else
                    Call ReloadDirectories(lblDirectory.Caption & "/" & Mid(.List(nCount), 1, InStr(1, .List(nCount), vbTab) - 1))
                End If
                Exit Sub
            End With
        End If
    Next nCount

End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    MsgBox "a"
End If

End Sub
