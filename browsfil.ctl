VERSION 5.00
Begin VB.UserControl FileBrowser 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   KeyPreview      =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   2535
   Begin VB.TextBox txtFolder 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Width           =   1680
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   270
      Left            =   1830
      Picture         =   "browsfil.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   270
   End
End
Attribute VB_Name = "FileBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Enum DialogType
    OpenFile = 0
    SaveFile = 1
End Enum

'Event Declarations:
Event Change() 'MappingInfo=txtFolder,txtFolder,-1,Change

'Default Property Values:
Const m_def_FileDialogType = OpenFile
Const m_def_Filter = "All Files (*.*) | *.*"
Const m_def_InitDir = "C:\"
Const m_def_Title = "Open"
'Property Variables:

Dim m_FileDialogType As DialogType
Dim m_Filter As String
Dim m_InitDir As String
Dim m_Title As String



Private Function SaveDialog(Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = UserControl.hWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
        A = GetSaveFileName(ofn)

        If (A) Then
            SaveDialog = Trim$(ofn.lpstrFile)
        Else
            SaveDialog = ""
        End If

End Function

Private Function OpenDialog(Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = UserControl.hWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        A = GetOpenFileName(ofn)

        If (A) Then
            OpenDialog = Trim$(ofn.lpstrFile)
        Else
            OpenDialog = ""
        End If

End Function
Private Sub cmdBrowse_Click()
    Dim sFile As String
    
    txtFolder.SetFocus 'so we dont see the Focus Rectangle
    'Show the browse for folder dialog
    If m_FileDialogType = OpenFile Then
         sFile = OpenDialog(m_Filter, m_Title, m_InitDir)
    Else
        sFile = SaveDialog(m_Filter, m_Title, m_InitDir)
    End If
    If sFile <> "" Then
        txtFolder.Text = sFile
    End If
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Keypreview is set, so we get all of the keypresses here first.
    'Check for keypresses which should cause the Browse dialog to show
    'Alt and down arrow.
    If KeyCode = vbKeyDown And Shift = 4 Then
        cmdBrowse_Click
    End If
End Sub

Private Sub UserControl_Paint()
    UserControl.Cls
    UserControl.Line (cmdBrowse.Left - Screen.TwipsPerPixelX, 0)-(cmdBrowse.Left - Screen.TwipsPerPixelX, UserControl.ScaleHeight), vb3DLight, BF
End Sub

Private Sub UserControl_Resize()
    'Position the constituent controls
    cmdBrowse.Move UserControl.ScaleWidth - cmdBrowse.Width, 0, cmdBrowse.Width, UserControl.ScaleHeight
    txtFolder.Move Screen.TwipsPerPixelX, Screen.TwipsPerPixelY, UserControl.ScaleWidth - (cmdBrowse.Width + (3 * Screen.TwipsPerPixelX)), UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()
    'Get the tooltip
    txtFolder.ToolTipText = UserControl.Extender.ToolTipText
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtFolder.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtFolder.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtFolder.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtFolder.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtFolder.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtFolder.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtFolder.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtFolder.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    txtFolder.Refresh
End Sub

Private Sub txtFolder_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtFolder.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtFolder.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtFolder.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtFolder.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtFolder.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtFolder.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtFolder.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtFolder.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtFolder.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtFolder.Text() = New_Text
    PropertyChanged "Text"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtFolder.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtFolder.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtFolder.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtFolder.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtFolder.Locked = PropBag.ReadProperty("Locked", False)
    txtFolder.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtFolder.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtFolder.SelText = PropBag.ReadProperty("SelText", "")
    txtFolder.Text = PropBag.ReadProperty("Text", "")
    m_FileDialogType = PropBag.ReadProperty("FileDialogType", m_def_FileDialogType)
    m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
    m_InitDir = PropBag.ReadProperty("InitDir", m_def_InitDir)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", txtFolder.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtFolder.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtFolder.Enabled, True)
    Call PropBag.WriteProperty("Font", txtFolder.Font, Ambient.Font)
    Call PropBag.WriteProperty("Locked", txtFolder.Locked, False)
    Call PropBag.WriteProperty("SelLength", txtFolder.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtFolder.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtFolder.SelText, "")
    Call PropBag.WriteProperty("Text", txtFolder.Text, "")
    Call PropBag.WriteProperty("FileDialogType", m_FileDialogType, m_def_FileDialogType)
    Call PropBag.WriteProperty("Filter", m_Filter, m_def_Filter)
    Call PropBag.WriteProperty("InitDir", m_InitDir, m_def_InitDir)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FileDialogType() As DialogType
    FileDialogType = m_FileDialogType
End Property

Public Property Let FileDialogType(ByVal New_FileDialogType As DialogType)
    m_FileDialogType = New_FileDialogType
    PropertyChanged "FileDialogType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,All Files (*.*) | *.*
Public Property Get Filter() As String
    Filter = m_Filter
End Property

Public Property Let Filter(ByVal New_Filter As String)
    m_Filter = New_Filter
    PropertyChanged "Filter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,C:\
Public Property Get InitDir() As String
    InitDir = m_InitDir
End Property

Public Property Let InitDir(ByVal New_InitDir As String)
    m_InitDir = New_InitDir
    PropertyChanged "InitDir"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Open
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Title As String)
    m_Title = New_Title
    PropertyChanged "Title"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileDialogType = m_def_FileDialogType
    m_Filter = m_def_Filter
    m_InitDir = m_def_InitDir
    m_Title = m_def_Title
End Sub

