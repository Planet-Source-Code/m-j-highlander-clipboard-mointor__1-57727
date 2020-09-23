VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmClipList 
   Caption         =   " Clipboard Monitor"
   ClientHeight    =   6375
   ClientLeft      =   1695
   ClientTop       =   2850
   ClientWidth     =   9510
   Icon            =   "Clip.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9450
      TabIndex        =   2
      Top             =   0
      Width           =   9510
      Begin ClipMon.CoolButton btnSaveAs 
         Height          =   480
         Left            =   3900
         TabIndex        =   3
         ToolTipText     =   "Save All to Text File"
         Top             =   60
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   ""
         MaskColor       =   16711935
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":08CA
      End
      Begin ClipMon.CoolButton cmdCopyAll 
         Height          =   495
         Left            =   630
         TabIndex        =   4
         ToolTipText     =   "Copy All Items"
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":0B5C
      End
      Begin ClipMon.CoolButton cmdClear 
         Height          =   495
         Left            =   1830
         TabIndex        =   5
         ToolTipText     =   "Clear All"
         Top             =   60
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   873
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":0DEE
      End
      Begin ClipMon.CoolButton cmdRemoveItem 
         Height          =   495
         Left            =   1290
         TabIndex        =   6
         ToolTipText     =   "Remove Selected Item"
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":1080
      End
      Begin ClipMon.CoolButton btnStop 
         Height          =   495
         Left            =   2685
         TabIndex        =   7
         ToolTipText     =   "Stop Monitoring"
         Top             =   60
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   873
         Caption         =   ""
         ForeColor       =   -2147483637
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":1312
      End
      Begin ClipMon.CoolButton cmdCopySelectedItem 
         Height          =   495
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Copy Selected Item"
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":15A4
      End
      Begin ClipMon.CoolButton btnResume 
         Height          =   495
         Left            =   3300
         TabIndex        =   9
         ToolTipText     =   "Resume Monitoring"
         Top             =   60
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   873
         Caption         =   ""
         ForeColor       =   -2147483637
         ShowFocusRect   =   0   'False
         BackPicture     =   "Clip.frx":1836
      End
   End
   Begin VB.TextBox txtPreview 
      BackColor       =   &H8000000F&
      Height          =   2700
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3645
      Width           =   9450
   End
   Begin VB.ListBox lstLocal 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   45
      TabIndex        =   0
      Top             =   945
      Width           =   9390
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   6180
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClipList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function JoinListItems(ByRef lstX As ListBox) As String
Dim sArray() As String
Dim lLower As Long, lUpper As Long, idx As Long

If lstX.ListCount = 0 Then
    JoinListItems = ""

Else
    
    lLower = 0
    lUpper = lstX.ListCount - 1
    ReDim sArray(lLower To lUpper)
    
    For idx = lLower To lUpper
        sArray(idx) = lstX.List(idx)
    Next
    
    JoinListItems = Join(sArray, vbCrLf)

End If


End Function

Private Sub btnResume_Click()
    

    btnStop.Enabled = True
    btnResume.Enabled = False

    HookForm Me
    
    lstLocal.SetFocus

End Sub

Private Sub btnSaveAs_Click()
Dim sText As String
Dim ff As Integer

sText = JoinListItems(lstLocal)

If Trim(sText) = "" Then
    MsgBox "Nothing to Save", vbExclamation, "Oops"
    Exit Sub
End If

dlgOpen.CancelError = False
dlgOpen.Filter = "Text files|*.*|All Files|*.*"
dlgOpen.DefaultExt = "txt"
'dlgOpen.FileName = ""
dlgOpen.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

dlgOpen.ShowSave

If dlgOpen.FileName <> "" Then
    ff = FreeFile
    Open dlgOpen.FileName For Output As #ff
    Print #ff, sText
    Close #ff
End If

lstLocal.SetFocus
    
End Sub
Private Sub btnStop_Click()

    UnHookForm Me
    btnStop.Enabled = False
    btnResume.Enabled = True
    
    lstLocal.SetFocus

End Sub
Private Sub cmdClear_Click()

If lstLocal.ListCount > 0 Then
    lstLocal.Clear
    txtPreview.Text = ""
End If

lstLocal.SetFocus

End Sub
Private Sub cmdCopyAll_Click()

If lstLocal.ListCount > 0 Then

    'Stop Monitoring
    UnHookForm Me
    
    DoEvents
    
    Clipboard.Clear
    Clipboard.SetText JoinListItems(lstLocal)
    
    DoEvents
    
    'Resume
    HookForm Me

End If

lstLocal.SetFocus


End Sub
Private Sub cmdCopySelectedItem_Click()

If lstLocal.ListCount > 0 And lstLocal.ListIndex <> -1 Then

    'Stop Monitoring
    UnHookForm Me
    
    DoEvents
    
    Clipboard.Clear
    Clipboard.SetText lstLocal.List(lstLocal.ListIndex)
    
    DoEvents
    
    'Resume
    HookForm Me

End If

lstLocal.SetFocus

End Sub
Private Sub cmdRemoveItem_Click()

If lstLocal.ListCount > 0 And lstLocal.ListIndex <> -1 Then
    
    lstLocal.RemoveItem lstLocal.ListIndex
    txtPreview.Text = ""

End If

lstLocal.SetFocus

End Sub
Private Sub Form_Activate()
On Error Resume Next

lstLocal.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'If KeyAscii = vbKeyEscape Then
'End If

'If KeyAscii = vbKeyReturn Then
'End If

End Sub
Private Sub Form_Load()

'txtPreview.Text = ""

'Disable Context Menu
 
 'OldTextBoxProc = SetWindowLong( _
       txtPreview.hWnd, GWL_WNDPROC, _
        AddressOf NewTextBoxProc)

    'Subclass this form
    HookForm Me
    'Register this form as a Clipboardviewer
    SetClipboardViewer Me.hWnd
    
    btnResume.Enabled = False

End Sub
Private Sub Form_Resize()
On Error Resume Next

lstLocal.Top = picButtons.Height
lstLocal.Left = 0
lstLocal.Width = ScaleWidth
lstLocal.Height = ScaleHeight \ 2

txtPreview.Top = lstLocal.Height + picButtons.Height
txtPreview.Left = 0
txtPreview.Width = ScaleWidth
txtPreview.Height = ScaleHeight - lstLocal.Height - picButtons.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)

''Return Default Handler
' OldTextBoxProc = SetWindowLong( _
'       txtPreview.hWnd, GWL_WNDPROC, _
'        AddressOf OldTextBoxProc)

    'Unhook the form
    UnHookForm Me

End Sub
Private Sub lstLocal_Click()

If lstLocal.ListCount >= 1 Then
    txtPreview.Text = lstLocal.List(lstLocal.ListIndex)
End If

End Sub

Private Sub lstLocal_GotFocus()

If lstLocal.ListIndex = -1 And lstLocal.ListCount > 0 Then
        lstLocal.ListIndex = 0
End If

End Sub
