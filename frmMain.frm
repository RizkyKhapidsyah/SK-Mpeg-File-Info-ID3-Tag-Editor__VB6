VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cIDV3 Class Module Example"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Tag"
      Default         =   -1  'True
      Height          =   420
      Left            =   4410
      TabIndex        =   4
      Top             =   6285
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   420
      Left            =   5580
      TabIndex        =   3
      Top             =   6285
      Width           =   1110
   End
   Begin VB.FileListBox File1 
      Height          =   5745
      Left            =   2865
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   420
      Width           =   4005
   End
   Begin VB.DirListBox Dir1 
      Height          =   5715
      Left            =   45
      TabIndex        =   1
      Top             =   420
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6825
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This was made quickly to show how to use the cIDV3 class module.
'I would appreciate anyone who wants to help make it better or fix
'any mistakes I may have to help with this project and email me any
'updates.  Thanks in advance.  Please send any updates to:
'sharmon@vpcusa.com


Public CurrentFile As String

Private Sub cmdEdit_Click()
  LoadEditor
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
  CheckForFile
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  LoadEditor
End Sub

Private Sub CheckForFile()
  If File1.ListCount > 0 Then
    cmdEdit.Enabled = True
  Else
    cmdEdit.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  CheckForFile
End Sub

Private Sub LoadEditor()
Dim strTemp As String
  
  strTemp = Dir1.Path
  If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
  CurrentFile = strTemp & File1.List(File1.ListIndex)
  Load frmID3Edit
  frmID3Edit.Show vbModal
End Sub
