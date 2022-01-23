VERSION 5.00
Begin VB.Form frmID3Edit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MPEG file info box + IDv3 tag editor"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Tag"
      Height          =   345
      Left            =   2205
      TabIndex        =   15
      Top             =   2475
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1140
      TabIndex        =   14
      Top             =   2475
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   345
      Left            =   90
      TabIndex        =   13
      Top             =   2475
      Width           =   975
   End
   Begin VB.TextBox txtComment 
      Height          =   300
      Left            =   855
      MaxLength       =   30
      TabIndex        =   12
      Top             =   2010
      Width           =   2520
   End
   Begin VB.TextBox txtYear 
      Height          =   300
      Left            =   855
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1635
      Width           =   570
   End
   Begin VB.TextBox txtAlbum 
      Height          =   300
      Left            =   855
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1260
      Width           =   2520
   End
   Begin VB.TextBox txtArtist 
      Height          =   300
      Left            =   855
      MaxLength       =   30
      TabIndex        =   4
      Top             =   885
      Width           =   2520
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   855
      MaxLength       =   30
      TabIndex        =   2
      Top             =   510
      Width           =   2520
   End
   Begin VB.Frame Frame1 
      Caption         =   "MPEG Info"
      Height          =   2430
      Left            =   3510
      TabIndex        =   16
      Top             =   435
      Width           =   2445
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.TextBox txtFilename 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   5895
   End
   Begin VB.ComboBox cboGenre 
      Height          =   315
      Left            =   2055
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   435
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   11
      Top             =   2055
      Width           =   660
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   195
      Index           =   3
      Left            =   435
      TabIndex        =   7
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   3
      Top             =   930
      Width           =   345
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   5
      Top             =   1305
      Width           =   435
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   195
      Index           =   0
      Left            =   465
      TabIndex        =   1
      Top             =   555
      Width           =   300
   End
End
Attribute VB_Name = "frmID3Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idv3 As New cIDV3

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  idv3.ClearTag
  Unload Me
End Sub

Private Sub cmdSave_Click()

  With idv3
    .Album = txtAlbum
    .Artist = txtArtist
    .Comments = txtComment
    .Title = txtTitle
    .Year = txtYear
    .Genre = cboGenre.ListIndex
    .WriteTag
  End With

  Unload Me

End Sub

Private Sub Form_Load()
  
  With idv3
    txtFilename = frmMain.CurrentFile
    
    .Filename = txtFilename
    
    If .HasTag Then
      .FillComboGenre cboGenre, .Genre
      txtAlbum = .Album
      txtArtist = .Artist
      txtComment = .Comments
      txtTitle = .Title
      txtYear = .Year
    Else
      .FillComboGenre cboGenre
    End If
    
    lblInfo = .InfoString
  End With

End Sub
