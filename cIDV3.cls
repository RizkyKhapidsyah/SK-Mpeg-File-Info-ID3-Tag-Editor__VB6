VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIDV3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "MP3 Tag Information/Editor class."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Public Enum eGenre
  [Blues] = 0
  [Classic Rock] = 1
  [Country] = 2
  [Dance] = 3
  [Disco] = 4
  [Funk] = 5
  [Grunge] = 6
  [Hip-Hop] = 7
  [Jazz] = 8
  [Metal] = 9
  [New Age] = 10
  [Oldies] = 11
  [Other] = 12
  [Pop] = 13
  [R&B] = 14
  [Rap] = 15
  [Reggae] = 16
  [Rock] = 17
  [Techno] = 18
  [Industrial] = 19
  [Alternative] = 20
  [Ska] = 21
  [Death Metal] = 22
  [Pranks] = 23
  [Soundtrack] = 24
  [Euro-Techno] = 25
  [Ambient] = 26
  [Trip-Hop] = 27
  [Vocal] = 28
  [Jazz+Funk] = 29
  [Fusion] = 30
  [Trance] = 31
  [Classical] = 32
  [Instrumental] = 33
  [Acid] = 34
  [House] = 35
  [Game] = 36
  [Sound Clip] = 37
  [Gospel] = 38
  [Noise] = 39
  [Alt. Rock] = 40
  [Bass] = 41
  [Soul] = 42
  [Punk] = 43
  [Space] = 44
  [Meditative] = 45
  [Instrumental Pop] = 46
  [Instrumental Rock] = 47
  [Ethnic] = 48
  [Gothic] = 49
  [Darkwave] = 50
  [Techno-Industrial] = 51
  [Electronic] = 52
  [Pop-Folk] = 53
  [Eurodance] = 54
  [Dream] = 55
  [Southern Rock] = 56
  [Comedy] = 57
  [Cult] = 58
  [Gangsta Rap] = 59
  [Top 40] = 60
  [Christian Rap] = 61
  [Pop/Funk] = 62
  [Jungle] = 63
  [Native American] = 64
  [Cabaret] = 65
  [New Wave] = 66
  [Phychedelic] = 67
  [Rave] = 68
  [Showtunes] = 69
  [Trailer] = 70
  [Lo-Fi] = 71
  [Tribal] = 72
  [Acid Punk] = 73
  [Acid Jazz] = 74
  [Polka] = 75
  [Retro] = 76
  [Musical] = 77
  [Rock & Roll] = 78
  [Hard Rock] = 79
  [Folk] = 80
  [Folk/Rock] = 81
  [National Folk] = 82
  [Swing] = 83
  [Fast-Fusion] = 84
  [Bebob] = 85
  [Latin] = 86
  [Revival] = 87
  [Celtic] = 88
  [Bluegrass] = 89
  [Avantegarde] = 90
  [Gothic Rock] = 91
  [Progressive Rock] = 92
  [Psychedelic Rock] = 93
  [Symphonic Rock] = 94
  [Slow Rock] = 95
  [Big Band] = 96
  [Chorus] = 97
  [Easy Listening] = 98
  [Acoustic] = 99
  [Humour] = 100
  [Speech] = 101
  [Chanson] = 102
  [Opera] = 103
  [Chamber Music] = 104
  [Sonata] = 105
  [Symphony] = 106
  [Booty Bass] = 107
  [Primus] = 108
  [Porn Groove] = 109
  [Satire] = 110
  [Slow Jam] = 111
  [Club] = 112
  [Tango] = 113
  [Samba] = 114
  [Folklore] = 115
  [Ballad] = 116
  [Power Ballad] = 117
  [Rhythmic Soul] = 118
  [Freestyle] = 119
  [Duet] = 120
  [Punk Rock] = 121
  [Drum Solo] = 122
  [A Capella] = 123
  [Euro-House] = 124
  [Dance Hall] = 125
  [Goa] = 126
  [Drum & Bass] = 127
  [Club-House] = 128
  [Hardcore] = 129
  [Terror] = 130
  [Indie] = 131
  [BritPop] = 132
  [Negerpunk] = 133
  [Polsk Punk] = 134
  [Beat] = 135
  [Christian Gangsta Rap] = 136
  [Heavy Metal] = 137
  [Black Metal] = 138
  [Crossover] = 139
  [Contemporary Christian] = 140
  [Christian Rock] = 141
  [Merengue] = 142
  [Salsa] = 143
  [Trash Metal] = 144
  [Anime] = 145
  [JPop] = 146
  [Synthpop] = 147
End Enum
'----------------------------------------------------

'----------------------------------------------------
Private Type tId3Info
  strTitle As String * 30
  strArtist As String * 30
  strAlbum As String * 30
  strYear  As String * 4
  strComments As String * 30
  byteGenre As Byte
End Type
'----------------------------------------------------

'----------------------------------------------------
Private Type tMPEGInfo
  lFileBytes As Long
  iSeconds As Integer
  iMpegVersion As Integer
  iLayerVersion As Integer
  iKbits As Integer
  lFrames As Long
  lHz As Long
  strMode As String
  strPrivate As String      'Yes/No
  strCRCs As String         'Yes/No
  strCopyrighted As String  'Yes/No
  strOriginal As String     'Yes/No
  strEmphasis As String     'Yes/No
  iChannels As Integer
  strInfoString As String
End Type
'----------------------------------------------------
  
'----------------------------------------------------
'Local variable(s) to hold property value(s)
Private mvarFilename As String
Private id3Info As tId3Info
Private MPEGInfo As tMPEGInfo
'----------------------------------------------------
'
'

'----------------------------------------------------
Private Sub Class_Initialize()
  ClearTagVariables
End Sub
'----------------------------------------------------

Public Property Get InfoString() As String
  InfoString = MPEGInfo.strInfoString
End Property

'----------------------------------------------------
Public Property Get FileBytes() As Long
  FileBytes = MPEGInfo.lFileBytes
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Seconds() As Integer
  Seconds = MPEGInfo.iSeconds
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get MpegVersion() As Integer
  MpegVersion = MPEGInfo.iMpegVersion
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get LayerVersion() As Integer
  LayerVersion = MPEGInfo.iLayerVersion
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Kbits() As Integer
  Kbits = MPEGInfo.iKbits
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Frames() As Long
  Frames = MPEGInfo.lFrames
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Hz() As Long
  Hz = MPEGInfo.lHz
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Mode() As String
  Mode = MPEGInfo.strMode
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get PrivateBit() As String
  PrivateBit = MPEGInfo.strPrivate
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get CRCs() As String
  CRCs = MPEGInfo.strCRCs
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Copyrighted() As String
  Copyrighted = MPEGInfo.strCopyrighted
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Original() As String
  Original = MPEGInfo.strOriginal
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Emphasis() As String
  Emphasis = MPEGInfo.strEmphasis
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Channels() As Integer
  Channels = MPEGInfo.iChannels
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Filename(ByVal vData As String)

  ClearTagVariables
  
  If ExistFile(vData) Then
    mvarFilename = vData
    ReadTag
  Else
    mvarFilename = ""
  End If
  
End Property

Public Property Get Filename() As String
    Filename = mvarFilename
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Comments(ByVal vData As String)
    id3Info.strComments = vData
End Property

Public Property Get Comments() As String
    Comments = RTrim(id3Info.strComments)
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Year(ByVal vData As String)
    id3Info.strYear = vData
End Property

Public Property Get Year() As String
    Year = RTrim(id3Info.strYear)
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Album(ByVal vData As String)
    id3Info.strAlbum = vData
End Property

Public Property Get Album() As String
    Album = RTrim(id3Info.strAlbum)
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Artist(ByVal vData As String)
    id3Info.strArtist = vData
End Property

Public Property Get Artist() As String
    Artist = RTrim(id3Info.strArtist)
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Title(ByVal vData As String)
    id3Info.strTitle = vData
End Property

Public Property Get Title() As String
    Title = RTrim(id3Info.strTitle)
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Genre(ByVal vData As eGenre)
Attribute Genre.VB_Description = "Returns/Sets current genre."
  id3Info.byteGenre = CByte(vData)
End Property

Public Property Get Genre() As eGenre
  Genre = CInt(id3Info.byteGenre)
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Function ClearTag() As Boolean
Dim Blank As String * 127
On Error GoTo clearTagError
  
  If ExistFile(mvarFilename) Then
    Open mvarFilename For Binary Access Write As #1
    Seek #1, LOF(1) - 127
    Put #1, , Blank
    Close #1
    ClearTag = True
  
  Else
    ClearTag = False
  End If
  
  Exit Function
  
clearTagError:
  ClearTag = False

End Function
'----------------------------------------------------

'----------------------------------------------------
Private Sub ReadTag()
On Error Resume Next

  If Not ExistFile(mvarFilename) Then Exit Sub
    
  If HasTag Then
    Open mvarFilename For Binary As #1
    Get #1, FileLen(mvarFilename) - 124, id3Info
    Close #1
  End If

  ReadMPEGInfo

End Sub
'----------------------------------------------------

'----------------------------------------------------
' Portions of this routine were found on the www.
' Part that finds frames is not working....
'
Private Sub ReadMPEGInfo()
Dim ByteArray(4) As Byte, x As Byte, VBR As Boolean
Dim XingH As String * 4, LayerVersion As String, bin As String
Dim Headstart As Integer, i As Integer, z As Integer, FIO As Integer
Dim sMode, sVersion, sLayer, sEmph, sFreq, sTemp, bRate, NoYes, YesNo
On Error Resume Next

  '----------------------------------------------------
  FIO = FreeFile
  Open mvarFilename For Binary Access Read As FIO
  MPEGInfo.lFileBytes = FileLen(mvarFilename)
  '----------------------------------------------------
  
  '----------------------------------------------------
  'If start position <>1 then id3v2 tag exists
  For i = 1 To 5000 'Check up to 5000 bytes for the header
    Get FIO, i, x
    If x = 255 Then 'Header always start with 255 followed by 250 or 251
      Get FIO, i + 1, x
      If x > 249 And x < 252 Then
        Headstart = i 'Set header start position
        Exit For
      End If
    End If
  Next i
  'End check start position for header'''''''''''''
  '----------------------------------------------------

  '----------------------------------------------------
  'Start check for XingHeader
  Get FIO, Headstart + 36, XingH
  If XingH = "Xing" Then
    VBR = True
    For z = 1 To 4
      Get #1, Headstart + 43 + z, ByteArray(z)  'Get framelength to array
    Next z
    MPEGInfo.lFrames = BinToDec(ByteToBit(ByteArray))   'Calculate # of frames
  Else
    VBR = False
  End If
   'End check for XingHeader
  '----------------------------------------------------

  '----------------------------------------------------
  'Start extract the first 4 bytes (32 bits) to an array
   For z = 1 To 4 '
     Get #1, Headstart + z - 1, ByteArray(z)
   Next z
  'Stop extract the first 4 bytes (32 bits) to an array
  '----------------------------------------------------

  Close FIO
  bin = ByteToBit(ByteArray)

  '----------------------------------------------------
  sVersion = Array(25, 0, 2, 1) 'Mpegversion table
  MPEGInfo.iMpegVersion = sVersion(BinToDec(Mid(bin, 12, 2))) 'Get mpegversion from table
  '----------------------------------------------------
  
  '----------------------------------------------------
  sLayer = Array(0, 3, 2, 1) 'Layer table
  MPEGInfo.iLayerVersion = sLayer(BinToDec(Mid(bin, 14, 2))) 'Get layer from table
  '----------------------------------------------------
  
  '----------------------------------------------------
  sMode = Array("Stereo", "Joint Stereo", "Dual Channel", "Single Channel") 'Mode table
  MPEGInfo.strMode = sMode(BinToDec(Mid(bin, 25, 2))) 'Get mode from table
  
  If MPEGInfo.strMode = "Single Channel" Then
    MPEGInfo.iChannels = 1
  Else
    MPEGInfo.iChannels = 2
  End If
  '----------------------------------------------------
  
  '----------------------------------------------------
  sEmph = Array("None", "50/15", "Reserved", "CCITT J 17") 'Emphasis table
  MPEGInfo.strEmphasis = sEmph(BinToDec(Mid(bin, 31, 2))) 'Get emphasis from table

  Select Case MPEGInfo.iMpegVersion 'Look for version to create right table
    Case 1 'For version 1
      sFreq = Array(44100, 48000, 32000)
    Case 2 Or 25 'For version 2 or 2.5
      sFreq = Array(22050, 24000, 16000)
    Case Else
      MPEGInfo.lHz = 0
      Exit Sub
  End Select
  '----------------------------------------------------
  
  MPEGInfo.lHz = sFreq(BinToDec(Mid(bin, 21, 2))) 'Look for frequency in table
  
  '----------------------------------------------------
  LayerVersion = MPEGInfo.iMpegVersion & MPEGInfo.iLayerVersion 'Combine version and layer to string
  Select Case Val(LayerVersion) 'Look for the right Bits table
    Case 11 'Version 1, Layer 1
      bRate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
    Case 12 'V1 L1
      bRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
    Case 13 'V1 L3
      bRate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
    Case 21 Or 251 'V2 L1 and 'V2.5 L1
      bRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
    Case 22 Or 252 Or 23 Or 253 'V2 L2 and 'V2.5 L2 etc...
      bRate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
    Case Else 'If variable Bits
      MPEGInfo.iKbits = 1 'E.G. for variable Bits
      Exit Sub
  End Select
  '----------------------------------------------------
    
  MPEGInfo.iKbits = bRate(BinToDec(Mid(bin, 17, 4)))
  
  '----------------------------------------------------
  NoYes = Array("No", "Yes")
  MPEGInfo.strOriginal = NoYes(Mid(bin, 30, 1))
  MPEGInfo.strCopyrighted = NoYes(Mid(bin, 29, 1))
  MPEGInfo.strPrivate = NoYes(Mid(bin, 24, 1))
  '----------------------------------------------------
  
  '----------------------------------------------------
  YesNo = Array("Yes", "No")
  MPEGInfo.strCRCs = YesNo(Mid(bin, 16, 1))
  '----------------------------------------------------
  
  MPEGInfo.iSeconds = Int(((MPEGInfo.lFileBytes * 8) / MPEGInfo.iKbits) / 1000)
  
  With MPEGInfo 'Formatted info string like Winamp's(tm) MPEG Info box
    .strInfoString = "Size: " & .lFileBytes & " bytes" & vbCrLf _
                    & "Length: " & .iSeconds & " seconds" & vbCrLf _
                    & "MPEG " & Format(.iMpegVersion, "#.0") & " layer " & .iLayerVersion & vbCrLf _
                    & .iKbits & "kbit, " & .lFrames & " frames" & vbCrLf _
                    & .lHz & "hz " & .strMode & vbCrLf _
                    & "Private: " & .strPrivate & vbCrLf _
                    & "CRCs: " & .strCRCs & vbCrLf _
                    & "Copyrighted: " & .strCopyrighted & vbCrLf _
                    & "Original: " & .strOriginal & vbCrLf _
                    & "Emphasis: " & .strEmphasis
  End With

End Sub
'----------------------------------------------------

'----------------------------------------------------
Private Sub ClearTagVariables()
Dim Inbuf As String * 124
  
  With id3Info
    .byteGenre = 12
    .strAlbum = ""
    .strArtist = ""
    .strComments = ""
    .strTitle = ""
    .strYear = ""
  End With

  With MPEGInfo
    .lFileBytes = 0
    .iSeconds = 0
    .iMpegVersion = 0
    .iLayerVersion = 0
    .iKbits = 0
    .lFrames = 0
    .lHz = 0
    .strMode = ""
    .strPrivate = ""
    .strCRCs = ""
    .strCopyrighted = ""
    .strOriginal = ""
    .strEmphasis = ""
    .iChannels = 0
  End With

End Sub
'----------------------------------------------------

'----------------------------------------------------
Public Function HasTag() As Boolean
On Error GoTo hasTagError
Dim tag As String * 3
  
  If Not ExistFile(mvarFilename) Then HasTag = False: Exit Function
  
  Open mvarFilename For Binary As #1
  Get #1, FileLen(mvarFilename) - 127, tag
  Close #1
  
  If UCase(tag) = "TAG" Then
    HasTag = True
  Else
    HasTag = False
  End If

  Exit Function

hasTagError:
  HasTag = False

End Function
'----------------------------------------------------

'----------------------------------------------------
Public Function WriteTag() As Boolean
On Error GoTo writeError
Dim tag As String * 3

  If Not ExistFile(mvarFilename) Then WriteTag = False: Exit Function
  
  tag = "TAG"
  
  Open mvarFilename For Binary Access Write As #1
  Seek #1, FileLen(mvarFilename) - 127
  Put #1, , tag
  Put #1, , id3Info
  Close #1

  WriteTag = True
  
  Exit Function
  
writeError:
  WriteTag = False

End Function
'----------------------------------------------------

'----------------------------------------------------
Public Function GenreText(iGenreID As Integer) As String
Attribute GenreText.VB_Description = "Returns genre item string based on input number from 0 to 147."
Dim aGenre() As Variant

  If iGenreID < 0 Or iGenreID > 147 Then
    GenreText = "Unknown"

  Else
  
  aGenre = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
    "Hip-Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", _
    "Rock", "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", _
    "Soundtrack", "Euro-Techno", "Ambient", "Trip-Hop", "Vocal", "Jazz+Funk", "Fusion", _
    "Trance", "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", _
    "Noise", "Alt. Rock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
    "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial", "Electronic", _
    "Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", _
    "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", "New Wave", _
    "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", "Acid Punk", "Acid Jazz", _
    "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", "Folk", "Folk/Rock", "National Folk", _
    "Swing", "Fast-Fusion", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", _
    "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", _
    "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty Bass", _
    "Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad", _
    "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", _
    "Euro-House", "Dance Hall", "Goa", "Drum & Bass", "Club-House", "Hardcore", "Terror", "Indie", "BritPop", _
    "Negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", _
    "Contemporary Christian", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop")

  GenreText = aGenre(iGenreID)

  End If
  
End Function
'----------------------------------------------------

'----------------------------------------------------
Public Sub FillComboGenre(inCombo As ComboBox, Optional iWhich As Integer)
On Error Resume Next
Dim i As Integer

  If iWhich < 0 Or iWhich > 147 Then iWhich = 0

  inCombo.Clear
  
  For i = 0 To 147
    inCombo.AddItem GenreText(i)
  Next i
  
  inCombo.ListIndex = iWhich
  Exit Sub
  
End Sub
'----------------------------------------------------

'----------------------------------------------------
'Test file existence with error trapping
Private Function ExistFile(ByVal sSpec As String) As Boolean
On Error Resume Next
  
  Call FileLen(sSpec)
  ExistFile = (Err = 0)

End Function
'----------------------------------------------------

'----------------------------------------------------
'This function converts 4*1 byte array to 4*8 bits
Private Function ByteToBit(ByteArray) As String
Dim z As Integer, i As Integer

  ByteToBit = ""
  For z = 1 To 4
    For i = 7 To 0 Step -1
      If Int(ByteArray(z) / (2 ^ i)) = 1 Then
        ByteToBit = ByteToBit & "1"
        ByteArray(z) = ByteArray(z) - (2 ^ i)
      Else
        If ByteToBit <> "" Then
          ByteToBit = ByteToBit & "0"
        End If
      End If
    Next i
  Next z
  
End Function
'----------------------------------------------------

'----------------------------------------------------
'This function converts Binary string to decimal integer
Private Function BinToDec(BinValue As String) As Long
Dim i As Integer

  BinToDec = 0
  For i = 1 To Len(BinValue)
    If Mid(BinValue, i, 1) = 1 Then
      BinToDec = BinToDec + 2 ^ (Len(BinValue) - i)
    End If
  Next i
  
End Function
'----------------------------------------------------
