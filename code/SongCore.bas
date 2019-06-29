Attribute VB_Name = "SongCore"
Public Type SongBlocker
    time As Single
    number As Long
    lastTime As Single
End Type
Public Type SongFollower
    Obj() As SongBlocker
End Type
Public Type SongInfo
    name As String
    Artist As String
    Path As String
    MakerCID As String
    Follower As SongFollower
End Type
Public SongList() As SongInfo
Public SongPath As String
Public SongRes As GResource, CurrentSong As Long
Public WelcomePage As WelcomePage, MainPage As MainPage, SetPage As SetPage
Public BGM As GMusic
Public Enum SongPic
    Bg = 0
    BlurBg = 1
    SCircle = 2
    BCircle = 3
End Enum
Function SongP(i As Long, k As SongPic) As Long
    If UBound(SongList) = 0 Then SongP = 1: Exit Function
    SongP = i * 4 + k - 3
End Function
Sub GetSongList()
    'ȡ�ø����ļ���
    SongPath = GetSpecialDir(MYDOCUMENTS) & "\Muing III\"
    If Dir(SongPath, vbDirectory) = "" Then MkDir GetSpecialDir(MYDOCUMENTS) & "\Muing III\"
    
    '��ʼ�������б�
    ReDim SongList(0)
    '��ʼ����Դ��
    Set SongRes = New GResource
    
    '�ݹ����и���
    Dim f As String
    f = Dir(SongPath, vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then '�ų��ϼ�Ŀ¼
            ReDim Preserve SongList(UBound(SongList) + 1)
            With SongList(UBound(SongList))
                .Path = SongPath & f
                .name = "Kiss me"
            End With
            SongRes.newImage SongPath & f & "\background.png", GW, GH, UBound(SongList) & ".png"
            SongRes.newImage SongPath & f & "\background.png", GW, GH, UBound(SongList) & " blur.png"
            SongRes.ApplyBlurEffect UBound(SongList) & " blur.png", 30, 0
            SongRes.newImage SongPath & f & "\background.png", 54, 54, UBound(SongList) & " circle.png"
            SongRes.ClipCircle UBound(SongList) & " circle.png"
            SongRes.newImage SongPath & f & "\background.png", 251, 251, UBound(SongList) & " Bcircle.png"
            SongRes.ClipCircle UBound(SongList) & " Bcircle.png"
        End If
        f = Dir(, vbDirectory)
        DoEvents
    Loop
    
    '����UI��Դ
    SongRes.NewImages App.Path & "\assets\"
    
    '������Դ��
    Set WelcomePage.Page.Res = SongRes
    Set MainPage.Page.Res = SongRes
    Set SetPage.Page.Res = SongRes
    
    '���ѡ��
    Randomize
    CurrentSong = Int(Rnd * UBound(SongList)) + 1
    If CurrentSong > UBound(SongList) Then CurrentSong = UBound(SongList)
    
    '����BGM������
    Set BGM = New GMusic
    BGM.Create SongList(CurrentSong).Path & "\music.mp3"
    BGM.Play
    If ESave.GetData("volume") <> "" Then BGM.Volume = Val(ESave.GetData("volume"))
End Sub

