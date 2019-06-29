VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
    Dim IconOn As Boolean
    Dim SY As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    If SY > 0 Then SY = 0
    If SY < -(UBound(SongList) - 1) * 80 Then SY = -(UBound(SongList) - 1) * 80
    
    Page.Clear '清空画布

    '背景区
    Page.DrawImage SongP(CurrentSong, BlurBg), 0, 0, Alpha:=1
    Page.DrawImage "shadow.png", 0, 0, Alpha:=1
    Page.DrawImage "listbackground.png", 0, 236, Alpha:=1
    
    '无歌曲
    If UBound(SongList) = 0 Then
        Page.Writes "There is not song in your folder .", GW / 2, 256, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
        GoTo SkipSongList
    End If
    
    '歌曲列表
    Dim I As Long, Y As Long
    For I = 1 To UBound(SongList)
        Y = 496 + (I - 1) * 80 + SY
        If Y >= 496 - 100 And Y <= GH + 100 Then
            If CurrentSong = I Then Page.DrawImage "songfocus.png", 65, 493 + (I - 1) * 80 + SY, Alpha:=1
            Page.DrawImage SongP(I, SCircle), 68, Y, Alpha:=1
            Page.Writes SongList(I).name, 135, 510 + (I - 1) * 80 + SY, 24, argb(255, 0, 0, 0)
            If CurrentSong <> I Then
                Page.DrawImage "songfocus.png", 65, 493 + (I - 1) * 80 + SY, Alpha:=0
                If CheckMouse2 = mMouseUp Then
                    BGM.Create SongList(I).Path & "\music.mp3"
                    BGM.Play
                    CurrentSong = I
                    ECore.NewTransform transFadeIn, 500
                End If
            End If
        End If
    Next
    Page.Paint 0, 0, 496 - 100, GW, 100, argb(255, 255, 255, 255)
    Page.DrawImage "listcover.png", 0, GH + 100, Alpha:=1, Pos:=posOnTop
    Page.DrawImage "listcover.png", 0, 376, Alpha:=1, Direction:=DirVertical
    
    '当前歌曲
    Dim SON As Boolean
    Page.DrawImage SongP(CurrentSong, BCircle), GW / 2, 236, Alpha:=1, Pos:=posOnCenter
    SON = (CheckMouse2 > 0)
    If SON Then Page.DrawImage "playbtn.png", GW / 2, 236, Alpha:=1, Pos:=posOnCenter
    If IconOn <> SON Then IconOn = SON: ECore.NewTransform
    
    Page.Writes SongList(CurrentSong).name, GW / 2, 376, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    Page.Writes "2333333 SSS", GW / 2, 416, 28, argb(255, 0, 176, 240), align:=StringAlignmentCenter
    
SkipSongList:
    
    '更多
    Page.DrawImage "morebutton.png", GW, 0, Alpha:=0.5, Pos:=posOnLeft
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, 300, "SetPage"
    
    '键盘输入
    If IsKeyPress(VK_UP) Then SY = SY + 2
    If IsKeyPress(VK_DOWN) Then SY = SY - 2
    
    If Not BGM.IsPlaying Then BGM.Play
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
    SY = SY + Direction * Depth * 20
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    SY = -(CurrentSong - 1) * 80
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    'Page.Res.NewImages App.path & "\assets"
    '创建页面
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
