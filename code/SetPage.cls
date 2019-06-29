VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SetPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布
    Page.DrawImage SongP(CurrentSong, Bg), 0, 0, Alpha:=1
    
    '垃圾动态频谱
    Dim MusicMap() As Single, Buff As Long, p(65)
    If BGM.IsPlaying Then
        MusicMap = BGM.GetMusicMap
        For i = 0 To 30
            If i > UBound(MusicMap) Then Exit Sub
            p(i * 2) = i * Int(GW / 31)
            p(i * 2 + 1) = GH - 320 - MusicMap(i) / 8
        Next
        p(62) = GW: p(63) = GH - 319: p(64) = 0: p(65) = GH - 319
        Page.PaintPolygon2 argb(255, 255, 255, 255), 0, p
    End If
    
    '背景区
    Page.DrawImage SongP(CurrentSong, BlurBg), 0, 0, cy:=0, ch:=130, Alpha:=1
    Page.Paint 0, 0, 0, GW, 131, argb(100, 255, 255, 255)
    Page.DrawImage SongP(CurrentSong, BlurBg), 0, GH - 320, cy:=GH - 320, ch:=320, Alpha:=1
    Page.Paint 0, 0, GH - 60, GW, 61, argb(255, 255, 255, 255)
    Page.Writes "MUING III (ver.1.0.0) , made by Error 404", GW / 2, GH - 48, 24, argb(200, 0, 0, 0), align:=StringAlignmentCenter
    Page.Paint 0, 0, GH - 320, GW, 49, argb(100, 255, 255, 255)
    Page.Writes SongList(CurrentSong).name, GW / 2, GH - 320 + 8, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    
    '按钮
    Page.DrawImage "returnbtn.png", 60, 130 / 2, Alpha:=1, Pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, 300, "MainPage"
    Page.DrawImage "editbtn.png", GW - 140, 130 / 2, Alpha:=1, Pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, 300, "EditPage"
    Page.DrawImage "delbtn.png", GW - 60, 130 / 2, Alpha:=1, Pos:=posOnCenter
    Page.DrawImage "createbutton.png", GW - 80, GH - 320, Alpha:=1, Pos:=posOnCenter
    
    '设置
    Page.Paint 0, 0, GH - 268, GW, 205, argb(190, 255, 255, 255)
    Page.Writes "Volume", 62, GH - 268 + 50, 28, argb(255, 0, 0, 0)
    Page.Writes Int(BGM.Volume * 100) & "%", GW - 62, GH - 268 + 50, 28, argb(255, 0, 176, 240), align:=StringAlignmentFar
    Dim vo As Single
    vo = BGM.Volume
    If Page.ShowSliderBar(vo, "volumebar.png", 65, GH - 268 + 112) = 3 Then
        ESave.PutData "volume", vo
    End If
    If vo <> BGM.Volume Then BGM.Volume = vo
    
    If Not BGM.IsPlaying Then BGM.Play
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
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
    '创建页面
    ECore.Add Page, "SetPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
