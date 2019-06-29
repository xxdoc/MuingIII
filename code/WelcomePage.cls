VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WelcomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
    Dim lHot As Long, OpenTime As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Dim MusicMap() As Single, Hot As Long
    
    If BGM.IsPlaying Then
        MusicMap = BGM.GetMusicMap
        For i = 0 To UBound(MusicMap)
            Hot = Hot + MusicMap(i)
        Next
    End If
    
    Hot = Hot / 300
    lHot = lHot + (Hot - lHot) / 3
    If Hot > 60 Then Hot = 60
    If Hot < 0 Then Hot = 0
    If lHot < 0 Then lHot = 0
    If lHot > 60 Then lHot = 60
    
    Page.DrawImage SongP(CurrentSong, BlurBg), 0, 0, Alpha:=1

    Page.DrawImage "welcomelight.png", 0, 0, Alpha:=1 - lHot / 60 * 0.4
    
    If UBound(SongList) = 0 Then lHot = 20
    
    Page.DrawImageEx "dancemusic.png", GW / 2, GH / 2 + Sin(GetTickCount / 100) * lHot - 50, 97 - Sin(GetTickCount / 100) * lHot, 95 - Cos(GetTickCount / 100) * lHot, posOnCenter
    
    Page.Writes "Dance with MUING III", GW / 2, GH - 200, 28, argb(100 + (lHot / 30) * 70, 0, 0, 0), align:=StringAlignmentCenter
    
    If GetTickCount - OpenTime >= 3000 Then
        ECore.NewTransform transFadeOut, 1000, "MainPage"
    End If
    
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
    OpenTime = GetTickCount
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
    'Page.Res.NewImages App.Path & "\assets"
    '创建页面
    ECore.Add Page, "WelcomePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
