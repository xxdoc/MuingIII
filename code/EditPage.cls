VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditPage"
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

    '背景
    Page.DrawImage SongP(CurrentSong, BlurBg), 0, 0, Alpha:=1
    Page.Paint 0, 0, 0, GW, 131, argb(140, 255, 255, 255)
    Page.Writes "Touch your numberpad !", GW / 2, 50, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    Page.Paint 0, 0, GH - 131, GW, 132, argb(140, 255, 255, 255)
    
    '按钮
    Page.DrawImage "returnbtn.png", 60, 130 / 2, Alpha:=1, pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, 300, "SetPage"
    
    '进度
    Page.Writes SongList(CurrentSong).name, GW / 2, GH - 115, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    Page.Writes ToTime(BGM.Position) & "/" & ToTime(BGM.length), GW - 47, GH - 107, 28, argb(255, 0, 0, 0), align:=StringAlignmentFar
    Dim pos As Single
    pos = BGM.Position / BGM.length
    Page.ShowSliderBar pos, "songbar.png", 60, GH - 65
    If Int(pos * 100) <> Int(BGM.Position / BGM.length * 100) Then BGM.Position = pos * BGM.length
    
    Page.DrawImage "startbtn.png", 60, GH - 130, Alpha:=IIf(BGM.IsPlaying, 0.3, 1), pos:=posOnCenter
    If (Not BGM.IsPlaying) And CheckMouse2 = mMouseUp Then BGM.Play
    Page.DrawImage "endbtn.png", 130, GH - 130, Alpha:=IIf(BGM.IsPlaying, 1, 0.3), pos:=posOnCenter
    If BGM.IsPlaying And CheckMouse2 = mMouseUp Then BGM.Stops
    Page.DrawImage "setbtn.png", GW - 60, GH - 130, Alpha:=IIf(BGM.IsPlaying, 1, 0.3), pos:=posOnCenter
    
    '游戏显示
    Page.Paint 1, GW / 2 - 148, GH / 2 - 148, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2 - 148, GH / 2, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2, GH / 2 - 148, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2, GH / 2, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2 + 148, GH / 2 + 148, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2 + 148, GH / 2, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2, GH / 2 + 148, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2 - 148, GH / 2 + 148, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
    Page.Paint 1, GW / 2 + 148, GH / 2 - 148, 128, 128, argb(180, 255, 255, 255), pos:=posOnCenter
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
    '导入游戏资源
    'Page.Res.NewImages App.Path & "\assets"
    '创建页面
    ECore.Add Page, "EditPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
