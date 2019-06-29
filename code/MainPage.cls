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
'   ҳ�������
    Public Page As GPage
    Dim SY As Long
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    If SY > 0 Then SY = 0
    If SY < -(UBound(SongList) - 1) * 80 Then SY = -(UBound(SongList) - 1) * 80
    
    Page.Clear '��ջ���

    '������
    Page.DrawImage SongP(CurrentSong, BlurBg), 0, 0, Alpha:=1
    Page.DrawImage "shadow.png", 0, 0, Alpha:=1
    Page.DrawImage "listbackground.png", 0, 236, Alpha:=1
    
    '�޸���
    If UBound(SongList) = 0 Then
        Page.Writes "There is not song in your folder .", GW / 2, 256, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
        GoTo SkipSongList
    End If
    
    '�����б�
    Dim i As Long, Y As Long
    For i = 1 To UBound(SongList)
        Y = 496 + (i - 1) * 80 + SY
        If Y >= 496 - 100 And Y <= GH + 100 Then
            If CurrentSong = i Then Page.DrawImage "songfocus.png", 65, 493 + (i - 1) * 80 + SY, Alpha:=1
            Page.DrawImage SongP(i, SCircle), 68, Y, Alpha:=1
            Page.Writes SongList(i).name, 135, 510 + (i - 1) * 80 + SY, 24, argb(255, 0, 0, 0)
            If CurrentSong <> i Then
                Page.DrawImage "songfocus.png", 65, 493 + (i - 1) * 80 + SY, Alpha:=0
                If CheckMouse2 = mMouseUp Then
                    BGM.Create SongList(i).Path & "\music.mp3"
                    BGM.Play
                    CurrentSong = i
                    ECore.NewTransform transFadeIn, 500
                End If
            End If
        End If
    Next
    Page.Paint 0, 0, 496 - 100, GW, 100, argb(255, 255, 255, 255)
    Page.DrawImage "listcover.png", 0, GH + 100, Alpha:=1, Pos:=posOnTop
    Page.DrawImage "listcover.png", 0, 376, Alpha:=1, Direction:=DirVertical
    
    '��ǰ����
    Page.DrawImage SongP(CurrentSong, BCircle), GW / 2, 236, Alpha:=1, Pos:=posOnCenter
    Page.Writes SongList(CurrentSong).name, GW / 2, 376, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    Page.Writes "2333333 SSS", GW / 2, 416, 28, argb(255, 0, 176, 240), align:=StringAlignmentCenter
    
SkipSongList:
    
    '����
    Page.DrawImage "morebutton.png", GW, 0, Alpha:=0.5, Pos:=posOnLeft
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, 300, "SetPage"
    
    '��������
    If IsKeyPress(VK_UP) Then SY = SY + 2
    If IsKeyPress(VK_DOWN) Then SY = SY - 2
    
    If Not BGM.IsPlaying Then BGM.Play
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
    SY = SY + Direction * Depth * 20
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
    SY = -(CurrentSong - 1) * 80
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    'Page.Res.NewImages App.path & "\assets"
    '����ҳ��
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub