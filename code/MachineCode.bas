Attribute VB_Name = "MachineCode"
'====================================================================================
'   Machine Code
'   Made: Error 404(1361778219)
'   Version: 1.0(09.3.9)
'   Describe: ���ڻ�ȡ�豸������
'   Note: ��Ҫ����MD5.cls
'====================================================================================
    '����ϵͳ��Ϣ
    '<Driver=Ŀ���豸>
    Function RetSystem(Driver As String) As String
        Dim o As Object, o2 As Object
        
        Set o = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_" & Driver)
        For Each o2 In o
            RetSystem = RetSystem & o2.Caption & vbCrLf
        Next
        
        Set o = Nothing
    End Function
    '�����豸������
    Function GetMachineCode() As String
        Dim r As String
        r = RetSystem("SoundDevice") & vbCrLf & _
            RetSystem("Processor") & vbCrLf & _
            RetSystem("DiskDrive") & vbCrLf & _
            RetSystem("MotherboardDevice") & vbCrLf & _
            RetSystem("VideoController") & vbCrLf & _
            RetSystem("Keyboard") & vbCrLf & _
            RetSystem("PointingDevice")
        Set o = New MD5
        GetMachineCode = o.Md5_String_Calc(r)
    End Function
'====================================================================================
