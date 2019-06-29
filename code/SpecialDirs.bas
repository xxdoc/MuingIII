Attribute VB_Name = "SpecialDirs"
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Enum SpecialDir
    DESKTOP = &H0& '����
    PROGRAMS = &H2& '����
    MYDOCUMENTS = &H5& '�ҵ��ĵ�
    MYFAVORITES = &H6& '�ղؼ�
    STARTUP = &H7& '����
    RECENT = &H8& '����򿪵��ļ�
    SENDTO = &H9& '����
    STARTMENU = &HB& '��ʼ�˵�
    NETHOOD = &H13& '�����ھ�
    FONTS = &H14& '����
    SHELLNEW = &H15& 'ShellNew
    APPDATA = &H1A& 'ApplicationData
    PRINTHOOD = &H1B& 'PrintHood
    PAGETMP = &H20& '��ҳ��ʱ�ļ�
    COOKIES = &H21& 'CookiesĿ¼
    HISTORY = &H22& '��ʷ
End Enum
Function GetSpecialDir(Dirs As SpecialDir) As String
    Dim sTmp As String * 200, nLength As Long, pidl As Long
    SHGetSpecialFolderLocation 0, Dirs, pidl
    SHGetPathFromIDList pidl, sTmp
    GetSpecialDir = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
End Function
