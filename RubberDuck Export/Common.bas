Attribute VB_Name = "Common"
Option Explicit

'https://msdn.microsoft.com/en-us/library/windows/desktop/bb776426(v=vs.85).aspx


Private Declare Function ShellExecute _
                         Lib "shell32.dll" Alias "ShellExecuteA" ( _
                         ByVal hWnd As Long, _
                         ByVal Operation As String, _
                         ByVal Filename As String, _
                         Optional ByVal Parameters As String, _
                         Optional ByVal Directory As String, _
                         Optional ByVal WindowStyle As Long = vbMinimizedFocus _
                         ) As Long

Public Sub OpenUrl(Url As String)

    Dim UrlOK As Long
    UrlOK = ShellExecute(0, "Open", Url)

End Sub


