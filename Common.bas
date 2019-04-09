Attribute VB_Name = "Common"
Option Explicit

'https://msdn.microsoft.com/en-us/library/windows/desktop/bb776426(v=vs.85).aspx

'https://msdn.microsoft.com/en-us/library/windows/desktop/bb776426(v=vs.85).aspx
#If VBA7 Then                                    '64-bit
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal Operation As String, _
    ByVal Filename As String, _
    Optional ByVal Parameters As String, _
    Optional ByVal Directory As String, _
    Optional ByVal WindowStyle As Long = vbMinimizedFocus _
    ) As Long
#Else                                            '32-bit
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                             ByVal hWnd As Long, _
                             ByVal Operation As String, _
                             ByVal Filename As String, _
                             Optional ByVal Parameters As String, _
                             Optional ByVal Directory As String, _
                             Optional ByVal WindowStyle As Long = vbMinimizedFocus _
                             ) As Long
#End If

Public Sub OpenUrl(Url As String)

    Dim UrlOK As Long
    UrlOK = ShellExecute(0, "Open", Url)

End Sub

Sub PrepareForm()
    Dim nrAlpha As Name
    Dim nrBeta  As Name
    Dim nrGamma As Name
    
    With ThisWorkbook
        ' Workbook level Name ranges
        Set nrAlpha = .Names("AlphaDeg")
        Set nrBeta = .Names("BetaDeg")
        Set nrGamma = .Names("GammaDeg")
        'Set nrAlpha = .Worksheets("Support").Names("AlphaDeg")
    End With
    
    frmRotationAxis.ShowInit nrAlpha, nrBeta, nrGamma
    
End Sub


