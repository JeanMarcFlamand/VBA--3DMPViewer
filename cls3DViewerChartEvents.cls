VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls3DViewerChartEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
''ref : https://docs.microsoft.com/en-us/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-embedded-charts
''ref https://docs.microsoft.com/en-us/office/vba/excel/concepts/events-worksheetfunctions-shapes/chart-object-events
'' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-get-statement
'

Option Explicit

Public WithEvents ThreeDViewerclass As Chart
Attribute ThreeDViewerclass.VB_VarHelpID = -1

Private Sub ThreeDViewerclass_MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    Dim ElementID As Long
    Dim Arg1 As Long
    Dim Arg2 As Long
    Dim varRange As Variant
    Dim myX, myY As Double
    'yes if the mouse-left is clicked
    If Button = xlPrimaryButton Then
        With ActiveChart
            'Get the coordinates of the selected element by the user
            .GetChartElement x, y, ElementID, Arg1, Arg2
            If ElementID = xlSeries Or ElementID = xlDataLabel Then
                If Arg2 > 0 Then
                    'Show the data
                    DisplayMassProperties Arg2
            
                End If
            End If
        End With
    End If
End Sub

Private Sub DisplayMassProperties(MyRow As Long)
    Dim MyString  As String
    Dim x As String
    Dim y As String
    Dim z As String
    Dim w As String
    Dim code1 As String
    Dim code2 As String
    Dim code3 As String
    Dim code4 As String

    w = ActiveSheet.Cells(MyRow + 1, 3)
    x = ActiveSheet.Cells(MyRow + 1, 4)
    y = ActiveSheet.Cells(MyRow + 1, 5)
    z = ActiveSheet.Cells(MyRow + 1, 6)

    MyString = MyString & "Mass properties found at Row " & MyRow + 1 & Chr(10)
    MyString = MyString & " Part Number = " & ActiveSheet.Cells(MyRow + 1, 2) & Chr(10)
    MyString = MyString & " Unit Weight = " & w & Chr(10)
    MyString = MyString & " X cg = " & x & Chr(10)
    MyString = MyString & " Y cg = " & y & Chr(10)
    MyString = MyString & " Z cg = " & z & Chr(10)


    MsgBox MyString

End Sub

