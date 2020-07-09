VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'定数
Const headerRow As Long = 5
Const logRow As Long = 6

'メンバ変数
Private bk As Workbook
Private st As Worksheet
Private lastColNum As Long
Private lastRowNum As Long
Private pvtCache As PivotCache


'コンストラクタ/デストラクタ
Private Sub Class_Initialize()
  Debug.Print "TMSLogHandler Initialized"

End Sub

Private Sub Class_Terminate()
  Debug.Print "TMSLogHandler Terminated."

End Sub


'getter/setter
Property Get GetWorkbook() As Workbook
  Set GetWorkbook = bk

End Property

Property Get GetWorksheet() As Worksheet
  Set GetWorksheet = st

End Property


'Private
Private Function OpenWorkbook(opPath As String) As Workbook
    'ファイルの生存確認
  Dim retBook As Workbook
On Error GoTo Catch
  Set retBook = Workbooks.Open(FileName:=opPath, _
                                        UpdateLinks:=0, _
                                        ReadOnly:=False)
  Set OpenWorkbook = retBook
  Exit Function
Catch:
  Err.Raise Number:=999, Description:="ファイルを開くことができませんでした。"
  Set OpenWorkbook = Nothing

End Function

Private Function CloseWorkbook(clsBook As Workbook) As Workbook
  Application.DisplayAlerts = False
  clsBook.Close
  Set CloseWorkbook = Nothing
  Application.DisplayAlerts = True

End Function


'Public
Public Function SetWorkbook(path As String)
  Set bk = OpenWorkbook(path)
  Set st = bk.Sheets(1)
  lastColNum = st.Cells(headerRow, 1).End(xlToRight).Column
  lastRowNum = st.Cells(headerRow, 1).End(xlDown).Row

End Function

Public Function MergeLogs(withFilePath As String)
  Dim bk2 As Workbook
  Dim st2 As Worksheet
  Set bk2 = OpenWorkbook(withFilePath)
  Set st2 = bk2.Sheets(1)

  Dim lastRowNum2 As Long
  Dim lastColNum2 As Long
  Dim RowCount2 As Long
  lastColNum2 = st2.Cells(headerRow, 1).End(xlToRight).Column
  lastRowNum2 = st2.Cells(headerRow, 1).End(xlDown).Row
  RowCount2 = lastRowNum2 - logRow + 1

  Dim tmpArr As Variant
  tmpArr = st2.Range(st2.Cells(logRow, 1), st2.Cells(lastRowNum2, lastColNum2))
  
  st.Cells(lastRowNum + 1, 1).Value = tmpArr
  st.Range(st.Cells(lastRowNum + 1, 1), st.Cells(lastRowNum + RowCount2, lastColNum2)) = tmpArr
  lastRowNum = lastRowNum + RowCount2
  Set bk2 = CloseWorkbook(bk2)

End Function

Public Function CreateDataTable(TableName As String)
  st.ListObjects.Add(xlSrcRange, st.Range(st.Cells(headerRow, 1), st.Cells(lastRowNum, lastColNum)), , xlYes).Name = TableName

End Function

Public Function CreatePivotCache(srcTableName As String)
  Set pvtCache = bk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcTableName, Version:=6)

End Function

Public Function CreatePivotGraph(graphName As String, RowFields() As String, ColFields() As String, DataFields() As PivotGraphDataField)
  Dim pvtChart As Shape
  Dim pvtSheet As Worksheet
  Dim pvtTable As pivotTable

  Set pvtSheet = bk.Sheets.Add()
  pvtSheet.Name = graphName & "P"
  Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=pvtSheet.Range("A3"), TableName:=graphName, DefaultVersion:=6)
  Set pvtChart = pvtSheet.Shapes.AddChart2(201, xlLine)
  pvtChart.Chart.SetSourceData Source:=pvtTable.DataBodyRange

  Dim counter As Long
  counter = 1
  If Not (Not RowFields) = -1 Then
    Dim r As Variant
    For Each r In RowFields
      With pvtChart.Chart.PivotLayout.pivotTable.PivotFields(r)
        .Orientation = xlRowField
        .Position = counter
      End With
      counter = counter + 1
    Next r
  End If

  counter = 1
  If Not (Not ColFields) = -1 Then
    Dim c As Variant
    For Each c In ColFields
      With pvtChart.Chart.PivotLayout.pivotTable.PivotFields(c)
        .Orientation = xlColumnField
        .Position = counter
      End With
      counter = counter + 1
    Next c
  End If

  If Not (Not DataFields) = -1 Then
    Dim d As Variant
    For Each d In DataFields
      With pvtChart.Chart.PivotLayout.pivotTable
        'Captionはスペースを結合しないとエラーになる(Fieldsと同じ名前がNG?)
        .AddDataField .PivotFields(d.Item), d.Item & " ", d.func
      End With
    Next d
  End If

  With pvtChart.Chart
    .SetElement (msoElementLegendRight)
    .ChartTitle.Text = graphName
    .Location Where:=xlLocationAsNewSheet, Name:=graphName
  End With

End Function

Public Function SaveAndCloseBook(toPath As String)
  bk.SaveAs FileName:=toPath, FileFormat:=xlOpenXMLWorkbook
  bk.Close

End Function

Public Function UpdateBoundary()
  lastColNum = st.Cells(headerRow, 1).End(xlToRight).Column
  lastRowNum = st.Cells(headerRow, 1).End(xlDown).Row

End Function
