VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DivYieldUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pWs As Worksheet
Private pDivRange As Range
Private pDivData As Collection

' Setter for the Worksheet
Public Sub SetWorksheet(sheetName As String)
    Set pWs = ThisWorkbook.Sheets(sheetName)
End Sub

' Getter for the Worksheet
Public Function GetWorksheet() As Worksheet
    Set GetWorksheet = pWs
End Function

' Setter for DivData Collection
Public Sub SetDivData(col As Collection)
    Set pDivData = col
End Sub

' Getter for DivData Collection
Public Function GetDivData() As Collection
    Set GetDivData = pDivData
End Function

'Setter for DivRange
Public Sub SetDivRange(rng As String)
    Set pDivRange = pWs.Range(rng)
End Sub

'Getter for DivRange
Public Function GetDivRange() As Range
    Set GetDivRange = pDivRange
End Function

Public Sub UpdateWorksheet()

    Dim element As Variant
    Dim cell As Range
    For Each element In pDivData
        For Each cell In pDivRange
            If element("dataId") = cell.value + "_C" Then
                cell.Offset(0, 1).value = element("yield")
            End If
        Next cell
    Next element

End Sub

Public Sub UpdateWorksheetEfficient()

    ' dataIds�� yields�� ���� mapping�� ������ Dictionary ��ü�� �����.
    Dim divDataDict As Object
    Set divDataDict = CreateObject("Scripting.Dictionary")
    
    Dim element As Variant
    
    ' dataIds���� key�� yields���� item���� Dictionary ��ü�� ä���.
    For Each element In pDivData
        Dim dataId As String
        dataId = element("dataId")
        Dim yield As Variant
        yield = element("yield")
        divDataDict(dataId) = yield
    Next element
    
    Dim cell As Range
    
    ' Ư�� Range�� cell�� iterate�Ѵ�.
    For Each cell In pDivRange
        Dim cellDataId As String
        cellDataId = cell.value + "_C"
        
        ' ���� cell�� dataId�� Dictionary�� �ִ��� Ȯ���Ѵ�.
        If divDataDict.Exists(cellDataId) Then
            ' �� Į���� �ִ� cell�� yield������ update�Ѵ�.
            cell.Offset(0, 1).value = divDataDict(cellDataId)
        End If
    Next cell

End Sub
