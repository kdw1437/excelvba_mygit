VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "postDisDivMissingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Module: postDisDiv
Private pWorksheet As Worksheet
Private pStartCell As Range
Private pK As Long

' Property for Worksheet
' @property Get Worksheet
' @return {Worksheet} worksheet ��ü
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

' @property Set Worksheet
' @param {Worksheet} value
Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' Property for StartCell
' @property Get StartCell
' @return {Range} starting cell range
Public Property Get StartCell() As Range
    Set StartCell = pStartCell
End Property

' @property Set StartCell
' @param {Range} value
Public Property Set StartCell(value As Range)
    Set pStartCell = value
End Property

' Property for K
' @property Get k
' @return {Long} K��
Public Property Get k() As Long
    k = pK
End Property

' @property Let k
' @param {Long} value - set K��
Public Property Let k(value As Long)
    pK = value
End Property

' Method to Generate JSON
'
' @method ReturnJSON
' @return {String} worksheet�� data�� ���� ������� JSON string
Public Function ReturnJSON() As String
    Dim dividendStreamCollection As New Collection '���� �ٱ��� jsonArray
    Dim i As Long, j As Long

    For i = 1 To pK
        Dim dataIdCell As Range
        Set dataIdCell = pStartCell.Offset(3, 1 + 3 * (i - 1))
        
        Dim dataId As String
        dataId = dataIdCell.value
        
        Dim disDivData As Object
        Set disDivData = CreateObject("Scripting.Dictionary")
        disDivData("dataId") = dataId '���� �ٱ��� jsonArray�� jsonObject�� key dataId�� value�� �Ҵ�
        
        Dim dividendsCollection As New Collection 'jsonObject�� key�� value���� �� jsonarray ����
        Set dividendsCollection = New Collection '���ʱ�ȭ (jsonArray ���ʱ�ȭ)
        j = 4 ' Start row offset for yields
        
        Do While pWorksheet.Cells(dataIdCell.row + j, dataIdCell.Column).value <> ""
            Dim divData As Object
            Set divData = CreateObject("Scripting.Dictionary") 'jsonarray�� ��Ұ� �� jsonObject ����
            divData("date") = CStr(pWorksheet.Cells(dataIdCell.row + j, dataIdCell.Column - 1).value)
            divData("value") = pWorksheet.Cells(dataIdCell.row + j, dataIdCell.Column).value
            dividendsCollection.Add divData 'divData jsonObject(Dictionary)�� jsonArray�� �־��ش�.
            j = j + 1
            DoEvents
        Loop
        
        Dim divsArray() As Variant
        ReDim divsArray(1 To dividendsCollection.Count) '������� dictionary�� ��ҷΰ����� collection�� Array���·� �ٲ㼭
        'structured�� json��ü�� �����Ѵ�.
        
        Dim idx As Long
        For idx = 1 To dividendsCollection.Count
            Set divsArray(idx) = dividendsCollection(idx) 'array�� ù��°�� collection�� ù��° dictionary�� �־��ش�.
            DoEvents
        Next idx
        
        disDivData("dividends") = divsArray 'dictionary�� key dividends�� divsArray(json array) value�� �Ҵ��Ѵ�.
        dividendStreamCollection.Add disDivData '��� ���� �Ҵ�� dictionary�� collection�� ������� �߰��Ѵ�.
        DoEvents
    Next i
    
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dividendStreamCollection) 'collection�� ConvertToJson�޼ҵ带 �̿��ؼ� jsonString���� ��ȯ�Ѵ�.
    
    Debug.Print jsonString
    ReturnJSON = jsonString
    
End Function

