Attribute VB_Name = "Functions"
' VBA������ static method�� �������� �����Ƿ�, class module�� �޼ҵ� �ۼ� �ÿ��� Ŭ������ �ν��Ͻ� ���� �Ŀ� ����� �ؾ��մϴ�.
' ��� �̷��� ����� ���ϱ� ���ؼ� ���� ����ϴ� �Լ��� Functions��⿡ ������ �Ǿ����ϴ�.
' SOC(Separation of Concern), ��� ����

' @method GetHttpResponseText
' @param {String} url
' @return {String}
' @usage - HTTP GET request�� �ϰ�, response text�� ���� �޴� �Լ�
Public Function GetHttpResponseText(url As String) As String
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        GetHttpResponseText = .ResponseText
    End With
End Function
' Asynchronous Process
' @method GetHttpResponseText2
' @param {String} url
' @return {String}
' @usage - HTTP GET request�� �ϰ�, response text�� ���� ���� ������ ������ responsive�ϵ��� ����� �ִ� �Լ�
Public Function GetHttpResponseText2(url As String) As String
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, True
        .Send
        
        While .readyState <> 4
            DoEvents
        Wend
        
        GetHttpResponseText2 = .ResponseText
    End With
End Function
' @method ParseJson
' @param {String} jsonString
' @return {Object}
' @usage - JSON string�� VBA ��ü�� �Ľ��ϴ� �Լ�
Public Function ConverterParseJson(JsonString As String) As Object
    ' JsonConverter ����� �����ϰ� �־�� �Ѵ�.
    Set ConverterParseJson = JsonConverter.ParseJson(JsonString)
End Function

' @method GetJsonResponse
' @param {String} url
' @return {Object}
' @usage - GetHttpResponseText�Լ��� ParseJson�Լ� �̿�. url�� input���� ������, �ش� json string�� �Ľ����ְ�, ��� vba ��ü�� return�Ѵ�.
Public Function GetJsonResponse(url As String) As Object
    Dim JsonString As String
    JsonString = GetHttpResponseText(url)
    Set GetJsonResponse = ConverterParseJson(JsonString)
End Function

' @method URLEncode
' @param {String} StringVal
' @param {Boolean} [SpaceAsPlus=False]
' @return {String}
' @usage - StringVal�� ���ڵ����ִ� �Լ�. EncodeChar�Լ� �̿�
Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long: StringLen = Len(StringVal)

    If StringLen > 0 Then
        ReDim result(StringLen) As String 'StringLen���̷� string�� dynamic array�� result�� �ʱ�ȭ�Ѵ�.
        Dim i As Long

        For i = 1 To StringLen
            result(i) = EncodeChar(Mid$(StringVal, i, 1), SpaceAsPlus) 'dynamic array�� result�� encoding�� character�� ������� ����ȴ�.
        Next i

        URLEncode = Join(result, "") 'Join�Լ�: result array�� ��� ��Ҹ� empty string("")�� separator�� �� �ٿ� ������. �ᱹ���� ��� character�� �̾ encoding�� string�� �����.
    End If
End Function

' @method EncodeChar
' @param {String} Char
' @param {Boolean} SpaceAsPlus
' @return {String}
' @usage - �ϳ��� character�� string�� input���� �޴´�. �ϳ��� character�� string�� ���ڵ��ϴ� �޼ҵ��̴�.
' SpaceAsPlus�� space�� encoding�Ǵ� ����� �����Ѵ�. true�� ��, +��, false�� ��, %20���� ���ڵ� �ȴ�.
Private Function EncodeChar(Char As String, SpaceAsPlus As Boolean) As String
    Dim CharCode As Integer: CharCode = Asc(Char)
    Dim Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
            EncodeChar = Char
        Case 32
            EncodeChar = Space
        Case 0 To 15
            EncodeChar = "%0" & Hex(CharCode)
        Case Else
            EncodeChar = "%" & Hex(CharCode)
    End Select
End Function
' @method IsDuplicate
' @param {String} dataId1
' @param {String} dataId2
' @param {Dictionary} uniquePairs
' @return {Boolean} - data ID pair�� �ߺ��̸�, Returns True. �ƴϸ� False
' @usage - �Լ��� 2���� dataId�� �Ѱ��� dictionary�� �޴´�. �Լ��� dataId1�� dataId2�� ���� Ȥ�� dataId2�� dataId1 ������ dictionary�� �����ϴ��� Ȯ���Ѵ�.
' duplicate�� ��Ÿ����, dictionary���� combination�� �ϳ��� �߰ߵǸ� True�� return�Ѵ�.
Function IsDuplicate(ByVal dataId1 As String, ByVal dataId2 As String, ByRef uniquePairs As Scripting.Dictionary) As Boolean
    ' Check if the combinedKey or its reverse is in the dictionary
    IsDuplicate = uniquePairs.Exists(dataId1 & ":" & dataId2) Or uniquePairs.Exists(dataId2 & ":" & dataId1)
End Function

' @method MapDataIdToCode
' @param {String} dataId
' @return {String}
' @usage - dataId ���� �����Ͽ� �ڵ� ���� ��ȯ�ϴ� �Լ�
Function MapDataIdToCode(dataId As String) As String
    ' mapping�Լ�
    Select Case dataId
        Case "KOSPI200_LOC"
            MapDataIdToCode = "KOSPI_LV" ' KOSPI200_LOC ���� KOSPI_LV�� ���εȴ�.
        Case "HSI_LOC"
            MapDataIdToCode = "HSI_LV" ' HSI_LOC ���� HSI_LV�� ���εȴ�.
        Case "N225_LOC"
            MapDataIdToCode = "NKY_LV" ' N225_LOC ���� NKY_LV�� ���εȴ�.
        Case "HSCEI_LOC"
            MapDataIdToCode = "HSCEI_LV" ' HSCEI_LOC ���� HSCEI_LV�� ���εȴ�.
        
        Case Else
            MapDataIdToCode = dataId ' mapping�� �ش�Ǵ� ��찡 ���� ����, �״�� return�Ѵ�.
    End Select
End Function

' @method ImportGreekData
' @param {String} jobId
' @usage - jobId�� �ش��ϴ� greek���� cell�� �ѷ��ش�.
Function ImportGreekData(jobId As String)
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim json As Object

    ' XMLHTTP ��ü ����
    Set objRequest = CreateObject("MSXML2.XMLHTTP")

    ' jobId�� url����
    strUrl = "http://urosys-web.juroinstruments.com/app/selectGreeks?jobId=" & jobId
    blnAsync = True

    ' get request
    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send

        ' ������ ��ٸ���.
        While .readyState <> 4
            DoEvents ' ������ ��� ���� �����Ѵ�.
        Wend

        ' ������ ��´�.
        strResponse = .ResponseText
    End With

    ' JSON response �Ľ�
    Set json = JsonConverter.ParseJson(strResponse)
    Dim rowNumber As Integer
    rowNumber = 9

    Dim colNumber As Integer ' �Լ����� ������ �ʴ´�.
    colNumber = 6

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OTC")

    ' column ����
    Dim colMapping As Object
    Set colMapping = CreateObject("Scripting.Dictionary")

    colMapping("rfCd") = 8
    colMapping("jobId") = 6
    colMapping("sensTyCd") = 11
    colMapping("delta") = 9
    colMapping("itemCd") = 7
    colMapping("gamma") = 10

    Dim key As Variant
    Dim Item As Object

    ' JSON data�� ��ȸ�Ѵ�.
    For Each Item In json("selectGreekValues")
        For Each key In Item.Keys
            If colMapping.Exists(key) Then
                ws.Cells(rowNumber, colMapping(key)).value = Item(key)
            End If
            DoEvents
        Next key
        rowNumber = rowNumber + 1
        DoEvents
    Next Item
End Function

