Attribute VB_Name = "Functions"
'VBA������ static method�� �������� �����Ƿ�, class module�� �޼ҵ� �ۼ� �ÿ��� Ŭ������ �ν��Ͻ� ���� �Ŀ� ����� �ؾ��մϴ�.
'��� �̷��� ����� ���ϱ� ���ؼ� ���� ����ϴ� �Լ��� Functions��⿡ ������ �Ǿ����ϴ�.
'SOC(Separation of Concern), ��� ����

'@method GetHttpResponseText
'@param {String} url
'@return {String}
'@usage - HTTP GET request�� �ϰ�, response text�� ���� �޴� �Լ�
Public Function GetHttpResponseText(url As String) As String
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        GetHttpResponseText = .responseText
    End With
End Function

'@method ParseJson
'@param {String} jsonString
'@return {Object}
'@usage - JSON string�� VBA ��ü�� �Ľ��ϴ� �Լ�
Public Function ConverterParseJson(jsonString As String) As Object
    ' JsonConverter ����� �����ϰ� �־�� �Ѵ�.
    Set ConverterParseJson = JsonConverter.ParseJson(jsonString)
End Function

'@method GetJsonResponse
'@param {String} url
'@return {Object}
'@usage - GetHttpResponseText�Լ��� ParseJson�Լ� �̿�. url�� input���� ������, �ش� json string�� �Ľ����ְ�, ��� vba ��ü�� return�Ѵ�.
Public Function GetJsonResponse(url As String) As Object
    Dim jsonString As String
    jsonString = GetHttpResponseText(url)
    Set GetJsonResponse = ConverterParseJson(jsonString)
End Function

'@method URLEncode
'@param {String} StringVal
'@param {Boolean} [SpaceAsPlus=False]
'@return {String}
'@usage - StringVal�� ���ڵ����ִ� �Լ�. EncodeChar�Լ� �̿�
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

'@method EncodeChar
'@param {String} Char
'@param {Boolean} SpaceAsPlus
'@return {String}
'@usage - �ϳ��� character�� string�� input���� �޴´�. �ϳ��� character�� string�� ���ڵ��ϴ� �޼ҵ��̴�.
'SpaceAsPlus�� space�� encoding�Ǵ� ����� �����Ѵ�. true�� ��, +��, false�� ��, %20���� ���ڵ� �ȴ�.
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
'@method IsDuplicate
'@param {String} dataId1
'@param {String} dataId2
'@param {Dictionary} uniquePairs
'@return {Boolean} - data ID pair�� �ߺ��̸�, Returns True. �ƴϸ� False
'@usage - �Լ��� 2���� dataId�� �Ѱ��� dictionary�� �޴´�. �Լ��� dataId1�� dataId2�� ���� Ȥ�� dataId2�� dataId1 ������ dictionary�� �����ϴ��� Ȯ���Ѵ�.
'duplicate�� ��Ÿ����, dictionary���� combination�� �ϳ��� �߰ߵǸ� True�� return�Ѵ�.
Function IsDuplicate(ByVal dataId1 As String, ByVal dataId2 As String, ByRef uniquePairs As Scripting.Dictionary) As Boolean
    ' Check if the combinedKey or its reverse is in the dictionary
    IsDuplicate = uniquePairs.Exists(dataId1 & ":" & dataId2) Or uniquePairs.Exists(dataId2 & ":" & dataId1)
End Function

Function MapDataIdToCode(dataId As String) As String
    ' mapping�Լ�
    Select Case dataId
        Case "KOSPI200_LOC"
            MapDataIdToCode = "KOSPI_LV"
        Case "HSI_LOC"
            MapDataIdToCode = "HSI_LV"
        Case "N225_LOC"
            MapDataIdToCode = "NKY_LV"
        Case "HSCEI_LOC"
            MapDataIdToCode = "HSCEI_LV"
        
        Case Else
            MapDataIdToCode = dataId ' mapping�� �ش�Ǵ� ��찡 ���� ����, �״�� return�Ѵ�.
    End Select
End Function


