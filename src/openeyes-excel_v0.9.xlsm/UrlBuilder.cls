VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UrlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pBaseURL As String
Private pDataParameter As String
Private pDataIds As String
Private pdataSetId As String
Private pVersion As String
Private pMatrixId As String
Private pBaseDt As String

'Getter, Setter 설정. URL을 구성하는 요소들 (property(field)이다.)
Public Property Get baseURL() As String
    baseURL = pBaseURL
End Property

Public Property Let baseURL(Value As String)
    pBaseURL = Value
End Property

Public Property Get DataParameter() As String
    DataParameter = pDateParameter
End Property

Public Property Let DataParameter(Value As String)
    pDataParameter = Value
End Property

Public Property Get DataIds() As String
    DataIds = pDataIds
End Property

Public Property Let DataIds(Value As String)
    pDataIds = Value
End Property

Public Property Get dataSetId() As String
    dataSetId = pdataSetId
End Property

Public Property Let dataSetId(Value As String)
    pdataSetId = Value
End Property

Public Property Get Version() As String
    Version = pVersion
End Property

Public Property Let Version(Value As String)
    pVersion = Value
End Property

Public Property Get MatrixId() As String
    MatrixId = pMatrixId
End Property

Public Property Let MatrixId(Value As String)
    pMatrixId = Value
End Property

Public Property Get baseDt() As String
    baseDt = pBaseDt
End Property

Public Property Let baseDt(Value As String)
    pBaseDt = Value

End Property
'URL을 만드는 method(property(field)값을 바탕으로 호출 시, URL을 return한다.)
Public Function MakeUrl() As String
    Dim url As String
    url = pBaseURL & pVersion & pDataParameter & pBaseDt & pDataIds & pdataSetId & pMatrixId
    MakeUrl = url
End Function
