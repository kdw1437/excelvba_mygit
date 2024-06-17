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

'Getter, Setter ����. URL�� �����ϴ� ��ҵ� (property(field)�̴�.)
Public Property Get baseurl() As String
    baseurl = pBaseURL
End Property

Public Property Let baseurl(value As String)
    pBaseURL = value
End Property

Public Property Get DataParameter() As String
    DataParameter = pDateParameter
End Property

Public Property Let DataParameter(value As String)
    pDataParameter = value
End Property

Public Property Get DataIds() As String
    DataIds = pDataIds
End Property

Public Property Let DataIds(value As String)
    pDataIds = value
End Property

Public Property Get dataSetId() As String
    dataSetId = pdataSetId
End Property

Public Property Let dataSetId(value As String)
    pdataSetId = value
End Property

Public Property Get Version() As String
    Version = pVersion
End Property

Public Property Let Version(value As String)
    pVersion = value
End Property

Public Property Get MatrixId() As String
    MatrixId = pMatrixId
End Property

Public Property Let MatrixId(value As String)
    pMatrixId = value
End Property

Public Property Get baseDt() As String
    baseDt = pBaseDt
End Property

Public Property Let baseDt(value As String)
    pBaseDt = value

End Property

' URL�� ����� method
' @method MakeUrl
' @return {String}
' @usage - property(field)���� �������� URL�� �����Ͽ� ��ȯ�Ѵ�.
Public Function MakeUrl() As String
    Dim url As String
    url = pBaseURL & pVersion & pDataParameter & pBaseDt & pDataIds & pdataSetId & pMatrixId
    MakeUrl = url
End Function