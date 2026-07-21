VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSignatureLigne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSignatureLigne - Calcul de signature pour lignes et linestrings
' Gere : msdElementTypeLine, msdElementTypeLineString, msdElementTypePointString
' Formatage delegue au module partage MFormatSig.
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureLigne(ByVal oLine As LineElement) As String
    ' StartPoint/EndPoint sont propres a ArcElement : sur une ligne on prend les
    ' extremites via GetVertices (voie confirmee).
    Dim pts() As Point3d
    pts = oLine.GetVertices
    SignatureLigne = MFormatSig.FormatPt(pts(LBound(pts)), m_dTol) & ">" & _
                     MFormatSig.FormatPt(pts(UBound(pts)), m_dTol)
End Function

Function SignatureLineString(ByVal oLine As LineElement) As String
    Dim pts() As Point3d
    Dim s As String
    Dim i As Long

    pts = oLine.GetVertices
    s = CStr(UBound(pts) - LBound(pts) + 1) & "V|"
    For i = LBound(pts) To UBound(pts)
        s = s & MFormatSig.FormatPt(pts(i), m_dTol) & ";"
    Next i
    SignatureLineString = s
End Function

Function SignaturePointString(ByVal oLine As LineElement) As String
    SignaturePointString = SignatureLineString(oLine)
End Function
