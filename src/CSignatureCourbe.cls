VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSignatureCourbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSignatureCourbe - Calcul de signature pour courbes et B-splines
' Gere : msdElementTypeCurve, msdElementTypeBsplineCurve,
'         msdElementTypeBsplineSurface
' Formatage delegue au module partage MFormatSig.
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureCurve(ByVal oCurve As CurveElement) As String
    Dim pts() As Point3d
    Dim s As String
    Dim i As Long

    pts = oCurve.GetVertices
    s = CStr(UBound(pts) - LBound(pts) + 1) & "V|"
    For i = LBound(pts) To UBound(pts)
        s = s & MFormatSig.FormatPt(pts(i), m_dTol) & ";"
    Next i
    SignatureCurve = s
End Function

Function SignatureBSplineCurve(ByVal oBSC As BsplineCurveElement) As String
    ' Order/GetPoles/GetKnots sont sur l'objet geometrie BsplineCurve, pas sur
    ' l'element : on l'extrait d'abord. Le flag "closed" n'a pas de propriete
    ' dediee et reste implicite dans les noeuds/poles -> non necessaire ici.
    Dim s As String
    Dim oBc As BsplineCurve
    Dim pts() As Point3d
    Dim knots() As Double
    Dim i As Long

    Set oBc = oBSC.ExtractBsplineCurve

    s = CStr(oBc.Order) & "|"
    pts = oBc.GetPoles
    s = s & CStr(UBound(pts) - LBound(pts) + 1) & "P|"
    For i = LBound(pts) To UBound(pts)
        s = s & MFormatSig.FormatPt(pts(i), m_dTol) & ";"
    Next i

    knots = oBc.GetKnots
    s = s & CStr(UBound(knots) - LBound(knots) + 1) & "K|"
    For i = LBound(knots) To UBound(knots)
        s = s & MFormatSig.FmtD(knots(i), m_dTol) & ";"
    Next i

    SignatureBSplineCurve = s
End Function

Function SignatureBSplineSurface(ByVal oElement As Element) As String
    Dim s As String
    Dim oRange As Range3d

    oRange = oElement.Range
    s = "BSURF|"
    s = s & MFormatSig.FormatPt(oRange.Low, m_dTol) & "|"
    s = s & MFormatSig.FormatPt(oRange.High, m_dTol)
    SignatureBSplineSurface = s
End Function
