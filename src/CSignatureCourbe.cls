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
        s = s & FormatPt(pts(i)) & ";"
    Next i
    SignatureCurve = s
End Function

Function SignatureBSplineCurve(ByVal oBSC As BsplineCurveElement) As String
    Dim s As String
    Dim pts() As Point3d
    Dim knots() As Double
    Dim i As Long

    pts = oBSC.GetPoles
    s = CStr(oBSC.Order) & "|"
    s = s & BoolStr(oBSC.IsClosed) & "|"
    s = s & CStr(UBound(pts) - LBound(pts) + 1) & "P|"

    For i = LBound(pts) To UBound(pts)
        s = s & FormatPt(pts(i)) & ";"
    Next i

    knots = oBSC.GetKnots
    s = s & CStr(UBound(knots) - LBound(knots) + 1) & "K|"
    For i = LBound(knots) To UBound(knots)
        s = s & FmtD(knots(i)) & ";"
    Next i

    SignatureBSplineCurve = s
End Function

Function SignatureBSplineSurface(ByVal oElement As Element) As String
    Dim s As String
    Dim oRange As Range3d

    oRange = oElement.Range
    s = "BSURF|"
    s = s & FormatPt(oRange.Low) & "|"
    s = s & FormatPt(oRange.High)
    SignatureBSplineSurface = s
End Function

' --- Utilitaires prives ---

Private Function FormatPt(ByRef pt As Point3d) As String
    FormatPt = FmtD(pt.X) & "," & FmtD(pt.Y) & "," & FmtD(pt.Z)
End Function

Private Function FmtD(ByVal d As Double) As String
    FmtD = Format$(Round(d / m_dTol) * m_dTol, "0.0000")
End Function

Private Function BoolStr(ByVal b As Boolean) As String
    If b Then BoolStr = "T" Else BoolStr = "F"
End Function
