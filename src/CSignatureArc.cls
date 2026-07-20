VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSignatureArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSignatureArc - Calcul de signature pour arcs et ellipses
' Gere : msdElementTypeArc, msdElementTypeEllipse
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureArc(ByVal oArc As ArcElement) As String
    Dim s As String
    s = FormatPt(oArc.CenterPoint) & "|"
    s = s & FmtD(oArc.PrimaryRadius) & "|"
    s = s & FmtD(oArc.SecondaryRadius) & "|"
    s = s & FmtD(oArc.StartAngle) & "|"
    s = s & FmtD(oArc.SweepAngle) & "|"
    s = s & FmtD(oArc.Rotation)
    SignatureArc = s
End Function

Function SignatureEllipse(ByVal oEllipse As EllipseElement) As String
    Dim s As String
    s = FormatPt(oEllipse.CenterPoint) & "|"
    s = s & FmtD(oEllipse.PrimaryRadius) & "|"
    s = s & FmtD(oEllipse.SecondaryRadius) & "|"
    s = s & BoolStr(oEllipse.IsFilled)
    SignatureEllipse = s
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
