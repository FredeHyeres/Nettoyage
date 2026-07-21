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
' Formatage delegue au module partage MFormatSig.
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureArc(ByVal oArc As ArcElement) As String
    Dim s As String
    s = MFormatSig.FormatPt(oArc.CenterPoint, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oArc.PrimaryRadius, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oArc.SecondaryRadius, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oArc.StartAngle, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oArc.SweepAngle, m_dTol) & "|"
    s = s & MFormatSig.FmtRot(oArc.Rotation, m_dTol)
    SignatureArc = s
End Function

Function SignatureEllipse(ByVal oEllipse As EllipseElement) As String
    Dim s As String
    s = MFormatSig.FormatPt(oEllipse.CenterPoint, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oEllipse.PrimaryRadius, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oEllipse.SecondaryRadius, m_dTol) & "|"
    s = s & MFormatSig.BoolStr(oEllipse.IsFilled)
    SignatureEllipse = s
End Function
