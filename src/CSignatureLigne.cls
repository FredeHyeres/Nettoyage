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
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureLigne(ByVal oLine As LineElement) As String
    Dim s As String
    s = FormatPt(oLine.StartPoint) & ">" & FormatPt(oLine.EndPoint)
    SignatureLigne = s
End Function

Function SignatureLineString(ByVal oLine As LineElement) As String
    Dim pts() As Point3d
    Dim s As String
    Dim i As Long

    pts = oLine.GetVertices
    s = CStr(UBound(pts) - LBound(pts) + 1) & "V|"
    For i = LBound(pts) To UBound(pts)
        s = s & FormatPt(pts(i)) & ";"
    Next i
    SignatureLineString = s
End Function

Function SignaturePointString(ByVal oLine As LineElement) As String
    SignaturePointString = SignatureLineString(oLine)
End Function

' --- Utilitaires prives ---

Private Function FormatPt(ByRef pt As Point3d) As String
    FormatPt = FmtD(pt.X) & "," & FmtD(pt.Y) & "," & FmtD(pt.Z)
End Function

Private Function FmtD(ByVal d As Double) As String
    FmtD = Format$(Round(d / m_dTol) * m_dTol, "0.0000")
End Function
