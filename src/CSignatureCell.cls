VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSignatureCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSignatureCell - Calcul de signature pour cellules, dimensions, multilignes
' Gere : msdElementTypeCellHeader, msdElementTypeDimension,
'         msdElementTypeMultiLine
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureCell(ByVal oCell As CellElement) As String
    Dim s As String
    s = oCell.Name & "|"
    s = s & FormatPt(oCell.Origin) & "|"
    s = s & FmtD(oCell.Rotation) & "|"
    s = s & FmtD(oCell.ScaleX) & "|"
    s = s & FmtD(oCell.ScaleY) & "|"
    s = s & FmtD(oCell.ScaleZ)
    SignatureCell = s
End Function

Function SignatureDimension(ByVal oDim As DimensionElement) As String
    Dim s As String
    Dim oRange As Range3d

    oRange = oDim.Range
    s = FormatPt(oRange.Low) & "|"
    s = s & FormatPt(oRange.High) & "|"

    On Error Resume Next
    s = s & oDim.Text
    On Error GoTo 0

    SignatureDimension = s
End Function

Function SignatureMultiLine(ByVal oML As MultiLineElement) As String
    Dim pts() As Point3d
    Dim s As String
    Dim i As Long

    pts = oML.GetVertices
    s = CStr(UBound(pts) - LBound(pts) + 1) & "V|"
    For i = LBound(pts) To UBound(pts)
        s = s & FormatPt(pts(i)) & ";"
    Next i
    s = s & BoolStr(oML.IsClosed)
    SignatureMultiLine = s
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
