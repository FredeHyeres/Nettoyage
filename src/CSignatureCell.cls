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
' Formatage delegue au module partage MFormatSig.
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureCell(ByVal oCell As CellElement) As String
    ' ScaleX/Y/Z n'existent pas : l'echelle est un Point3d (oCell.Scale).
    ' Rotation est une Matrix3d (serialisee par FmtRot, pas traitee en Double).
    Dim s As String
    s = oCell.Name & "|"
    s = s & MFormatSig.FormatPt(oCell.Origin, m_dTol) & "|"
    s = s & MFormatSig.FmtRot(oCell.Rotation, m_dTol) & "|"
    s = s & MFormatSig.FormatPt(oCell.Scale, m_dTol)
    SignatureCell = s
End Function

Function SignatureDimension(ByVal oDim As DimensionElement) As String
    Dim s As String
    Dim oRange As Range3d

    oRange = oDim.Range
    s = MFormatSig.FormatPt(oRange.Low, m_dTol) & "|"
    s = s & MFormatSig.FormatPt(oRange.High, m_dTol) & "|"

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
        s = s & MFormatSig.FormatPt(pts(i), m_dTol) & ";"
    Next i
    s = s & MFormatSig.BoolStr(oML.IsClosed)
    SignatureMultiLine = s
End Function
