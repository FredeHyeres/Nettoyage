VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSignatureTexte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSignatureTexte - Calcul de signature pour textes et text nodes
' Gere : msdElementTypeText, msdElementTypeTextNode
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureTexte(ByVal oText As TextElement) As String
    Dim s As String
    s = oText.Text & "|"
    s = s & FormatPt(oText.Origin) & "|"
    s = s & FmtD(oText.TextStyle.Height) & "|"
    s = s & FmtD(oText.TextStyle.Width) & "|"
    s = s & CStr(oText.TextStyle.Font.ID)
    SignatureTexte = s
End Function

Function SignatureTextNode(ByVal oNode As TextNodeElement) As String
    Dim s As String
    Dim i As Long

    s = FormatPt(oNode.Origin) & "|"
    s = s & CStr(oNode.TextLinesCount) & "|"

    For i = 1 To oNode.TextLinesCount
        s = s & oNode.TextLine(i) & ";"
    Next i

    s = s & FmtD(oNode.TextStyle.Height) & "|"
    s = s & FmtD(oNode.TextStyle.Width) & "|"
    s = s & CStr(oNode.TextStyle.Font.ID)
    SignatureTextNode = s
End Function

' --- Utilitaires prives ---

Private Function FormatPt(ByRef pt As Point3d) As String
    FormatPt = FmtD(pt.X) & "," & FmtD(pt.Y) & "," & FmtD(pt.Z)
End Function

Private Function FmtD(ByVal d As Double) As String
    FmtD = Format$(Round(d / m_dTol) * m_dTol, "0.0000")
End Function
