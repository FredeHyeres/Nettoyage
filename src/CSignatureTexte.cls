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
' Formatage delegue au module partage MFormatSig.
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureTexte(ByVal oText As TextElement) As String
    Dim s As String
    s = oText.Text & "|"
    s = s & MFormatSig.FormatPt(oText.Origin, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oText.TextStyle.Height, m_dTol) & "|"
    s = s & MFormatSig.FmtD(oText.TextStyle.Width, m_dTol) & "|"
    s = s & CStr(oText.TextStyle.Font.ID)
    SignatureTexte = s
End Function

Function SignatureTextNode(ByVal oNode As TextNodeElement) As String
    ' Un TextNode n'expose pas .Text ni .TextStyle : ils vivent sur les
    ' TextElement enfants (chaque ligne a son propre style). On itere donc les
    ' sous-elements et on integre texte + style de CHAQUE ligne a la signature.
    Dim s As String
    Dim ee As ElementEnumerator
    Dim oLigne As TextElement

    s = MFormatSig.FormatPt(oNode.Origin, m_dTol) & "|"
    s = s & CStr(oNode.TextLinesCount) & "|"

    Set ee = oNode.GetSubElements
    Do While ee.MoveNext
        If ee.Current.IsTextElement Then
            Set oLigne = ee.Current
            s = s & oLigne.Text & "~"
            s = s & MFormatSig.FmtD(oLigne.TextStyle.Height, m_dTol) & "~"
            s = s & MFormatSig.FmtD(oLigne.TextStyle.Width, m_dTol) & "~"
            s = s & CStr(oLigne.TextStyle.Font.ID) & ";"
        End If
    Loop

    SignatureTextNode = s
End Function
