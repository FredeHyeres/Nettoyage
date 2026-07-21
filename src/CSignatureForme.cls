VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSignatureForme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSignatureForme - Calcul de signature pour shapes et complex strings/shapes
' Gere : msdElementTypeShape, msdElementTypeComplexString,
'         msdElementTypeComplexShape
' Formatage delegue au module partage MFormatSig.
' =============================================================================

Private m_dTol As Double

Sub Init(ByVal dTolerance As Double)
    m_dTol = dTolerance
End Sub

Function SignatureShape(ByVal oShape As ShapeElement) As String
    Dim pts() As Point3d
    Dim s As String
    Dim i As Long

    pts = oShape.GetVertices
    s = CStr(UBound(pts) - LBound(pts) + 1) & "V|"
    s = s & MFormatSig.BoolStr(oShape.IsFilled) & "|"
    For i = LBound(pts) To UBound(pts)
        s = s & MFormatSig.FormatPt(pts(i), m_dTol) & ";"
    Next i
    SignatureShape = s
End Function

Function SignatureComplexString(ByVal oComplex As ComplexStringElement) As String
    Dim s As String
    Dim oEnum As ElementEnumerator
    Dim oSub As Element
    Dim lCount As Long

    s = ""
    lCount = 0
    Set oEnum = oComplex.GetSubElements

    Do While oEnum.MoveNext
        Set oSub = oEnum.Current
        lCount = lCount + 1
        s = s & SignatureSousElement(oSub) & "#"
    Loop

    SignatureComplexString = CStr(lCount) & "S|" & s
End Function

Function SignatureComplexShape(ByVal oComplex As ComplexShapeElement) As String
    Dim s As String
    Dim oEnum As ElementEnumerator
    Dim oSub As Element
    Dim lCount As Long

    s = ""
    lCount = 0
    Set oEnum = oComplex.GetSubElements

    Do While oEnum.MoveNext
        Set oSub = oEnum.Current
        lCount = lCount + 1
        s = s & SignatureSousElement(oSub) & "#"
    Loop

    SignatureComplexShape = MFormatSig.BoolStr(oComplex.IsFilled) & "|" & CStr(lCount) & "S|" & s
End Function

' --- Signature simplifiee d'un sous-element (sans recurse infinie) ---
Private Function SignatureSousElement(ByVal oSub As Element) As String
    Dim s As String
    Dim oRange As Range3d

    s = CStr(oSub.Type) & ":"
    oRange = oSub.Range
    s = s & MFormatSig.FormatPt(oRange.Low, m_dTol) & ">" & MFormatSig.FormatPt(oRange.High, m_dTol)
    SignatureSousElement = s
End Function
