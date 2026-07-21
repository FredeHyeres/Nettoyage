VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMoteurScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CMoteurScan - Orchestration du scan et detection des doublons
' Connait l'API MicroStation (scan, suppression)
' Delegue le calcul de signature aux classes CSignatureXxx
' =============================================================================

' --- Classes metier ---
Private m_oSigLigne     As CSignatureLigne
Private m_oSigTexte     As CSignatureTexte
Private m_oSigArc       As CSignatureArc
Private m_oSigCourbe    As CSignatureCourbe
Private m_oSigForme     As CSignatureForme
Private m_oSigCell      As CSignatureCell

' --- Resultats ---
Private m_lNbScannes    As Long
Private m_lNbTraites    As Long
Private m_lNbIgnores    As Long
Private m_lNbDoublons   As Long
Private m_lNbSupprimes  As Long
Private m_aDoublons()   As DLong
Private m_lCountDoub    As Long

' --- Surbrillance ---
Private m_bSurbrillance As Boolean

' --- Dictionnaire des signatures ---
Private m_dictSigs      As Object

' --- Proprietes en lecture ---
Public Property Get NbElementsScannes() As Long
    NbElementsScannes = m_lNbScannes
End Property

Public Property Get NbDoublonsTrouves() As Long
    NbDoublonsTrouves = m_lNbDoublons
End Property

Public Property Get NbSupprimes() As Long
    NbSupprimes = m_lNbSupprimes
End Property

' Elements scannes dont la signature etait vide (type non gere ou erreur
' d'extraction avalee) : exclus de la detection. Utile au diagnostic.
Public Property Get NbIgnores() As Long
    NbIgnores = m_lNbIgnores
End Property

' --- Initialisation des classes de signature ---
Private Sub InitSignateurs(ByVal oSettings As CSettings)
    Set m_oSigLigne = New CSignatureLigne
    m_oSigLigne.Init oSettings.dTolerance

    Set m_oSigTexte = New CSignatureTexte
    m_oSigTexte.Init oSettings.dTolerance

    Set m_oSigArc = New CSignatureArc
    m_oSigArc.Init oSettings.dTolerance

    Set m_oSigCourbe = New CSignatureCourbe
    m_oSigCourbe.Init oSettings.dTolerance

    Set m_oSigForme = New CSignatureForme
    m_oSigForme.Init oSettings.dTolerance

    Set m_oSigCell = New CSignatureCell
    m_oSigCell.Init oSettings.dTolerance
End Sub

' =============================================================================
' Analyser - Scanne le modele et identifie les doublons
' =============================================================================
Sub Analyser(ByVal oModel As ModelReference, ByVal oSettings As CSettings)
    Dim oScanCriteria As ElementScanCriteria
    Dim oEnum As ElementEnumerator
    Dim oElement As Element
    Dim sSig As String

    InitSignateurs oSettings

    If Not m_dictSigs Is Nothing Then Set m_dictSigs = Nothing
    Set m_dictSigs = CreateObject("Scripting.Dictionary")
    m_dictSigs.CompareMode = vbBinaryCompare

    m_lNbScannes = 0
    m_lNbTraites = 0
    m_lNbIgnores = 0
    m_lNbDoublons = 0
    m_lNbSupprimes = 0
    m_lCountDoub = 0
    Erase m_aDoublons
    ReDim m_aDoublons(0)

    Set oScanCriteria = New ElementScanCriteria
    oScanCriteria.ExcludeAllTypes
    ConfigurerCriteres oScanCriteria, oSettings

    ShowCommand "Nettoyage Doublons"
    ShowPrompt "Analyse en cours..."

    Set oEnum = oModel.Scan(oScanCriteria)

    Do While oEnum.MoveNext
        Set oElement = oEnum.Current
        m_lNbScannes = m_lNbScannes + 1

        If Not oElement.IsComponentElement Then
            sSig = GenererSignature(oElement, oSettings)

            If Len(sSig) > 0 Then
                m_lNbTraites = m_lNbTraites + 1

                If m_dictSigs.Exists(sSig) Then
                    m_lNbDoublons = m_lNbDoublons + 1
                    m_lCountDoub = m_lCountDoub + 1
                    ReDim Preserve m_aDoublons(m_lCountDoub - 1)
                    m_aDoublons(m_lCountDoub - 1) = oElement.ID
                Else
                    m_dictSigs.Add sSig, oElement.ID
                End If
            Else
                ' Signature vide : type non gere ou erreur d'extraction avalee
                ' par le On Error de SignatureGeometrique -> element ignore.
                m_lNbIgnores = m_lNbIgnores + 1
            End If
        End If

        If m_lNbScannes Mod 1000 = 0 Then
            ShowPrompt "Analyse: " & m_lNbScannes & " scannes, " & _
                       m_lNbDoublons & " doublons..."
            DoEvents
        End If
    Loop

    ShowPrompt "Analyse terminee: " & m_lNbDoublons & " doublons sur " & _
               m_lNbScannes & " elements (" & m_lNbIgnores & " ignores)."
End Sub

' =============================================================================
' SupprimerDoublons - Supprime les elements identifies comme doublons
' =============================================================================
Sub SupprimerDoublons(ByVal oModel As ModelReference)
    Dim i As Long
    Dim oElem As Element

    m_lNbSupprimes = 0
    ShowPrompt "Suppression en cours..."

    For i = 0 To m_lCountDoub - 1
        On Error Resume Next
        Set oElem = Nothing
        Set oElem = oModel.GetElementByID(m_aDoublons(i))
        If Not oElem Is Nothing Then
            oModel.RemoveElement oElem
            m_lNbSupprimes = m_lNbSupprimes + 1
        End If
        On Error GoTo 0

        If m_lNbSupprimes Mod 100 = 0 And m_lNbSupprimes > 0 Then
            ShowPrompt "Suppression: " & m_lNbSupprimes & " / " & m_lNbDoublons
            DoEvents
        End If
    Next i

    CommandState.StartDefaultCommand
    ShowPrompt "Termine: " & m_lNbSupprimes & " doublons supprimes."
End Sub

' =============================================================================
' SurlignierDoublons - Selectionne les doublons (selection set)
' =============================================================================
Sub SurlignierDoublons(ByVal oModel As ModelReference)
    Dim i As Long
    Dim oElem As Element

    If m_lCountDoub = 0 Then Exit Sub

    oModel.UnselectAllElements

    For i = 0 To m_lCountDoub - 1
        On Error Resume Next
        Set oElem = Nothing
        Set oElem = oModel.GetElementByID(m_aDoublons(i))
        If Not oElem Is Nothing Then
            oModel.SelectElement oElem
        End If
        On Error GoTo 0
    Next i

    m_bSurbrillance = True
End Sub

' =============================================================================
' EffacerSurbrillance - Deselectionne tous les elements
' =============================================================================
Sub EffacerSurbrillance(ByVal oModel As ModelReference)
    If Not m_bSurbrillance Then Exit Sub
    oModel.UnselectAllElements
    m_bSurbrillance = False
End Sub

' =============================================================================
' Configuration des criteres de scan selon les options
' =============================================================================
Private Sub ConfigurerCriteres(ByVal oSC As ElementScanCriteria, _
                               ByVal oSettings As CSettings)
    If oSettings.bLignes Then
        oSC.IncludeType msdElementTypeLine
        oSC.IncludeType msdElementTypeLineString
        oSC.IncludeType msdElementTypePointString
    End If
    If oSettings.bTextes Then
        oSC.IncludeType msdElementTypeText
        oSC.IncludeType msdElementTypeTextNode
    End If
    If oSettings.bArcs Then
        oSC.IncludeType msdElementTypeArc
        oSC.IncludeType msdElementTypeEllipse
    End If
    If oSettings.bCourbes Then
        oSC.IncludeType msdElementTypeCurve
        oSC.IncludeType msdElementTypeBsplineCurve
        oSC.IncludeType msdElementTypeBsplineSurface
    End If
    If oSettings.bFormes Then
        oSC.IncludeType msdElementTypeShape
        oSC.IncludeType msdElementTypeComplexString
        oSC.IncludeType msdElementTypeComplexShape
    End If
    If oSettings.bCellules Then
        oSC.IncludeType msdElementTypeCellHeader
    End If
    If oSettings.bDimensions Then
        oSC.IncludeType msdElementTypeDimension
    End If
    If oSettings.bMultiLignes Then
        oSC.IncludeType msdElementTypeMultiLine
    End If
End Sub

' =============================================================================
' Generation de la signature complete d'un element
' =============================================================================
Private Function GenererSignature(ByVal oElement As Element, _
                                  ByVal oSettings As CSettings) As String
    Dim sSig As String
    Dim sGeo As String
    Dim lVal As Long

    sSig = CStr(oElement.Type) & "|"

    On Error Resume Next

    If oSettings.bComparerNiveau Then
        lVal = -1
        lVal = oElement.Level.Number
        sSig = sSig & CStr(lVal) & "|"
    End If
    If oSettings.bComparerCouleur Then
        lVal = -1
        lVal = oElement.Color
        sSig = sSig & CStr(lVal) & "|"
    End If
    If oSettings.bComparerEpaisseur Then
        lVal = -1
        lVal = oElement.LineWeight
        sSig = sSig & CStr(lVal) & "|"
    End If
    If oSettings.bComparerStyle Then
        lVal = -1
        lVal = oElement.LineStyle.Number
        sSig = sSig & CStr(lVal) & "|"
    End If

    On Error GoTo 0

    sGeo = SignatureGeometrique(oElement)
    If Len(sGeo) = 0 Then
        GenererSignature = ""
    Else
        GenererSignature = sSig & sGeo
    End If
End Function

' =============================================================================
' Dispatch vers la bonne classe de signature selon le type
' =============================================================================
Private Function SignatureGeometrique(ByVal oElement As Element) As String
    Dim s As String
    s = ""

    On Error Resume Next

    Select Case oElement.Type
        Case msdElementTypeLine
            s = m_oSigLigne.SignatureLigne(oElement.AsLineElement)

        Case msdElementTypeLineString
            s = m_oSigLigne.SignatureLineString(oElement.AsLineElement)

        Case msdElementTypePointString
            s = m_oSigLigne.SignaturePointString(oElement.AsLineElement)

        Case msdElementTypeText
            s = m_oSigTexte.SignatureTexte(oElement.AsTextElement)

        Case msdElementTypeTextNode
            s = m_oSigTexte.SignatureTextNode(oElement.AsTextNodeElement)

        Case msdElementTypeArc
            s = m_oSigArc.SignatureArc(oElement.AsArcElement)

        Case msdElementTypeEllipse
            s = m_oSigArc.SignatureEllipse(oElement.AsEllipseElement)

        Case msdElementTypeCurve
            s = m_oSigCourbe.SignatureCurve(oElement.AsCurveElement)

        Case msdElementTypeBsplineCurve
            s = m_oSigCourbe.SignatureBSplineCurve(oElement.AsBsplineCurveElement)

        Case msdElementTypeBsplineSurface
            s = m_oSigCourbe.SignatureBSplineSurface(oElement)

        Case msdElementTypeShape
            s = m_oSigForme.SignatureShape(oElement.AsShapeElement)

        Case msdElementTypeComplexString
            s = m_oSigForme.SignatureComplexString(oElement.AsComplexStringElement)

        Case msdElementTypeComplexShape
            s = m_oSigForme.SignatureComplexShape(oElement.AsComplexShapeElement)

        Case msdElementTypeCellHeader
            s = m_oSigCell.SignatureCell(oElement.AsCellElement)

        Case msdElementTypeDimension
            s = m_oSigCell.SignatureDimension(oElement.AsDimensionElement)

        Case msdElementTypeMultiLine
            s = m_oSigCell.SignatureMultiLine(oElement.AsMultiLineElement)
    End Select

    On Error GoTo 0

    SignatureGeometrique = s
End Function
