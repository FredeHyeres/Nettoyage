VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNettoyage 
   Caption         =   "Nettoyage Doublons"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3600
   OleObjectBlob   =   "frmNettoyage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNettoyage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' frmNettoyage - Formulaire modeless pour le nettoyage des doublons
' Controles crees au runtime dans ConstruireControles
' =============================================================================

Private m_bConstruit    As Boolean
Private m_bInit         As Boolean
Private m_oSettings     As CSettings

' --- Controles declares WithEvents pour recevoir les evenements ---
Private WithEvents chkLignes        As MSForms.CheckBox
Attribute chkLignes.VB_VarHelpID = -1
Private WithEvents chkTextes        As MSForms.CheckBox
Attribute chkTextes.VB_VarHelpID = -1
Private WithEvents chkArcs          As MSForms.CheckBox
Attribute chkArcs.VB_VarHelpID = -1
Private WithEvents chkCourbes       As MSForms.CheckBox
Attribute chkCourbes.VB_VarHelpID = -1
Private WithEvents chkFormes        As MSForms.CheckBox
Attribute chkFormes.VB_VarHelpID = -1
Private WithEvents chkCellules      As MSForms.CheckBox
Attribute chkCellules.VB_VarHelpID = -1
Private WithEvents chkDimensions    As MSForms.CheckBox
Attribute chkDimensions.VB_VarHelpID = -1
Private WithEvents chkMultiLignes   As MSForms.CheckBox
Attribute chkMultiLignes.VB_VarHelpID = -1

Private WithEvents chkNiveau        As MSForms.CheckBox
Attribute chkNiveau.VB_VarHelpID = -1
Private WithEvents chkCouleur       As MSForms.CheckBox
Attribute chkCouleur.VB_VarHelpID = -1
Private WithEvents chkEpaisseur     As MSForms.CheckBox
Attribute chkEpaisseur.VB_VarHelpID = -1
Private WithEvents chkStyle         As MSForms.CheckBox
Attribute chkStyle.VB_VarHelpID = -1

Private WithEvents btnAnalyser      As MSForms.CommandButton
Attribute btnAnalyser.VB_VarHelpID = -1
Private WithEvents btnSupprimer     As MSForms.CommandButton
Attribute btnSupprimer.VB_VarHelpID = -1
Private WithEvents btnFermer        As MSForms.CommandButton
Attribute btnFermer.VB_VarHelpID = -1

Private WithEvents txtTolerance     As MSForms.TextBox
Attribute txtTolerance.VB_VarHelpID = -1
Private WithEvents txtAngleTol      As MSForms.TextBox
Attribute txtAngleTol.VB_VarHelpID = -1

Private lblResultat                 As MSForms.Label
Private lblStatus                   As MSForms.Label

' =============================================================================
Private Sub UserForm_Initialize()
    If Not m_bConstruit Then ConstruireControles
End Sub

' =============================================================================
Public Sub Initialiser(ByVal oSettings As CSettings)
    m_bInit = True
    Set m_oSettings = oSettings

    If Not m_bConstruit Then ConstruireControles

    chkLignes.Value = oSettings.bLignes
    chkTextes.Value = oSettings.bTextes
    chkArcs.Value = oSettings.bArcs
    chkCourbes.Value = oSettings.bCourbes
    chkFormes.Value = oSettings.bFormes
    chkCellules.Value = oSettings.bCellules
    chkDimensions.Value = oSettings.bDimensions
    chkMultiLignes.Value = oSettings.bMultiLignes

    chkNiveau.Value = oSettings.bComparerNiveau
    chkCouleur.Value = oSettings.bComparerCouleur
    chkEpaisseur.Value = oSettings.bComparerEpaisseur
    chkStyle.Value = oSettings.bComparerStyle

    txtTolerance.Text = Format$(oSettings.dTolerance, "0.0000")
    txtAngleTol.Text = Format$(oSettings.dAngleTolerance, "0.0000")

    lblResultat.Caption = ""
    lblStatus.Caption = "Pret."
    btnSupprimer.Enabled = False

    m_bInit = False
End Sub

' =============================================================================
Public Sub AfficherResultat(ByVal lScannes As Long, ByVal lDoublons As Long)
    lblResultat.Caption = "Scannes: " & lScannes & vbCrLf & _
                          "Doublons: " & lDoublons
    If lDoublons > 0 Then
        btnSupprimer.Enabled = True
        lblStatus.Caption = lDoublons & " doublons prets a supprimer."
    Else
        btnSupprimer.Enabled = False
        lblStatus.Caption = "Aucun doublon trouve."
    End If
End Sub

' =============================================================================
Public Sub AfficherSuppression(ByVal lSupprimes As Long)
    lblStatus.Caption = lSupprimes & " doublons supprimes."
    lblResultat.Caption = lblResultat.Caption & vbCrLf & _
                          "Supprimes: " & lSupprimes
    btnSupprimer.Enabled = False
End Sub

' =============================================================================
Private Sub ConstruireControles()
    Dim fraTypes As MSForms.Frame
    Dim fraOptions As MSForms.Frame
    Dim fraActions As MSForms.Frame
    Dim nTop As Long

    Me.Width = 240
    Me.Height = 430

    ' --- Frame : Types d'elements ---
    Set fraTypes = Me.Controls.Add("Forms.Frame.1", "fraTypes")
    fraTypes.Caption = "Types a scanner"
    fraTypes.Left = 6: fraTypes.Top = 6: fraTypes.Width = 222: fraTypes.Height = 138

    nTop = 14
    Set chkLignes = fraTypes.Controls.Add("Forms.CheckBox.1", "chkLignes")
    chkLignes.Caption = "Lignes / LineStrings"
    chkLignes.Left = 8: chkLignes.Top = nTop: chkLignes.Width = 140: chkLignes.Height = 14
    nTop = nTop + 14

    Set chkTextes = fraTypes.Controls.Add("Forms.CheckBox.1", "chkTextes")
    chkTextes.Caption = "Textes / TextNodes"
    chkTextes.Left = 8: chkTextes.Top = nTop: chkTextes.Width = 140: chkTextes.Height = 14
    nTop = nTop + 14

    Set chkArcs = fraTypes.Controls.Add("Forms.CheckBox.1", "chkArcs")
    chkArcs.Caption = "Arcs / Ellipses"
    chkArcs.Left = 8: chkArcs.Top = nTop: chkArcs.Width = 140: chkArcs.Height = 14
    nTop = nTop + 14

    Set chkCourbes = fraTypes.Controls.Add("Forms.CheckBox.1", "chkCourbes")
    chkCourbes.Caption = "Courbes / B-Splines"
    chkCourbes.Left = 8: chkCourbes.Top = nTop: chkCourbes.Width = 140: chkCourbes.Height = 14
    nTop = nTop + 14

    Set chkFormes = fraTypes.Controls.Add("Forms.CheckBox.1", "chkFormes")
    chkFormes.Caption = "Shapes / Complex"
    chkFormes.Left = 8: chkFormes.Top = nTop: chkFormes.Width = 140: chkFormes.Height = 14
    nTop = nTop + 14

    Set chkCellules = fraTypes.Controls.Add("Forms.CheckBox.1", "chkCellules")
    chkCellules.Caption = "Cellules"
    chkCellules.Left = 8: chkCellules.Top = nTop: chkCellules.Width = 140: chkCellules.Height = 14
    nTop = nTop + 14

    Set chkDimensions = fraTypes.Controls.Add("Forms.CheckBox.1", "chkDimensions")
    chkDimensions.Caption = "Dimensions"
    chkDimensions.Left = 8: chkDimensions.Top = nTop: chkDimensions.Width = 140: chkDimensions.Height = 14
    nTop = nTop + 14

    Set chkMultiLignes = fraTypes.Controls.Add("Forms.CheckBox.1", "chkMultiLignes")
    chkMultiLignes.Caption = "MultiLignes"
    chkMultiLignes.Left = 8: chkMultiLignes.Top = nTop: chkMultiLignes.Width = 140: chkMultiLignes.Height = 14

    ' --- Frame : Criteres de comparaison ---
    Set fraOptions = Me.Controls.Add("Forms.Frame.1", "fraOptions")
    fraOptions.Caption = "Criteres de comparaison"
    fraOptions.Left = 6: fraOptions.Top = 148: fraOptions.Width = 222: fraOptions.Height = 126

    nTop = 14
    Set chkNiveau = fraOptions.Controls.Add("Forms.CheckBox.1", "chkNiveau")
    chkNiveau.Caption = "Niveau (Level)"
    chkNiveau.Left = 8: chkNiveau.Top = nTop: chkNiveau.Width = 100: chkNiveau.Height = 14

    Set chkCouleur = fraOptions.Controls.Add("Forms.CheckBox.1", "chkCouleur")
    chkCouleur.Caption = "Couleur"
    chkCouleur.Left = 112: chkCouleur.Top = nTop: chkCouleur.Width = 100: chkCouleur.Height = 14
    nTop = nTop + 14

    Set chkEpaisseur = fraOptions.Controls.Add("Forms.CheckBox.1", "chkEpaisseur")
    chkEpaisseur.Caption = "Epaisseur"
    chkEpaisseur.Left = 8: chkEpaisseur.Top = nTop: chkEpaisseur.Width = 100: chkEpaisseur.Height = 14

    Set chkStyle = fraOptions.Controls.Add("Forms.CheckBox.1", "chkStyle")
    chkStyle.Caption = "Style"
    chkStyle.Left = 112: chkStyle.Top = nTop: chkStyle.Width = 100: chkStyle.Height = 14
    nTop = nTop + 20

    Dim lblTol As MSForms.Label
    Set lblTol = fraOptions.Controls.Add("Forms.Label.1", "lblTol")
    lblTol.Caption = "Tol. distance :"
    lblTol.Left = 8: lblTol.Top = nTop + 2: lblTol.Width = 80: lblTol.Height = 14

    Set txtTolerance = fraOptions.Controls.Add("Forms.TextBox.1", "txtTolerance")
    txtTolerance.Left = 92: txtTolerance.Top = nTop: txtTolerance.Width = 60: txtTolerance.Height = 18
    nTop = nTop + 22

    Dim lblAngTol As MSForms.Label
    Set lblAngTol = fraOptions.Controls.Add("Forms.Label.1", "lblAngTol")
    lblAngTol.Caption = "Tol. angle :"
    lblAngTol.Left = 8: lblAngTol.Top = nTop + 2: lblAngTol.Width = 80: lblAngTol.Height = 14

    Set txtAngleTol = fraOptions.Controls.Add("Forms.TextBox.1", "txtAngleTol")
    txtAngleTol.Left = 92: txtAngleTol.Top = nTop: txtAngleTol.Width = 60: txtAngleTol.Height = 18

    ' --- Frame : Actions ---
    Set fraActions = Me.Controls.Add("Forms.Frame.1", "fraActions")
    fraActions.Caption = "Actions"
    fraActions.Left = 6: fraActions.Top = 278: fraActions.Width = 222: fraActions.Height = 98

    Set btnAnalyser = fraActions.Controls.Add("Forms.CommandButton.1", "btnAnalyser")
    btnAnalyser.Caption = "Analyser"
    btnAnalyser.Left = 8: btnAnalyser.Top = 16: btnAnalyser.Width = 80: btnAnalyser.Height = 22

    Set btnSupprimer = fraActions.Controls.Add("Forms.CommandButton.1", "btnSupprimer")
    btnSupprimer.Caption = "Supprimer"
    btnSupprimer.Left = 96: btnSupprimer.Top = 16: btnSupprimer.Width = 80: btnSupprimer.Height = 22
    btnSupprimer.Enabled = False

    Set lblResultat = fraActions.Controls.Add("Forms.Label.1", "lblResultat")
    lblResultat.Caption = ""
    lblResultat.Left = 8: lblResultat.Top = 44: lblResultat.Width = 200: lblResultat.Height = 40
    lblResultat.WordWrap = True

    ' --- Barre de statut en bas ---
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    lblStatus.Caption = "Pret."
    lblStatus.Left = 6: lblStatus.Top = 382: lblStatus.Width = 222: lblStatus.Height = 14

    ' --- Bouton Fermer ---
    Set btnFermer = Me.Controls.Add("Forms.CommandButton.1", "btnFermer")
    btnFermer.Caption = "Fermer"
    btnFermer.Left = 150: btnFermer.Top = 382: btnFermer.Width = 78: btnFermer.Height = 20

    m_bConstruit = True
End Sub

' =============================================================================
' Evenements des checkboxes - Types
' =============================================================================
Private Sub chkLignes_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bLignes = chkLignes.Value
End Sub

Private Sub chkTextes_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bTextes = chkTextes.Value
End Sub

Private Sub chkArcs_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bArcs = chkArcs.Value
End Sub

Private Sub chkCourbes_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bCourbes = chkCourbes.Value
End Sub

Private Sub chkFormes_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bFormes = chkFormes.Value
End Sub

Private Sub chkCellules_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bCellules = chkCellules.Value
End Sub

Private Sub chkDimensions_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bDimensions = chkDimensions.Value
End Sub

Private Sub chkMultiLignes_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bMultiLignes = chkMultiLignes.Value
End Sub

' =============================================================================
' Evenements des checkboxes - Criteres
' =============================================================================
Private Sub chkNiveau_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bComparerNiveau = chkNiveau.Value
End Sub

Private Sub chkCouleur_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bComparerCouleur = chkCouleur.Value
End Sub

Private Sub chkEpaisseur_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bComparerEpaisseur = chkEpaisseur.Value
End Sub

Private Sub chkStyle_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bComparerStyle = chkStyle.Value
End Sub

' =============================================================================
' Evenements des TextBox - Tolerances
' =============================================================================
Private Sub txtTolerance_Change()
    If m_bInit Then Exit Sub
    Dim d As Double
    If IsNumeric(txtTolerance.Text) Then
        d = CDbl(txtTolerance.Text)
        If d > 0 Then m_oSettings.dTolerance = d
    End If
End Sub

Private Sub txtTolerance_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTolerance.Text = Format$(m_oSettings.dTolerance, "0.0000")
    End If
End Sub

Private Sub txtAngleTol_Change()
    If m_bInit Then Exit Sub
    Dim d As Double
    If IsNumeric(txtAngleTol.Text) Then
        d = CDbl(txtAngleTol.Text)
        If d > 0 Then m_oSettings.dAngleTolerance = d
    End If
End Sub

Private Sub txtAngleTol_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtAngleTol.Text = Format$(m_oSettings.dAngleTolerance, "0.0000")
    End If
End Sub

' =============================================================================
' Boutons d'action
' =============================================================================
Private Sub btnAnalyser_Click()
    btnAnalyser.Enabled = False
    lblStatus.Caption = "Analyse en cours..."
    lblResultat.Caption = ""
    btnSupprimer.Enabled = False
    DoEvents
    LancerAnalyse
    btnAnalyser.Enabled = True
End Sub

Private Sub btnSupprimer_Click()
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Confirmer la suppression des doublons ?", _
                 vbYesNo + vbQuestion, "Nettoyage Doublons")
    If rep = vbYes Then
        lblStatus.Caption = "Suppression en cours..."
        DoEvents
        LancerSuppression
    End If
End Sub

Private Sub btnFermer_Click()
    NettoyerAvantFermeture
    Me.Hide
    CommandState.StartDefaultCommand
End Sub

' =============================================================================
' Fermeture par la croix
' =============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        NettoyerAvantFermeture
        Me.Hide
        CommandState.StartDefaultCommand
    End If
End Sub

