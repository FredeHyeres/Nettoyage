Attribute VB_Name = "NettoyageDoublonsV1"
Option Explicit

' =============================================================================
' NettoyageDoublonsV1 - Suppression des elements en doublon dans un DGN
' MicroStation V8i SS3 - VBA 6.x
' Ne prend pas en compte les fichiers en reference
' =============================================================================

' --- Globals partages ---
Public g_oSettings      As CSettings
Public g_oMoteur        As CMoteurScan

'----------------------------------------------------------------------
Sub NettoyageDoublons()
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Nettoyage Doublons"
        Exit Sub
    End If

    Set g_oSettings = New CSettings
    g_oSettings.Init

    Set g_oMoteur = New CMoteurScan

    frmNettoyage.Initialiser g_oSettings
    frmNettoyage.Show vbModeless
End Sub

'----------------------------------------------------------------------
Sub LancerAnalyse()
    If g_oMoteur Is Nothing Then Exit Sub

    If g_oMoteur.NbDoublonsTrouves > 0 Then
        g_oMoteur.EffacerSurbrillance ActiveModelReference
    End If

    g_oMoteur.Analyser ActiveModelReference, g_oSettings

    If g_oMoteur.NbDoublonsTrouves > 0 Then
        g_oMoteur.SurlignierDoublons ActiveModelReference
    End If
    frmNettoyage.AfficherResultat g_oMoteur.NbElementsScannes, _
                                  g_oMoteur.NbDoublonsTrouves
End Sub

'----------------------------------------------------------------------
Sub LancerSuppression()
    If g_oMoteur Is Nothing Then Exit Sub
    If g_oMoteur.NbDoublonsTrouves = 0 Then Exit Sub
    g_oMoteur.EffacerSurbrillance ActiveModelReference
    g_oMoteur.SupprimerDoublons ActiveModelReference
    frmNettoyage.AfficherSuppression g_oMoteur.NbSupprimes
End Sub

'----------------------------------------------------------------------
Sub NettoyerAvantFermeture()
    If g_oMoteur Is Nothing Then Exit Sub
    g_oMoteur.EffacerSurbrillance ActiveModelReference
End Sub
