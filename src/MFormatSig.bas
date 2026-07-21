Attribute VB_Name = "MFormatSig"
Option Explicit

' =============================================================================
' MFormatSig - Fonctions de formatage partagees par les classes CSignature*
'
' Regroupe FmtD / FormatPt / BoolStr, jusqu'ici copiees a l'identique dans les
' six classes de signature. La tolerance de quantification est passee en
' parametre (elle est commune a un scan : CSettings.dTolerance). Fonctions pures,
' sans etat.
' =============================================================================

' Quantifie d sur la grille de pas dTol, formate a 4 decimales.
Public Function FmtD(ByVal d As Double, ByVal dTol As Double) As String
    FmtD = Format$(Round(d / dTol) * dTol, "0.0000")
End Function

' Formate un point (X,Y,Z) quantifie.
Public Function FormatPt(ByRef pt As Point3d, ByVal dTol As Double) As String
    FormatPt = FmtD(pt.X, dTol) & "," & FmtD(pt.Y, dTol) & "," & FmtD(pt.Z, dTol)
End Function

' "T" / "F" pour un booleen.
Public Function BoolStr(ByVal b As Boolean) As String
    If b Then BoolStr = "T" Else BoolStr = "F"
End Function

' Serialise une rotation (Matrix3d) par ses vecteurs RowX et RowY quantifies.
' Deux orientations identiques -> meme chaine ; suffit a la deduplication et
' evite de traiter la Matrix3d comme un scalaire (piege corrige).
Public Function FmtRot(ByRef rot As Matrix3d, ByVal dTol As Double) As String
    FmtRot = FmtD(rot.RowX.X, dTol) & "," & FmtD(rot.RowX.Y, dTol) & "," & _
             FmtD(rot.RowX.Z, dTol) & "," & _
             FmtD(rot.RowY.X, dTol) & "," & FmtD(rot.RowY.Y, dTol) & "," & _
             FmtD(rot.RowY.Z, dTol)
End Function
