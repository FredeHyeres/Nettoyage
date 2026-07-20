VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSettings - Parametres du nettoyage
' Aucune dependance a l'API MicroStation
' =============================================================================

Public dTolerance       As Double
Public dAngleTolerance  As Double

' --- Types a scanner ---
Public bLignes          As Boolean
Public bTextes          As Boolean
Public bArcs            As Boolean
Public bCourbes         As Boolean
Public bFormes          As Boolean
Public bCellules        As Boolean
Public bDimensions      As Boolean
Public bMultiLignes     As Boolean

' --- Options ---
Public bComparerNiveau      As Boolean
Public bComparerCouleur     As Boolean
Public bComparerEpaisseur   As Boolean
Public bComparerStyle       As Boolean

Sub Init()
    dTolerance = 0.01
    dAngleTolerance = 0.01

    bLignes = True
    bTextes = True
    bArcs = True
    bCourbes = True
    bFormes = True
    bCellules = True
    bDimensions = True
    bMultiLignes = True

    bComparerNiveau = True
    bComparerCouleur = True
    bComparerEpaisseur = True
    bComparerStyle = True
End Sub
