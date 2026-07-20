# NettoyageDoublons - MicroStation V8i VBA

## Encodage obligatoire des fichiers VBA

Apres chaque creation ou modification de fichier `.cls`, `.bas` ou `.frm`,
normaliser en **CRLF + ANSI (Windows-1252)** avec :

    $t = [IO.File]::ReadAllText($path)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($path, $t, [Text.Encoding]::GetEncoding(1252))

Raison : le Write tool produit du LF/UTF-8. MicroStation importe les `.cls`
en LF comme des modules standard au lieu de classes, ce qui casse
`Implements` et `New`.

## Formulaire (FRM/FRX)

- Le formulaire est construit entierement au runtime dans
  `ConstruireControles` : le `.frx` est un blob binaire du formulaire vide.
- Si seul le **code** du `.frm` change : le `.frx` reste valide.
- Si les **proprietes du designer** changent (en-tete `Begin...End`) :
  regenerer le couple via `scripts\export_frm_via_excel.ps1`.

## Architecture

- Un metier = une classe (CSignatureLigne, CSignatureTexte, etc.)
- CMoteurScan orchestre le scan et delegue aux classes de signature
- CSettings : parametres sans dependance API MicroStation
- Le module .bas ne contient que le point d'entree et les globals

```
NettoyageDoublons()                    ' .bas : entree + globals
      |
      +--- CSettings (g_oSettings)     ' parametres (tolerance, filtres)
      |
      +--- CMoteurScan (g_oMoteur)     ' scan + detection doublons (API MST)
      |         +-- CSignatureLigne    '   lignes, linestrings, pointstrings
      |         +-- CSignatureTexte    '   textes, text nodes
      |         +-- CSignatureArc      '   arcs, ellipses
      |         +-- CSignatureCourbe   '   curves, bsplines
      |         +-- CSignatureForme    '   shapes, complex strings/shapes
      |         +-- CSignatureCell     '   cellules, dimensions, multilignes
      |
      +--- frmNettoyage                ' modeless, controles au runtime
```

## Lancement

Key-in MicroStation : `vba run [NettoyageDoublonsV1]NettoyageDoublons`

## Fichiers

```
src/
  NettoyageDoublonsV1.bas   - Module principal
  CSettings.cls             - Parametres
  CSignatureLigne.cls       - Signature lignes
  CSignatureTexte.cls       - Signature textes
  CSignatureArc.cls         - Signature arcs/ellipses
  CSignatureCourbe.cls      - Signature courbes/bsplines
  CSignatureForme.cls       - Signature shapes/complex
  CSignatureCell.cls        - Signature cellules/dimensions/multilignes
  CMoteurScan.cls           - Moteur de scan et orchestration
  frmNettoyage.frm          - Formulaire (code)
  frmNettoyage.frx          - Blob designer (binaire)
scripts/
  export_frm_via_excel.ps1  - Generation du couple .frm/.frx
```
