# Pointage de factures et pièces justificatives

Ce dépôt privé contient le **code de traitement uniquement**. Les documents comptables ne doivent jamais être ajoutés aux commits Git, aux Issues ou aux journaux du workflow.

## Objectif

Pour chaque ligne des classeurs de pointage, rechercher la pièce source la plus plausible à partir du numéro, des références, du partenaire, de la date et des montants. La colonne `Case à cocher` reçoit le **nom exact du fichier source** seulement si la correspondance franchit le seuil défini. Les cas incertains sont signalés au lieu d'être attribués arbitrairement.

- `P.MBS.xlsx` : feuille cible `Sheet1 (3)`
- `P.DISTRIMAX.xlsx` : feuille cible `Sheet1 (2)`
- Types acceptés : PDF et images JPG/JPEG/PNG/TIF/TIFF.
- Sorties : deux classeurs pointés et `rapport_pointage.csv`.

## Transfert privé temporaire

Le workflow manuel télécharge l'archive `pointage-input.zip` depuis la release privée portant le tag `factures-a-pointer`. L'archive doit contenir exactement cette structure à sa racine :

```
input/
  P.MBS.xlsx
  P.DISTRIMAX.xlsx
documents/
  (les 1 984 documents, sous-dossiers conservés si nécessaire)
```

Pour préparer l'archive sur Windows :
1. Créer un dossier de préparation, par exemple `pointage-input`.
2. Dans ce dossier, créer `input` et y placer les deux classeurs Excel avec leurs noms exacts.
3. Créer `documents` et y copier les pièces justificatives en conservant les sous-dossiers éventuels.
4. Ouvrir le dossier `pointage-input`, sélectionner les dossiers `input` et `documents`, puis les compresser en ZIP. Le fichier final doit s'appeler `pointage-input.zip` et contenir ces deux dossiers à sa racine, sans dossier parent supplémentaire.

Téléverser ensuite ce ZIP comme **asset d'une release privée** du dépôt, avec le tag `factures-a-pointer` et le nom d'asset `pointage-input.zip`. Ne pas l'ajouter au dépôt comme fichier ordinaire et ne pas le committer.

## Lancement et résultats

Dans GitHub : onglet **Actions** → **Pointage des factures** → **Run workflow**. Le traitement installe Tesseract avec les langues française et anglaise, puis indexe les pièces. Les résultats sont publiés dans l'artefact `resultats-pointage`, conservé 7 jours, puis téléchargeables depuis la page du workflow.

Les fichiers comptables sont temporaires sur le runner GitHub Actions, mais le runner et les artefacts ne doivent pas être considérés comme un coffre-fort absolu. Téléchargez les résultats et supprimez l'asset de la release et la release après vérification si vous ne souhaitez plus conserver l'archive sur GitHub. La suppression ultérieure ne garantit pas l'effacement de toutes les copies ou caches éventuels.

## Confidentialité technique

Les chemins `input/`, `documents/`, `output/` et les extensions de documents comptables sont exclus de Git par `.gitignore`. Vérifier malgré tout avant chaque commit qu'aucun document ou résultat comptable n'est ajouté au dépôt.
