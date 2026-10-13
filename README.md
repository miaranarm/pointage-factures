# Pointage de factures et pièces justificatives

Ce dépôt contient le **code de traitement uniquement**. Ne jamais y ajouter les classeurs comptables, PDF, images, scans, exports OCR ou résultats contenant des données comptables.

## Objectif
Pour chaque ligne des feuilles de pointage, retrouver le fichier source le plus plausible à partir du numéro, des références, du partenaire, de la date, du montant HT signé, de la taxe signée et du total TTC signé. La colonne `Case à cocher` reçoit le **nom exact du fichier source** uniquement quand la correspondance est suffisamment fiable. Les ambiguïtés sont signalées et ne sont pas attribuées arbitrairement.

Aucun type de pièce n'est exclu : facture, note de débit, pro forma, bon de livraison, etc.

## Classeurs repérés
- `P.MBS.xlsx` : feuille cible à confirmer `Sheet1 (3)`, avec `Case à cocher`.
- `P.DISTRIMAX.xlsx` : feuille cible à confirmer `Sheet1 (2)`, avec `Case à cocher`.
Les feuilles `Sheet1` semblent être des feuilles distinctes sans colonne `Case à cocher`. Le traitement ne doit pas modifier les feuilles sources sans confirmation.

## Confidentialité — important
Le dépôt est public selon le réglage choisi. Le code ici est public. **Ne téléversez jamais les documents, classeurs, journaux OCR ou résultats comptables dans ce dépôt public**, ni dans les Issues, commits ou logs Actions.

## État actuel
Le moteur doit encore être raccordé à une source temporaire et privée pour les 1 984 documents (environ 1,08 Go). GitHub ne peut pas lire les fichiers qui restent sur votre ordinateur. Le workflow ne lancera pas le pointage complet tant que les fichiers n'auront pas été transmis au runner par un canal adapté. Ne pas stocker les pièces dans l'historique Git.

## Résultat attendu
- Deux classeurs de sortie préservant les feuilles et colonnes originales.
- `Case à cocher` = nom exact du fichier source lorsqu'une correspondance fiable est trouvée.
- Un rapport d'anomalies pour les lignes non pointées, ambiguës ou présentant des conflits.
- Les documents ne sont pas modifiés.
