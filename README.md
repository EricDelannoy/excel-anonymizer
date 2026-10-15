# excel-anonymizer

Script Python qui ouvre un classeur Excel (`.xlsx` / `.xlsm`), **supprime l'auteur et le dernier modificateur**, **détecte les données cachées et les onglets masqués**, et **supprime tout ce qui n'est pas nécessaire**.

Le fichier d'origine n'est jamais modifié (sortie : `<nom>.clean.xlsx`). Le traitement se fait directement sur le XML du classeur, sans passer par openpyxl : graphiques, images, tableaux croisés dynamiques et mises en forme sont conservés.

## Installation

```bash
pip install -r requirements.txt
```

Prérequis : Python 3.10+ et `lxml` 4.9.2 ou plus récent (testé avec lxml 4.9.2 sur Python 3.11, et avec lxml 6.x sur Python 3.13).

## Utilisation

```bash
python excel_anonymizer.py classeur.xlsx                  # -> classeur.clean.xlsx
python excel_anonymizer.py classeur.xlsx --audit-only     # rapport seul, rien n'est écrit
python excel_anonymizer.py classeur.xlsx -o propre.xlsx --author "Service Finance"   # inscrit un nom d'auteur au choix
```

Le script affiche un audit avant nettoyage, la liste des actions, puis un audit du fichier produit.

| Option | Effet |
|---|---|
| `-o, --output` | fichier de sortie |
| `--in-place` | écrase le fichier d'origine (déconseillé) |
| `-a, --author NOM` | inscrit ce nom comme auteur et dernier modificateur (et auteur des commentaires). **Par défaut, l'auteur est totalement supprimé** (champs vides). |
| `--audit-only` | affiche seulement le rapport |
| `--keep-hidden-sheets` | conserve les onglets masqués |
| `--remove-comments` | supprime aussi commentaires et notes |
| `--purge-hidden-cells` | vide le contenu des lignes/colonnes masquées (les lignes masquées par un filtre sont conservées) |
| `--keep-dates` | conserve les dates de création/modification |
| `--force` | remplace par leur valeur les formules qui dépendent d'un onglet masqué supprimé |

## Ce qui est anonymisé

- auteur et dernier modificateur (`docProps/core.xml`) : **champs laissés vides par défaut**, comme avec « Fichier > Informations > Inspecter le document » d'Excel (ou remplacés par le nom donné avec `--author`) ; dates de création/modification remplacées par la date du traitement ;
- société, responsable et liste des noms d'onglets (`docProps/app.xml`) ;
- auteurs des commentaires, personnes des commentaires modernes ;
- chemin absolu du fichier d'origine, nom d'utilisateur de partage ;
- sujet, mots-clés, description, catégorie.

Le classeur est aussi réglé pour **s'ouvrir sur la première feuille visible**, avec une seule feuille sélectionnée. Sur **chaque feuille**, le défilement est remis en haut et la cellule active en A1 (avec des volets figés, la vue est replacée à l'origine de chaque volet ; les volets fractionnés non figés ne sont pas modifiés).

### Comparaison avec « Inspecter le document » d'Excel

| Catégorie de l'inspecteur | Ce script |
|---|---|
| Propriétés du document et informations personnelles (auteur, dernier modificateur, société, responsable…) | oui (auteur et dernier modificateur vides par défaut) |
| Commentaires et annotations | auteurs vidés ; suppression avec `--remove-comments` |
| Feuilles de calcul masquées | oui (supprimées) |
| Lignes et colonnes masquées | signalées ; vidées avec `--purge-hidden-cells` |
| Contenu invisible (objets masqués) | oui |
| Données XML personnalisées | oui |
| En-têtes et pieds de page | non |
| Documents incorporés, données de modèle | signalés seulement |

Le titre du document n'est pas modifié.

## Ce qui est détecté (rapport)

- onglets masqués et « très masqués » ; lignes et colonnes masquées ; cellules au format `;;;` ;
- objets graphiques masqués ; noms définis masqués ou cassés ;
- formules d'onglets visibles qui dépendent d'onglets masqués ;
- macros VBA, liens externes (avec leur cible), caches de tableaux croisés dynamiques, connexions de données, objets incorporés ;
- noms propres retrouvés dans le contenu des cellules ou des commentaires.

## Ce qui est supprimé

- onglets masqués (et leurs commentaires, dessins, tableaux…) — **le texte correspondant est aussi purgé de `sharedStrings.xml`**, sinon il resterait lisible dans le fichier ;
- objets graphiques masqués, noms définis masqués/cassés inutilisés ;
- miniature, propriétés personnalisées, XML personnalisé, paramètres d'imprimante, chaîne de calcul, historique des modifications suivies ;
- fichiers internes devenus orphelins.

## Limites à connaître

- Les noms qui apparaissent **dans le texte des cellules** (« Contact : Jean ») ne sont pas modifiés : ils sont seulement signalés.
- Les macros VBA, liens externes et caches de tableaux croisés dynamiques sont **signalés mais conservés** (les supprimer changerait le comportement du classeur).
- Si des formules d'onglets visibles dépendent d'un onglet masqué, le script s'arrête par sécurité ; `--force` les remplace par leurs valeurs.
- Les dépendances depuis des mises en forme conditionnelles, validations de données ou graphiques vers un onglet masqué supprimé ne sont pas réécrites.
- Format `.xls` (ancien) non pris en charge : enregistrez-le d'abord en `.xlsx`.
- Les formats `.xlsb` et les fichiers protégés par mot de passe ne sont pas pris en charge.

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```
