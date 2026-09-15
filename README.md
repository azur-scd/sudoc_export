# SudocExport

Outil local (Python + interface Tkinter) de récupération de notices bibliographiques Sudoc à partir d’une liste de PPN, avec export Excel.

## Objectif

SudocExport permet de transformer une liste de PPN en tableau Excel exploitable (contrôle, enrichissement, rapprochement), sans écrire de code.

Le traitement conserve la logique “une entrée non vide = une ligne exportée”, dans l’ordre de la saisie.

## Périmètre fonctionnel

### Inclus
- Saisie manuelle de PPN (zone de texte) ou import d’un fichier texte.
- Traitement séquentiel des PPN via `https://www.sudoc.fr/{ppn}.xml`.
- Extraction bibliographique vers un fichier Excel `.xlsx`.
- Suivi de progression et journal de traitement.
- Sauvegardes automatiques intermédiaires (paramétrables) + export final manuel.

### Hors périmètre
- Recherche Sudoc par titre/auteur/ISBN/ISSN.
- Extraction d’exemplaires/localisations détaillées.
- Synchronisation SIGB (Koha, etc.).
- Export MARCXML/ISO2709 complet.
- Interface web / service centralisé.

## Prérequis

- Python 3.9+
- Dépendances Python :
  - `requests`
  - `openpyxl`

Installation :

```bash
pip install requests openpyxl
```

## Lancement

```bash
python sudoc_export.py
```

## Utilisation

1. Coller une liste de PPN dans la zone de saisie **ou** importer un fichier texte.
2. (Optionnel) Régler :
   - délai entre requêtes (ms),
   - fréquence de sauvegarde automatique (nombre d’entrées traitées).
3. Cliquer sur **Lancer le traitement**.
4. Suivre l’avancement (barre, PPN courant, journal).
5. Exporter le résultat en `.xlsx`.

## Format d’entrée

- Séparateurs acceptés : retours à la ligne, virgules, points-virgules (et tabulations côté import fichier).
- Les entrées vides sont ignorées.
- Les doublons sont conservés (donc exportés plusieurs fois).

## Colonnes exportées (18)

Feuille principale : **Notices Sudoc**

1. PPN  
2. Type de document  
3. Titre  
4. Mention de resp.  
5. Auteur(s)  
6. Éditeur  
7. Lieu d’édition  
8. Date de publication  
9. Date codée  
10. Pays d’édition  
11. Description physique  
12. Collection  
13. Langue  
14. ISBN  
15. EAN  
16. ISSN  
17. Lien Sudoc  
18. Statut  

## Règles d’extraction (résumé)

- Source : XML Sudoc récupéré via URL directe `{ppn}.xml`.
- Titre : basé sur la zone 200 (notamment sous-zones `a/e/h/i`).
- Responsabilités : zones 700/701/702 et collectivités 710/711.
- Adresse bibliographique : zones 214, avec repli possible sur 210 selon présence des sous-zones.
- Lien Sudoc : URL vers la notice.

> Les règles métier détaillées et arbitrages de consolidation sont documentés dans `docs/PRD_SudocExport.md`.

## Statuts et erreurs

Le traitement distingue au minimum :
- succès (`OK`),
- non-récupération (ex. HTTP 404),
- incidents réseau / timeout,
- erreurs XML / structure inattendue.

Le journal permet d’identifier les PPN en échec et la nature des erreurs.

## Sauvegardes

- Sauvegarde automatique périodique configurable (par défaut : toutes les 100 entrées traitées).
- Fichier autosauvegardé horodaté dans le dossier utilisateur (comportement actuel).
- Export final manuel via la boîte de dialogue d’enregistrement.

## Construction d’un exécutable Windows (optionnel)

Exemple avec PyInstaller :

```bash
pyinstaller --onefile --windowed sudoc_export.py
```

Le binaire est généré dans `dist/`.

## Limites connues

- Outil de bureau séquentiel (pas de parallélisation réseau).
- Qualité du résultat dépend des PPN fournis et de la disponibilité du service distant.
- Certaines règles bibliographiques nécessitent validation métier continue (voir PRD).

## Documentation projet

- PRD : `docs/PRD_SudocExport.md`
- Dépôt : https://github.com/azur-scd/sudoc_export
