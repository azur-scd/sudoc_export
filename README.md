# Sudoc PPN → Excel  
**Outil de récupération de notices MARCXML depuis le Sudoc**

---

## Utilisation directe (avec Python installé)

### 1. Installer les dépendances
```
pip install requests openpyxl
```

### 2. Lancer le programme
```
python sudoc_export.py
```

---

## Créer un .exe Windows autonome (pour diffusion)

### Prérequis
- Python 3.9+ installé sur le poste de compilation
- Les dépendances ci-dessus installées

### Étapes

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "SudocExport" sudoc_export.py
```

Le fichier `SudocExport.exe` est généré dans le dossier `dist/`.  
Il est **autonome** : aucune installation requise sur le poste du bibliothécaire.

---

## Mode d'emploi de l'outil

1. **Collez vos PPN** dans la zone de texte (un PPN par ligne).  
   Formats acceptés : séparés par retour à la ligne, virgule ou point-virgule.
2. Cliquez sur **▶ Lancer la récupération**.  
   Le journal affiche en temps réel l'avancement.
3. Une fois terminé, cliquez sur **💾 Exporter en Excel**.  
   Choisissez l'emplacement et le nom du fichier.

---

## Champs extraits dans le fichier Excel

| Colonne               | Source MARC       |
|-----------------------|-------------------|
| PPN                   | 001               |
| Type de document      | Leader pos. 6     |
| Titre                 | 200 $a            |
| Sous-titre            | 200 $e            |
| Mention de resp.      | 200 $f $g         |
| Auteur(s)             | 700/701/702/710/711 |
| Éditeur               | 214 $c ou 210 $c  |
| Lieu d'édition        | 214 $a ou 210 $a  |
| Date de publication   | 214 $d ou 210 $d  |
| Description physique  | 215 $a $c $d      |
| Collection            | 225 $a            |
| Langue                | 101 $a            |
| ISBN                  | 010 $a            |
| ISSN                  | 011 $a            |
| Lien Sudoc            | (construit)       |
| Statut                | OK / erreur       |

---

## Remarques

- Le programme interroge l'**API SRU publique** de l'ABES (pas de clé requise).  
- Les lignes en **rouge pâle** dans Excel indiquent un PPN introuvable ou une erreur réseau.  
- Une connexion Internet est nécessaire durant l'utilisation.
- Pour de très grandes listes (>500 PPN), prévoir quelques minutes de traitement.
