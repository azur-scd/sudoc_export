# PRD — SudocExport

**Récupération de données bibliographiques à partir de PPN et export Excel**

| Élément | Valeur |
| --- | --- |
| Version du document | 1.0 — proposition de cadrage |
| Date | 15 septembre 2026 |
| Sources du dépôt | `README.md` et `sudoc_export.py` |
| Produit | Application graphique locale en Python, diffusion Windows envisagée |
| Périmètre de ce PRD | Description de l’existant et spécification d’une version consolidée |
| Statut des décisions | Les exigences nouvelles et les arbitrages proposés restent à valider par le responsable du projet |

Ce document s’appuie sur la lecture intégrale du script et du README à la révision indiquée. Les mentions **existant** décrivent le code ; les mentions **cible** définissent des comportements proposés. Une fonction présente dans le code n’est pas nécessairement validée en conditions réelles.

Des vérifications ciblées ont été réalisées avec des réponses HTTP simulées et un classeur temporaire. Elles ont confirmé deux défauts : la lecture incorrecte d’un MARCXML comportant un espace de noms et l’écriture d’une chaîne commençant par `=` comme formule Excel. La connexion réelle au service Sudoc n’a pas pu être validée dans l’environnement d’analyse. L’interface graphique et l’exécutable Windows n’ont pas été testés. Aucun changement n’a été apporté au dépôt.

## 1. Vision du produit et problème à résoudre

**SudocExport doit permettre à un bibliothécaire de transformer une liste de PPN déjà constituée en tableau Excel exploitable, sans écrire de code ni consulter manuellement chaque notice.**

Une liste de PPN issue d’un SIGB, d’un travail sur les collections ou d’un rapprochement de fichiers ne contient pas toujours les informations nécessaires à son exploitation : titre, responsabilités, éditeur, date, langue ou identifiants commerciaux. La consultation successive des notices prend du temps et rend les recopies hétérogènes.

Le produit automatise cette récupération et produit une ligne traçable pour chaque entrée. Sa valeur repose autant sur l’extraction des informations que sur la capacité à distinguer les résultats utilisables des données absentes, des erreurs de saisie et des échecs techniques.

**Positionnement :** outil ponctuel d’enrichissement et de contrôle de listes bibliographiques pour un usage de bureau. Le fichier Excel est une sélection aplatie de métadonnées ; il ne constitue pas un export intégral et réversible des notices UNIMARC.

## 2. Utilisateurs et cas d’usage

Les profils et usages suivants sont proposés à partir du fonctionnement du script ; ils ne résultent pas d’entretiens utilisateurs.

| Profil | Besoin | Résultat attendu |
| --- | --- | --- |
| Bibliothécaire travaillant sur les collections | Documenter une liste de PPN extraite d’un outil local | Tableau lisible, filtrable, rapprochable avec la liste d’origine |
| Agent chargé d’un contrôle bibliographique | Repérer les notices récupérées, incomplètes ou indisponibles | Statuts explicites et liens vers le Sudoc |
| Référent données ou administrateur de SIGB | Préparer un travail de comparaison ou d’analyse | Identifiants conservés en texte, ordre préservé, règles d’extraction documentées |
| Mainteneur du script | Corriger ou faire évoluer l’outil | Installation reproductible, fonctions testables et jeux d’essai représentatifs |

### Cas d’usage prioritaires

1. **Enrichir une courte liste.** L’utilisateur colle des PPN, lance le traitement puis enregistre le résultat en Excel.
2. **Traiter un fichier.** L’utilisateur importe un fichier texte contenant quelques centaines ou quelques milliers d’identifiants, règle éventuellement le rythme des requêtes et suit l’avancement.
3. **Comprendre les échecs.** L’utilisateur filtre les statuts pour distinguer les erreurs de saisie, les notices non récupérées et les incidents réseau.
4. **Conserver un traitement partiel.** Une interruption ne doit pas rendre inutilisables les notices déjà récupérées.

## 3. Objectifs et indicateurs de réussite

| Objectif | Indicateur proposé | Critère d’acceptation |
| --- | --- | --- |
| Fiabiliser l’extraction | Comparaison à un corpus XML de référence validé métier | Résultats attendus obtenus pour tous les cas du corpus |
| Éviter les faux succès | Notices sans identifiant exploitable classées `OK` | Aucun cas dans les tests de recette |
| Assurer la traçabilité | Correspondance entre entrées et lignes exportées | Une ligne par entrée non vide, dans le même ordre, selon la politique retenue |
| Préserver les identifiants | Zéros initiaux, clé `X`, ISBN/EAN longs | Conservation exacte en texte après réouverture du fichier |
| Réduire le travail manuel | Réalisation du parcours principal par un agent | Export réussi sans intervention du mainteneur après lecture du mode d’emploi |
| Préserver les résultats | Sauvegarde finale et arrêt contrôlé | Dernier état récupérable et statut du traitement explicite |
| Permettre la diffusion | Exécution sur un poste Windows de recette sans Python | Lancement, récupération et export réussis |

Le taux de récupération dépend aussi de la qualité des PPN et du service distant. Il doit être mesuré séparément des erreurs propres à l’application. Aucun engagement chiffré de disponibilité ou de durée du service Sudoc n’est posé dans ce PRD.

## 4. État actuel du produit

### 4.1. Parcours existant

1. Ouverture d’une fenêtre Tkinter en français.
2. Collage de PPN ou import d’un fichier texte remplaçant le contenu de la zone de saisie.
3. Paramétrage facultatif du délai entre requêtes et de la fréquence des sauvegardes.
4. Lancement d’un traitement séquentiel dans un thread secondaire.
5. Affichage d’une progression, du PPN en cours et d’un journal.
6. Sauvegardes Excel intermédiaires aux seuils configurés.
7. Bilan des notices marquées `OK` et des erreurs.
8. Export manuel du résultat dans le fichier choisi.

### 4.2. Capacités et limites observées

| Fonction | Comportement existant |
| --- | --- |
| Saisie | Découpage sur retours à la ligne, virgules et points-virgules ; suppression des entrées vides |
| Import texte | Lecture UTF-8 puis repli Latin-1 ; tabulations acceptées en plus des séparateurs de la saisie |
| Validation des PPN | Aucune validation de structure ou de clé ; les tokens non vides sont interrogés |
| Doublons | Conservés, interrogés plusieurs fois et exportés plusieurs fois |
| Source distante | Requête GET sur `https://www.sudoc.fr/{ppn}.xml` |
| Accès | Aucun mécanisme de clé API ou d’authentification dans le script |
| Délai | 500 ms par défaut entre deux requêtes ; 0 possible |
| Timeout | `requests.get(..., timeout=15)` ; ce réglage n’est pas une garantie de durée totale maximale de 15 secondes |
| Reprises réseau | Aucune nouvelle tentative automatique ni réutilisation explicite d’une session HTTP |
| Traitement XML | Recherche d’un élément `record`, puis extraction de ses champs |
| Export | Classeur `.xlsx` à 18 colonnes, feuille « Notices Sudoc » |
| Présentation Excel | En-tête bleu, alternance de lignes, erreurs en rouge pâle, liens cliquables, filtres, première ligne figée |
| Sauvegarde automatique | Toutes les 100 entrées traitées par défaut ; fichier horodaté dans le dossier personnel ; écrasement du même fichier à chaque seuil |
| Fin du traitement | Pas de sauvegarde automatique finale si le dernier résultat ne tombe pas sur un seuil |
| Arrêt et reprise | Aucun bouton d’arrêt ; aucune reprise de session après fermeture |
| Effacement | Le bouton « Effacer » reste disponible pendant le traitement |
| Diffusion | Commande PyInstaller décrite dans le README ; aucun exécutable suivi dans l’arbre étudié |

Source : [script à la révision étudiée](https://github.com/azur-scd/sudoc_export/blob/0355ce118f31e47c16bc1adfe21bede6525fb7a7/sudoc_export.py).

### 4.3. Écarts entre README et code

| Sujet | README | Code |
| --- | --- | --- |
| Type de document | Présente seulement la position 6 du leader | Combine positions 6 et 7, puis affine avec `105$a` |
| Description physique | Annonce `215$a$c$d` | Extrait uniquement `215$a` |
| Import et réglages | Documentation partielle | Import texte, délai réglable et sauvegardes périodiques présents |

Le README devra être aligné sur le comportement validé de la version livrée. Source : [README étudié](https://github.com/azur-scd/sudoc_export/blob/0355ce118f31e47c16bc1adfe21bede6525fb7a7/README.md).

## 5. Périmètre de la version consolidée

### Inclus dans la cible

- Application locale en français, destinée en priorité à Windows.
- Saisie et import de listes de PPN.
- Normalisation contrôlée, validation et compte rendu des entrées.
- Récupération séquentielle des XML et extraction des 18 colonnes actuelles.
- Gestion fiable des erreurs, progression et arrêt contrôlé.
- Export Excel, sauvegardes intermédiaires et sauvegarde finale.
- Traçabilité du PPN demandé et du PPN retourné.
- Documentation des règles bibliographiques et recette sur corpus.

### Hors périmètre initial

- Recherche de notices à partir d’un titre, d’un auteur, d’un ISBN ou d’un ISSN.
- Recherche par RCR, ILN ou établissement et extraction exhaustive d’un catalogue.
- Données d’exemplaires, états de collection, disponibilité et localisation.
- Modification du Sudoc, import automatique dans Koha ou synchronisation avec un SIGB.
- Export intégral MARCXML ou ISO 2709.
- Interface web, comptes utilisateurs, service centralisé ou planification des traitements.
- Enrichissement automatique par IA et correction des notices sources.

Ces fonctions peuvent constituer d’autres projets ; elles ne sont pas nécessaires au parcours principal.

## 6. Exigences fonctionnelles

**P0 :** indispensable avant diffusion d’une version fiable. **P1 :** nécessaire à la consolidation pour un usage régulier. **P2 :** évolution facultative. Les priorités ci-dessous sont des propositions.

### 6.1. Entrée et préparation du lot

| ID | Priorité | Exigence cible | Situation actuelle |
| --- | --- | --- | --- |
| ENT-01 | P0 | Accepter les mêmes séparateurs en collage et en import : retour à la ligne, virgule, point-virgule et tabulation | Partiel |
| ENT-02 | P0 | Conserver les identifiants sous forme de chaînes ; retirer les espaces extérieurs et un éventuel BOM ; normaliser un `x` final en `X` | Partiel |
| ENT-03 | P0 | Contrôler la structure attendue du PPN avant toute requête ; ne pas compléter arbitrairement les zéros manquants | Absent |
| ENT-04 | P1 | Ajouter le contrôle de clé après validation de l’algorithme et de ses exemples auprès d’une référence ABES | Absent |
| ENT-05 | P0 | Conserver une ligne pour chaque entrée non vide, y compris invalide, dans l’ordre initial | Pas de classification des invalides |
| ENT-06 | P1 | Afficher les nombres d’entrées, d’identifiants distincts, de doublons et d’entrées invalides | Absent |
| ENT-07 | P1 | Accepter les fichiers UTF-8 avec ou sans BOM et UTF-16 avec BOM ; signaler les encodages non reconnus | Repli Latin-1 insuffisant |
| ENT-08 | P1 | Signaler que l’import remplace la saisie lorsque celle-ci contient déjà des données | Remplacement direct |

**Politique de doublons proposée :** conserver les occurrences dans l’export pour permettre un rapprochement ligne à ligne. Dans un même traitement, réutiliser le résultat obtenu pour un PPN déjà traité, sans nouvelle requête ; la relance explicite des échecs pourra renouveler les requêtes. Les compteurs distinguent entrées et requêtes.

Le fichier texte est une liste d’identifiants, sans colonne supplémentaire ni en-tête. Un import CSV ou Excel avec sélection de colonne relève de P2.

### 6.2. Récupération et validation du XML

| ID | Priorité | Exigence cible | Situation actuelle |
| --- | --- | --- | --- |
| REC-01 | P0 | Interroger l’URL directe Sudoc pour chaque identifiant valide ; documenter exactement ce service | Présent, README incorrect |
| REC-02 | P0 | Lire le MARCXML avec espace de noms par défaut, préfixé, ou sans espace de noms dans les fichiers de test | Défaillant pour les champs qualifiés |
| REC-03 | P0 | Accepter un `record` racine ou une enveloppe contenant une notice ; signaler une absence ou une ambiguïté de notices | Prend le premier `record` trouvé |
| REC-04 | P0 | Contrôler la présence de `001` et ne jamais annoncer `OK` pour un XML dont les données bibliographiques n’ont pas été reconnues | Absent |
| REC-05 | P0 | Conserver séparément PPN demandé et PPN retourné ; signaler une différence sans conclure automatiquement à une fusion | Absent |
| REC-06 | P0 | Distinguer erreur de saisie, 404, autre erreur HTTP, timeout, réseau, XML invalide et structure inattendue | Partiel |
| REC-07 | P1 | Réessayer les incidents transitoires avec une limite ; tenir compte de `Retry-After` lorsque pertinent | Absent |
| REC-08 | P1 | Maintenir un rythme séquentiel configurable et réutiliser une session HTTP | Délai présent, session absente |
| REC-09 | P0 | Une exception sur une notice ne doit pas laisser l’interface bloquée ; restaurer un état cohérent et conserver les résultats acquis | Protection globale absente |

**Réglages initiaux proposés :** délai de 500 ms conservé ; au maximum deux nouvelles tentatives pour un incident transitoire ; attentes de 2 puis 5 secondes en l’absence de consigne du serveur. Ces valeurs sont des choix applicatifs à confirmer selon les recommandations du service, et non des limites officielles attribuées à l’ABES. Les erreurs permanentes, comme un PPN invalide ou une réponse 404, ne sont pas relancées automatiquement.

### 6.3. Traitement et interface

| ID | Priorité | Exigence cible | Situation actuelle |
| --- | --- | --- | --- |
| IHM-01 | P0 | Afficher progression, PPN en cours, compteurs et bilan sans bloquer l’interface | Présent, à sécuriser |
| IHM-02 | P0 | Effectuer les mises à jour Tkinter depuis le thread principal | Nombreux appels depuis le thread de travail |
| IHM-03 | P0 | Empêcher un effacement ou un remplacement de lot de modifier les résultats d’un traitement actif | Effacer et import restent disponibles |
| IHM-04 | P1 | Proposer « Arrêter » ; ne plus démarrer de nouvelle requête et conserver les résultats déjà obtenus | Absent |
| IHM-05 | P1 | À l’arrêt, qualifier les entrées restantes « Non traité — arrêt utilisateur » dans l’export du lot | Absent |
| IHM-06 | P0 | Protéger les résultats non enregistrés avant effacement, nouveau traitement ou fermeture | Absent |
| IHM-07 | P1 | Exporter après un arrêt ou un incident global, dès qu’un résultat ou un statut peut être fourni | Export disponible à la fin normale |
| IHM-08 | P1 | Permettre la navigation clavier et rendre tous les statuts compréhensibles sans dépendre des couleurs ou des pictogrammes | À recetter |

Un arrêt ne promet pas l’interruption instantanée de la requête en cours. L’interface doit rester utilisable pendant son achèvement ou son expiration. Le lot actif et ses réglages sont figés au lancement ; la saisie suivante ne doit pas affecter ce lot.

### 6.4. Sauvegarde et export

| ID | Priorité | Exigence cible | Situation actuelle |
| --- | --- | --- | --- |
| EXP-01 | P0 | Conserver les 18 colonnes actuelles, leur ordre et leurs libellés dans la feuille principale | Présent |
| EXP-02 | P0 | Écrire toutes les valeurs bibliographiques comme texte littéral, y compris celles commençant par `=` | Défaut confirmé |
| EXP-03 | P0 | Produire un classeur ouvrable, avec liens Sudoc, filtres, ligne d’en-tête figée et statuts lisibles | Présent, recette à compléter |
| EXP-04 | P0 | Avec l’autosauvegarde activée, écrire le dernier état à la fin normale et lors d’un arrêt contrôlé, même sous le seuil périodique | Absent |
| EXP-05 | P0 | Éviter qu’une écriture interrompue ne détruise la dernière sauvegarde valide | Écriture directe |
| EXP-06 | P0 | Signaler clairement un dossier non accessible ou un fichier verrouillé ; permettre un autre emplacement sans perdre les données en mémoire | Gestion d’erreurs partielle |
| EXP-07 | P1 | Permettre le choix du dossier des sauvegardes et afficher son emplacement | Dossier personnel imposé |
| EXP-08 | P1 | Ajouter une feuille « Traitement » avec contexte, compteurs et correspondance entre entrées et résultats | Absent |
| EXP-09 | P2 | Réouvrir une session sauvegardée et reprendre les PPN restants | Absent |

L’autosauvegarde reste réglable : 100 entrées traitées par défaut, 0 pour la désactiver. L’interface doit expliciter qu’avec 0, seul un enregistrement manuel conserve le résultat. Une sauvegarde périodique seule peut perdre les résultats acquis depuis le dernier seuil lors d’un arrêt brutal ; elle ne constitue pas une reprise de session.

## 7. Dictionnaire des données exportées

Ce tableau décrit **exactement l’extraction actuelle** et les points à consolider. Toutes les valeurs bibliographiques sont destinées à être exportées en texte. Un champ absent reste vide ; il ne devient ni zéro ni donnée inventée.

| N° | Colonne | Source et règle actuelles | Exigence de consolidation |
| --- | --- | --- | --- |
| 1 | PPN | `001` ; PPN demandé sur les lignes d’erreur | Conserver aussi le PPN demandé dans la traçabilité ; aucune perte en cas de `001` absent |
| 2 | Type de document | Leader positions 6 et 7 ; surcharge par `105$a[4]` si `m` ou `7` | Vérifier les libellés et les cas de repli ; ne pas promettre une détection exhaustive du support électronique |
| 3 | Titre | `200$a$e$h$i`, dans l’ordre du XML ; ponctuation reconstruite | Traiter les répétitions et éviter la concaténation de deux `$a` sans séparateur |
| 4 | Mention de resp. | `200$f$g` | Préserver l’ordre pertinent et documenter la ponctuation |
| 5 | Auteur(s) | `700/701/702`, puis `710/711`, sous-champs `$a$b` | Signaler que la colonne contient aussi d’autres responsabilités ; étudier `712` et les rôles en P2 |
| 6 | Éditeur | Tous les `214$c` si présents, sinon `210$c` | Distinguer publication, production, diffusion et fabrication |
| 7 | Lieu d’édition | `214$a` si le test sur `214$c` réussit, sinon `210$a` | Ne pas supprimer un lieu présent en 214 parce que `$c` est absent |
| 8 | Date de publication | `214$d` selon la même condition, sinon `210$d` | Ne pas supprimer une date présente en 214 parce que `$c` est absent ; qualifier les dates de repli |
| 9 | Date codée | Première `100$a`, tranche Python `[9:13]`, sans nettoyage préalable | Conserver la valeur textuelle ; ce n’est pas une validation d’année |
| 10 | Pays d’édition | `102$a` | Conserver les codes ; libellés éventuels dans une évolution séparée |
| 11 | Description physique | `215$a` uniquement | Maintenir ce périmètre ou décider explicitement d’ajouter `$c$d` |
| 12 | Collection | `225$a` | Conserver les occurrences ; pas de numérotation ou d’ISSN de collection dans l’existant |
| 13 | Langue | `101$a` | Conserver les codes et les répétitions |
| 14 | ISBN | `010$a` | Conserver les graphies et occurrences en texte |
| 15 | EAN | `073$a` | Conserver en texte ; aucun calcul ou déduction à partir de l’ISBN |
| 16 | ISSN | `011$a` | Conserver les graphies et occurrences en texte |
| 17 | Lien Sudoc | `https://www.sudoc.fr/` suivi du `001` | Lien vers le PPN retourné, ou vers le PPN demandé s’il est valide et non résolu |
| 18 | Statut | `OK` ou texte d’erreur | Vocabulaire contrôlé, avec détail dans la feuille de traçabilité |

Source : fonctions `parse_record`, `get_subfields`, `get_titre_200` et constante `COLUMNS` dans le [script étudié](https://github.com/azur-scd/sudoc_export/blob/0355ce118f31e47c16bc1adfe21bede6525fb7a7/sudoc_export.py).

### 7.1. Règles bibliographiques à fixer

**Travaux universitaires.** Le script distingue `m` et `7` à la position 4 de `105$a`. La documentation Sudoc associe ces codes aux thèses et mémoires originels, et explique la correspondance entre sous-zones de saisie et positions du format d’export. La recette doit couvrir cette conversion et éviter de tester uniquement des exemples au format de saisie. Référence : [ABES — zone 105](https://documentation.abes.fr/sudoc/formats/unmb/zones/105.htm).

**Adresse bibliographique.** Les indicateurs de 214 distinguent notamment publication, production, diffusion, fabrication et copyright. Le regroupement actuel de tous les `$c` sous « Éditeur » est donc insuffisant pour garantir la signification de la colonne. Référence : [ABES — zone 214](https://documentation.abes.fr/sudoc/formats/unmb/zones/214.htm).

**Règle cible proposée pour 214/210 :** sélectionner d’abord les occurrences de publication (`ind2=0`) ; en leur absence, permettre les occurrences de production (`ind2=1`) en signalant la nature de la mention. Extraire séparément lieu, nom et date pour que l’absence du nom n’efface pas les deux autres informations. Un repli vers 210 doit être documenté et tracé ; les mentions de diffusion, fabrication et copyright ne doivent pas être présentées silencieusement comme des mentions de publication. Les cas complexes ou mixtes restent à valider sur corpus avant implémentation.

**Ponctuation.** Le nettoyage actuel retire systématiquement certains signes finaux dans `get_subfields`. La cible doit préserver les ponctuations porteuses de sens, notamment les points d’abréviation. Les règles d’affichage ne doivent pas altérer silencieusement les données bibliographiques.

**Valeurs multiples.** Les occurrences sont conservées. Le séparateur ` ; ` entre occurrences est maintenu par défaut ; la ponctuation interne dépend du champ. Cette représentation ne permet pas de reconstruire sans ambiguïté la structure UNIMARC complète.

### 7.2. Traçabilité proposée sans changer les 18 colonnes

La feuille « Traitement » comprend un résumé de session et un tableau à une ligne par entrée : numéro d’ordre, valeur saisie, PPN normalisé, PPN retourné, statut technique, message, avertissements et nombre de tentatives. Le résumé mentionne la version de l’application, les dates de début et de fin, le service utilisé, les réglages et les compteurs.

La colonne principale « PPN » conserve le `001` sur les notices récupérées. En cas d’erreur, elle reprend l’identifiant demandé. La feuille « Traitement » rend explicite cette différence et permet les rapprochements avec la liste d’origine.

## 8. États et gestion des erreurs

| Situation | Statut proposé | Comportement attendu |
| --- | --- | --- |
| Notice exploitable et identifiant cohérent | OK | Export des données |
| Notice lue, mais anomalie ou donnée essentielle manquante | À vérifier | Export des données disponibles et motif précis |
| PPN retourné différent du PPN demandé | À vérifier | Conserver les deux identifiants ; ne pas affirmer automatiquement une fusion |
| Entrée incorrecte | PPN invalide | Aucun appel réseau ; valeur d’origine conservée |
| Réponse 404 | Non récupéré — HTTP 404 | Ne pas interpréter cela comme une preuve de suppression de la notice |
| Timeout après tentatives | Délai dépassé | Conserver l’entrée et permettre une relance |
| Erreur de connexion | Erreur réseau | Détail exploitable dans le journal |
| HTTP 429 ou 5xx persistant | Service indisponible | Tentatives bornées et respect des indications du serveur |
| XML illisible ou structure non reconnue | Erreur XML | Aucune ligne vide marquée `OK` |
| Arrêt demandé avant traitement de l’entrée | Non traité | Présence dans l’export du lot avec motif d’arrêt |
| Export impossible | Échec d’enregistrement | Préserver les données et proposer un autre chemin |

Le bilan doit distinguer `OK`, `À vérifier`, échecs et non traités. Leur somme doit correspondre au nombre d’entrées non vides du lot, doublons compris selon la politique retenue. Un champ secondaire absent, comme l’ISBN, n’est pas à lui seul un échec.

## 9. Exigences non fonctionnelles

### Fiabilité et intégrité

- Aucun résultat perdu silencieusement lors d’un effacement, d’une fermeture ou d’un nouveau traitement.
- Écriture des sauvegardes dans un fichier temporaire du même dossier, puis remplacement contrôlé de la sauvegarde précédente.
- Extraction indépendante de la présentation graphique et de l’écriture Excel.
- Erreurs par notice isolées ; incidents globaux présentés clairement.
- Aucune interprétation des métadonnées comme formule Excel.

### Performance

Le traitement reste séquentiel. Pour `N` requêtes, une estimation simplifiée est : somme des durées des appels + `(N − 1) × délai` + temps d’export et de sauvegarde. À 500 ms, 500 requêtes représentent environ 250 secondes d’attente interrequêtes, auxquelles s’ajoute le réseau. Ce calcul n’est pas une mesure du temps réel du service.

La recette doit mesurer séparément réseau, extraction et export sur des lots de 100, 500 et 1 000 entrées. Un essai plus volumineux, par exemple 5 000 notices simulées, permettra de vérifier mémoire et réactivité sans solliciter inutilement le service public. Les volumes officiellement pris en charge seront annoncés après ces mesures.

### Compatibilité et diffusion

Windows 11 constitue la plateforme cible proposée. Le README mentionne Python 3.9+ pour construire l’exécutable ; cette indication n’est pas une matrice de compatibilité vérifiée. La version consolidée devra déclarer une version Python maintenue et testée, ses dépendances et la procédure de création de l’exécutable. Les autres systèmes restent hors garantie initiale.

Le test du programme source ne suffit pas à valider le `.exe` : la recette doit inclure un poste sans Python, les chemins contenant des espaces ou des accents et les restrictions usuelles du poste de travail.

### Accessibilité et confidentialité

Les opérations principales doivent être réalisables au clavier, avec un focus visible et des messages textuels compréhensibles. Le redimensionnement et l’affichage Windows à 125 % et 150 % sont à vérifier.

Le produit fonctionne localement et envoie les PPN au service interrogé. Aucun ajout de télémétrie, de compte utilisateur ou de stockage distant n’est prévu. Les listes et fichiers restent sous le contrôle de l’utilisateur.

## 10. Défauts et risques prioritaires identifiés

| ID | Priorité | Constat et preuve | Conséquence | Réponse attendue |
| --- | --- | --- | --- | --- |
| DEF-01 | P0 | Les sélecteurs de champs sont sans espace de noms ; un XML qualifié donne PPN et titre vides avec statut `OK` dans le test ciblé | Faux succès bibliographique | REC-02 et REC-04 |
| DEF-02 | P0 | `parse_record` fixe `OK` sans vérifier `001`, le titre ou la cohérence de l’identifiant | Absence de données non détectée | Statuts et contrôles explicites |
| DEF-03 | P0 | « Effacer » réinitialise `_results` pendant que le worker peut encore y ajouter des résultats ; les widgets sont manipulés dans ce worker | Résultats incomplets et état d’interface incohérent possibles | IHM-02 et IHM-03 |
| DEF-04 | P0 | Sauvegarde uniquement aux multiples du seuil ; pas de sauvegarde de fin | Pour 150 entrées et un seuil de 100, la sauvegarde périodique ne contient que les 100 premières | EXP-04 |
| DEF-05 | P0 | Une valeur de titre `=1+1` est enregistrée avec le type cellule `f` dans le test | Métadonnée interprétée comme formule | EXP-02 |
| DEF-06 | P0 | Présence de `214$c` utilisée comme condition commune pour éditeur, lieu et date | Perte possible de lieu ou de date présents en 214 | Règles d’extraction indépendantes |
| DEF-07 | P1 | Indicateurs 214 ignorés et ponctuation finale supprimée génériquement | Interprétation trompeuse ou altération des valeurs | Validation métier des règles |
| DEF-08 | P1 | Encodages, tabulations et doublons traités de manière incomplète ou hétérogène | Échecs évitables et appels répétés | ENT-01 à ENT-08 |
| DEF-09 | P1 | Pas de relance réseau, de reprise de session ou d’arrêt contrôlé | Coût des incidents sur les lots longs | Consolidation progressive |
| DEF-10 | P1 | Documentation en décalage ; aucun fichier de dépendances ou test suivi dans l’arbre étudié | Diffusion et maintenance fragiles | Documentation et recette versionnées |

DEF-01 et DEF-05 sont reproduits dans des tests ciblés. Les autres constats résultent de la lecture du code et doivent être couverts par la recette. Le comportement du service en production et les incidents d’interface ne sont pas présentés comme des observations terrain.

## 11. Architecture et livrables techniques proposés

Le code actuel tient dans un fichier : constantes, helpers MARCXML, requête réseau, export Excel et classe Tkinter. Une refonte lourde n’est pas nécessaire pour répondre au besoin.

La cible doit rendre séparément testables cinq responsabilités : préparation des entrées, client Sudoc, extraction bibliographique, écriture Excel, interface et orchestration. Le traitement réseau peut rester dans un worker ; une file d’événements consommée par le thread principal permettrait de mettre à jour l’interface.

Livrables attendus pour la consolidation :

- code corrigé et versionné ;
- dépendances déclarées et procédure d’installation reproductible ;
- README aligné sur les options et colonnes réellement livrées ;
- corpus XML de recette avec résultats attendus et provenance ;
- tests ciblant extraction, erreurs et intégrité des exports ;
- procédure de construction et de recette du `.exe` ;
- notes de version et limites connues.

La licence de diffusion devra être explicitée par le responsable du dépôt : aucun fichier de licence n’est présent dans l’arbre étudié. Le choix d’une licence n’est pas présumé par ce PRD.

## 12. Plan de recette

Les scénarios ci-dessous définissent des tests à réaliser sur la version corrigée. Ils ne sont pas tous exécutés dans le cadre de ce document.

| Test | Situation | Résultat attendu |
| --- | --- | --- |
| T01 | XML équivalents sans namespace, avec namespace par défaut et avec préfixe | Données bibliographiques identiques |
| T02 | XML malformé, HTML d’erreur ou document sans notice | Erreur explicite, jamais `OK` |
| T03 | `001` absent ou différent du PPN demandé | Absence signalée ou statut « À vérifier » ; identifiants traçables |
| T04 | PPN avec zéro initial et clé `X` | Valeur préservée de l’entrée à l’Excel |
| T05 | Séparateurs mixtes et tabulations | Même découpage en collage et en import |
| T06 | Fichiers UTF-8, UTF-8 BOM et UTF-16 BOM | Aucune pollution du premier PPN ; erreurs d’encodage compréhensibles |
| T07 | Liste avec doublons et entrées invalides | Ordre et occurrences conservés ; aucun appel pour un invalide |
| T08 | HTTP 404, 429, 500, timeout et erreur de connexion simulés | Classification correcte et tentatives conformes |
| T09 | 214 contenant une date ou un lieu mais pas de `$c` | Valeurs présentes conservées |
| T10 | Plusieurs 214 de natures différentes et coexistence avec 210 | Sélection conforme à la règle métier, sans mélange silencieux |
| T11 | Titres à plusieurs `$a`, compléments, parties et écritures | Lisibilité, ordre et caractères préservés |
| T12 | Thèse, mémoire, périodique, ressource avec leader incomplet | Type conforme au corpus ; repli explicite |
| T13 | Titre `=1+1`, identifiants longs, caractères non latins | Cellules textuelles et classeur réouvrable |
| T14 | Lots de 99, 100, 101 et 150 entrées avec seuil 100 | Sauvegarde finale contenant tous les résultats |
| T15 | Effacer, importer, relancer ou fermer pendant le travail | Lot actif et résultats protégés |
| T16 | Arrêt volontaire au milieu d’un lot | Données acquises conservées ; restantes marquées non traitées |
| T17 | Fichier ouvert dans Excel ou dossier non accessible | Message clair et possibilité de choisir un autre fichier |
| T18 | Exécutable sur Windows sans Python, clavier et zoom système | Parcours principal utilisable et export réussi |

### Vérifications effectivement réalisées pour ce PRD

| Vérification | Résultat observé |
| --- | --- |
| XML synthétique sans namespace | PPN `123456789`, titre « Titre de test », type « Monographie », statut `OK` |
| Même XML avec namespace MARCXML | PPN vide, titre vide, type « Inconnu », statut `OK` : défaut confirmé |
| Export d’un PPN commençant par zéro | Valeur `012345678` conservée comme chaîne |
| Export du titre synthétique `=1+1` | Type cellule formule : défaut confirmé |
| Structure du classeur temporaire | 18 colonnes, filtre `A1:R2`, première ligne figée avec `A2` |

Ces contrôles utilisent un substitut de l’appel HTTP ; ils n’établissent ni la disponibilité actuelle de l’endpoint ni la forme exacte de toutes ses réponses réelles. La validation finale doit intégrer un petit corpus récupéré effectivement auprès du Sudoc, puis conservé pour les tests reproductibles.

## 13. Feuille de route proposée

| Étape | Contenu | Condition de sortie |
| --- | --- | --- |
| 1 — Fiabiliser l’existant | Namespaces, faux `OK`, PPN demandé/retourné, protection des résultats, UI, formules Excel et sauvegarde finale | Tous les P0 validés, aucun faux succès sur le corpus |
| 2 — Stabiliser les règles métier | Titres, 214/210, types, ponctuation, valeurs multiples, alignement du README | Corpus validé par un référent bibliographique |
| 3 — Consolider l’usage | Encodages, doublons, arrêt, relances réseau, traçabilité, diffusion Windows | Recette fonctionnelle et Windows réussie |
| 4 — Évolutions facultatives | Reprise de session, import CSV/XLSX, champs configurables, export XML, enrichissement des responsabilités | Besoin confirmé et périmètre distinct accepté |

Les étapes se terminent par des résultats vérifiables. Aucun calendrier de développement n’est fixé avant examen des corrections et disponibilité du référent métier.

## 14. Arbitrages à valider

| Question | Proposition retenue pour ce PRD | Effet |
| --- | --- | --- |
| Quel périmètre produit ? | Application de bureau PPN → Excel | Préserve un outil simple à diffuser |
| Faut-il supprimer les doublons ? | Conserver les lignes, mutualiser les requêtes dans le lot | Facilite le rapprochement avec la source |
| Faut-il modifier les colonnes ? | Garder les 18 colonnes, ajouter une feuille de traçabilité | Limite les ruptures pour les usages existants |
| Que faire d’une notice incomplète ? | Exporter les données disponibles avec avertissement si nécessaire | Rend le résultat exploitable sans masquer l’anomalie |
| Quelle adresse bibliographique choisir ? | Publication prioritaire ; autres natures et replis explicités | Évite de confondre éditeur, producteur et imprimeur |
| Faut-il un bouton d’arrêt ? | Oui en P1, avec conservation du lot partiel | Permet les traitements longs |
| Faut-il reprendre après fermeture ? | P2, au-delà de la sauvegarde Excel | Évite de confondre sauvegarde et reprise |
| Quel volume garantir ? | Décision après mesure ; recette à 100, 500 et 1 000 entrées | Engagement fondé sur des essais |
| Quels systèmes garantir ? | Windows 11 en premier | Recette de diffusion maîtrisable |

## 15. Conditions de validation de la version consolidée

La version peut être considérée comme prête à diffuser lorsque :

1. chaque exigence P0 dispose d’une preuve de recette satisfaisante ;
2. le corpus métier couvre les types de notices et variantes XML utilisés ;
3. les règles d’extraction et leurs limites sont relues par un référent bibliographique ;
4. les lots comportant des erreurs restent exportables et traçables ;
5. l’enregistrement final conserve tous les résultats et les cellules restent textuelles ;
6. le parcours complet est validé sur Windows avec le mode de distribution retenu ;
7. la documentation correspond exactement à la version distribuée.

## 16. Sources et portée de l’analyse

- [Dépôt SudocExport](https://github.com/azur-scd/sudoc_export) — point d’entrée du projet.
- [Code au commit étudié](https://github.com/azur-scd/sudoc_export/blob/0355ce118f31e47c16bc1adfe21bede6525fb7a7/sudoc_export.py) — source principale de la description de l’existant et des constats de code.
- [README au commit étudié](https://github.com/azur-scd/sudoc_export/blob/0355ce118f31e47c16bc1adfe21bede6525fb7a7/README.md) — installation, usage annoncé et diffusion envisagée.
- [ABES — zone 200](https://documentation.abes.fr/sudoc/formats/unmb/zones/200.htm) — distinction entre titre parallèle et complément du titre.
- [ABES — zone 105](https://documentation.abes.fr/sudoc/formats/unmb/zones/105.htm) — nature du contenu et correspondance vers le format d’export.
- [ABES — zone 214](https://documentation.abes.fr/sudoc/formats/unmb/zones/214.htm) — sens des indicateurs et nature des mentions bibliographiques.

Sources consultées le 15 septembre 2026. Les règles métier nouvelles, priorités, indicateurs et choix d’architecture sont des propositions de ce PRD. La description du code est une analyse statique complétée par les vérifications ciblées indiquées ; elle ne constitue pas une certification de fonctionnement en production.
