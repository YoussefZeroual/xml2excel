# Documentation — `xml2xlsx`
## Convertisseur de corpus XML annoté vers Excel

**Public cible :** linguistes et annotateurs maîtrisant XML et les corpus écrits, sans formation Python particulière

---

## Vue d'ensemble

`xml2xlsx` transforme des fichiers XML annotés en tableaux Excel lisibles, avec mise en couleur automatique des balises. Chaque paragraphe `<p>` du corpus devient une ligne du tableau ; chaque balise imbriquée devient une colonne.

La détection et l'extraction des balises est **entièrement générique** : l'outil ne présuppose aucun schéma d'annotation particulier. Il s'adapte automatiquement à n'importe quelle hiérarchie de balises et d'attributs à partir de la DTD fournie.

---

## 1. Lancement

Depuis un terminal, dans le dossier du projet :

```
python -m xml2xlsx.gui
```

---

## 2. L'interface graphique pas à pas

### 2.1 Source d'entrée

| Option | Usage |
|--------|-------|
| **Dossier** | Traite tous les `.xml` d'un répertoire ; génère un fichier Excel par corpus + un fichier maître fusionné |
| **Fichier XML** | Traite un seul fichier `.xml` |

### 2.2 Chemin

Cliquez sur **Parcourir…** pour sélectionner votre dossier ou fichier.

### 2.3 Schéma DTD

Champ **obligatoire**. La DTD (Document Type Definition) décrit la structure de votre XML : quelles balises existent, dans quel ordre, avec quels attributs. `xml2xlsx` s'en sert pour savoir quelles colonnes créer dans Excel.

### 2.4 Options

| Option | Effet |
|--------|-------|
| **Générer un fichier XLSX individuel par fichier XML** | Crée un `.xlsx` à côté de chaque `.xml` traité (mode dossier uniquement) |
| **Générer un fichier XLSX maître** | Crée un fichier `master_output.xlsx` fusionnant tous les fichiers du dossier |
| **Calculer la colonne POSITION_INTRODD** | Ajoute une colonne indiquant si `<INTRODD>` est en position `ANTE`, `POST` ou `AUTRE` par rapport à `<PPI>` dans chaque paragraphe |

### 2.5 Lancer la conversion

Cliquez sur **▶ Lancer la conversion**. Le **Journal** (panneau noir en bas) affiche en temps réel les fichiers traités, les fichiers Excel créés, et les éventuelles erreurs.

---

## 3. Structure du fichier Excel produit

### 3.1 Colonnes générées

Le tableau suit la hiérarchie de votre XML. Pour chaque `<p>`, les colonnes sont :

| Colonne | Contenu |
|---------|---------|
| `source_file` | Nom du fichier XML d'origine |
| `p_id` | Identifiant extrait du paragraphe (première expression entre guillemets) |
| `paragraph_text` | Texte complet du paragraphe avec balises XML colorées |
| `POSITION_INTRODD` | Position d'INTRODD (si l'option est activée) |
| `NOM_BALISE_text` | Texte contenu dans `<NOM_BALISE>` |
| `NOM_BALISE_attribut` | Valeur de l'attribut `attribut` sur `<NOM_BALISE>` |

Pour les balises imbriquées, les noms de colonnes reflètent la hiérarchie par chaînage : `INTRODD_VDD_text`, `INTRODD_VDD_lemme`, etc. Quand une balise apparaît plusieurs fois dans un même paragraphe, les occurrences supplémentaires reçoivent un suffixe numérique : `EXPANSION_text`, `EXPANSION_2_text`, `EXPANSION_3_text`.

### 3.2 Mise en couleur dans la colonne `paragraph_text`

Le texte du paragraphe est restitué avec les balises XML visibles et colorées. La table ci-dessous correspond au schéma PREFAB ; avec une DTD différente, les couleurs s'appliquent aux balises du même nom si elles existent.

| Balise | Signification | Couleur |
|--------|--------------|---------|
| `PPI` | Phrase préfabriquée des interactions | Rouge foncé |
| `NONPPI` | Segment non PPI | Vert foncé |
| `DD` | Discours direct | Turquoise |
| `INTRODD` | Introducteur du discours direct | Bleu foncé |
| `VDD` | Verbe introducteur du discours direct | Violet |
| `EXPANSION` | Expansion de l'introducteur du discours direct | Orange |
| `MD` | Marqueur du discours | Bleu moyen |
| `MOD` | Modifieur | Marron |
| `APP` | Appellatif | Gris foncé |

### 3.3 Casse du texte

- En **mode fichier unique** : la casse originale est conservée.
- En **mode dossier** : les colonnes de texte sont mises en minuscules pour faciliter la comparaison et le tri, sauf `paragraph_text` et quelques colonnes contextuelles.

---

## 4. Vérification d'intégrité automatique

Après chaque conversion, l'outil compare le contenu du fichier XML source avec le tableau Excel produit, et affiche un rapport dans le Journal :

- `OK: PPI: Excel=12, xml=12` → tout correspond
- `Erreur: VDD: Excel=10, xml=11` → décalage détecté, avec indication de la position divergente

> **Note :** contrairement au moteur d'extraction — entièrement générique — la vérification d'intégrité a été développée spécifiquement pour les fichiers traités dans le cadre du projet en cours. Elle peut produire des résultats incorrects sur des corpus de structure différente.

---

## 5. Glossaire

| Terme | Signification |
|-------|--------------|
| DTD | Document Type Definition — fichier décrivant les balises autorisées et leur hiérarchie |
| Fichier maître | Fichier Excel regroupant les données de tous les XML d'un dossier dans un seul tableau |
| POSITION_INTRODD | Colonne indiquant si l'introducteur du DD précède (ANTE) ou suit (POST) la PPI dans le même paragraphe |
| PPI | Phrase préfabriquée des interactions |
| NONPPI | Segment non PPI |
| DD | Discours direct |
| INTRODD | Introducteur du discours direct |
| VDD | Verbe introducteur du discours direct |
| EXPANSION | Expansion de l'introducteur du discours direct |
| MD | Marqueur du discours |
| MOD | Modifieur |
| APP | Appellatif |
