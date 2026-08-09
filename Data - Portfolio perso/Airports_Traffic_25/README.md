# ✈️ [Analyse du trafic aérien et de la charge piste en Europe]

## 🎯 Objectif du projet

Analyse du trafic aérien sur l'année 2025 à l'échelle Européenne+Caucase afin d'identifier tendances, répartition, charges et saisonnalité géograhique notamment.

## 📊 Source des données

- **Origine :** Google Sheets (.csv importé depuis Kaggle - dataset by Nilesh Kadam)
- **Contenu supplémentaire :** Données sur les pistes d'aéroports (surface, dimensions, localisation) - dataset by OurAirports
- **Volume :** [116020 lignes, 34128 colonnes après nettoyage]

## 🛠️ Outils utilisés

- **Google Colab** — environnement d'exploration et de nettoyage
- **pandas / numpy** — manipulation des données
- **Google Sheets** (export CSV via URL) — source des données

## 🔍 Démarche

### 1. Exploration (EDA)
- Inspection structurelle (types, dimensions, aperçu des données)
- Analyse des valeurs manquantes et doublons
- Étude des variables catégorielles (ex: types de surface de piste) via `value_counts()`
- Statistiques descriptives sur les variables numériques

### 2. Nettoyage
- Conversion des types (dates en `datetime`, colonnes numériques en `int` après traitement des valeurs manquantes)
- Traitement des valeurs manquantes (imputation à 0 ou suppression selon la signification métier de la donnée)
- Suppression des lignes non pertinentes (ex: lignes où deux indicateurs clés étaient simultanément à 0)
- Normalisation des catégories textuelles (harmonisation de la casse, regroupement des variantes de "surface" en catégories standardisées : ASP, CON, OTH...)

### 3. Transformation

en cours

### 4. Analyse

en cours

## 📈 Résultats clés

en cours

## 📂 Structure du dossier

```
projet-trafic-aerien/
├── README.md
├── notebook_exploration_nettoyage.ipynb
└── data/
    └── 
```

## 🚀 Pour reproduire l'analyse

1. Ouvrir le notebook dans Google Colab
2. Le dataset est chargé directement depuis Google Sheets via une URL d'export CSV — aucune installation locale nécessaire
3. Exécuter les cellules dans l'ordre (Exploration → Nettoyage → Transformation → Analyse)

## 🔮 Pistes d'amélioration

...
