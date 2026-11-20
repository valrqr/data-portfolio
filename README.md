# Valentin Rouquier — Data Analyst Portfolio

## 👋 À propos

FR: Disposant d’une solide expérience de plus de 12 ans dans la coordination et la gestion des opérations au sein de secteurs exigeants (aéroportuaire et portuaire), j'ai mené de nombreux projets d'améliorations et d'optimisation des process internes dans les services où j'ai évolué grâce à mes compétences en Excel & VBA. Devant ce constat, j'ai récemment décidé de me former aux compétences techniques de Data Analyst. 
J'espère ainsi grâce à mon excellente vision terrain et à une forte capacité à automatiser et analyser les données pour l'aide à la décision pouvoir évoluer dans ce nouveau domaine avec succès.

EN: With over 12 years of solid experience in coordinating and monitoring operations within highly demanding environnements, I have successfully led numerous internal process optimization and improvement projects throughout my career thanks to my Excel & VBA skills.
Driven by this experience, I recently decided to follow a challenging training in Data Analytics. Combining strong on-the-ground operational insight with a proven ability to automate and analyze data for strategic decision-making, I am now eager to successfully transition into this new field.

**Compétences clés :** Python (pandas, numpy, matplotlib/seaborn) · SQL · Google Sheets · Jupyter/Colab · Git/GitHub · PowerBI · BigQuery · dbt · Machine Learning · n8n · FastAPI · Docker · Microsoft Excel · Visual Basic

📫 [LinkedIn](https://www.linkedin.com/in/valentin-rouquier-286833347) · [Email](mailto:val.rouquier@gmail.com) · [CV](https://drive.google.com/file/d/1-GrR1cVQz7kPAXfyNl9JleDdZTI1zn7y/view?usp=sharing)

---

## 📁 Projets

Chaque projet présenté ici vise à mettre en avant mes compétences acquises en data suite à ma formation auprès de La Capsule mais également celles sur Excel à travers la présentation de certains projets menés dans le cadre de mes précédentes fonctions chez Intercruises et Aéroports de la Côte d'Azur.

### 🌊 [SILLAGE – La flotte de porte-conteneurs CMA CGM vue par les données publiques européennes] - Projet Perso

Que disent les données publiques sur la performance énergétique d'une flotte, navire par navire ? À partir du registre européen EU MRV (déclarations 2024 et 2025), le projet analyse les 197 porte-conteneurs gérés par le groupe CMA CGM et les compare à plus de 2 100 autres. Principaux constats : la flotte a ralenti de 1,2 nœud en un an (-2,6 % de carburant par mille) ; en comparant chaque navire à ses sister-ships, 4 navires consomment au moins 15 % de plus que leur série ; avec l'arrivée des navires au GNL, le méthane pèse désormais 5 % des émissions en équivalent CO₂ de la flotte.

**Techniques utilisées :** nettoyage de fichiers réglementaires Excel, grandeurs dérivées (distance, vitesse moyenne), comparaison à des groupes de pairs (indice vs série), modélisation dbt avec tests, tableau de bord Power BI avec mesures DAX sur Microsoft Fabric (Lakehouse, Direct Lake)
**Outils :** Python (pandas, pytest), Google BigQuery, dbt, Power BI / Microsoft Fabric
🔗 [Voir le projet](https://github.com/valrqr/sillage)

### ✈️ [MAINTAIN – Les paliers en descente vers Nice Côte d'Azur] - Projet Perso

Combien de carburant coûtent les paliers imposés pendant la descente vers Nice ? À partir de données ADS-B publiques collectées pendant 3 jours (OpenSky Network), le projet reconstruit 385 arrivées, détecte 427 paliers et estime le carburant et le CO₂ en jeu. Principaux constats : 63 % des paliers se font entre 3 000 et 4 500 ft, juste avant l'axe final ; ils concernent 1 arrivée sur 3 aux heures calmes, mais 3 sur 4 dès 4 arrivées par heure ; et 35 % des paliers concentrent 80 % du carburant en jeu.

**Techniques utilisées :** collecte d'API avec authentification OAuth2 et gestion des quotas, découpage de trajectoires en vols, détection d'événements sur séries temporelles, modèle de consommation de carburant, statistiques circulaires (dispersion des trajectoires), tests unitaires sur trajectoires synthétiques, modélisation dbt avec tests, cartographie
**Outils :** Python (pandas, matplotlib, OpenAP, pytest), API OpenSky, Google BigQuery, dbt
🔗 [Voir le projet](https://github.com/valrqr/maintain-4000)

### 🏢 [Mitoyen – Assistant open source pour syndics bénévoles de copropriété] - Projet Perso

Outil qui permet à un copropriétaire bénévole de gérer une copropriété de moins de 50 lots, sur un principe simple : l'agent prépare, le syndic valide et signe. Première brique livrée : calcul des appels de fonds au centime près selon les tantièmes de chaque lot, génération des avis en PDF, puis envoi automatisé après validation du syndic par mail.

**Techniques utilisées :** modélisation de données (lots, tantièmes, clés de répartition), contrôles de cohérence, répartition au centime par la méthode du plus fort reste, API REST sécurisée, conteneurisation, workflow avec validation humaine
**Outils :** Python, FastAPI, ReportLab, n8n, Docker Compose, YAML
🔗 [Voir le projet](https://github.com/valrqr/mitoyen)

NB: en cours - échéancier automatique et relances des impayés à suivre

### ✈️ [Analyse du trafic aérien et de la charge piste en Europe] - Projet Perso

Exploration et nettoyage d'un dataset des données trafic 2025 d'aéroports européens, avec construction de tables de dimension et gestion de la qualité des données (valeurs manquantes, doublons, cohérence des catégories).

**Techniques utilisées :** pandas, jointures, feature engineering, nettoyage de données catégorielles
**Outils :** Google Colab, Google Sheets
🔗 [Voir le projet](https://github.com/valrqr/data-portfolio/tree/Projet_solo_EU_Airports_Traffic25)

NB: en cours - transformation des données à suivre

### [Etude des facteurs de réadmission chez la patientèle diabétique aux US 1999-2008] - Projet de fin de formation

Projet réalisé en binôme à partir des données Health Facts (130 hôpitaux américains, 1999-2008) : prédire la réadmission à moins de 30 jours des patients diabétiques et identifier des leviers d'action pour l'éviter.

- **Pipeline de données complet** : ingestion avec Airbyte, modélisation dbt sur BigQuery (staging → intermediate → marts en schéma en étoile, tests de qualité), orchestration Airflow sous Docker
- **Machine Learning** : segmentation KMeans, régression logistique et XGBoost (AUC ≈ 0,66). Les trois approches convergent : l'historique d'hospitalisations antérieures est le premier facteur de risque. Un groupe à haut risque de 6 % des patients se réadmet à 25 %, contre 11 % en moyenne
- **Restitution pour un public non technique** : tableau de bord Power BI (matrice de confusion en langage métier, scores du modèle, calculateur de groupe à risque avec estimation des hospitalisations évitées)
- **Recommandations** : suivi renforcé à deux niveaux selon le risque, éviter les changements de traitement inutiles et renforcer l'éducation thérapeutique des patients

**Techniques utilisées :** modélisation en étoile, tests dbt, prévention des fuites de données au niveau patient (GroupShuffleSplit), gestion du déséquilibre de classes, clustering, classification supervisée, sélection de variables (9 variables sur 115 pour une performance comparable), mesures DAX
**Outils :** Google Colab, Google Sheets, Google BigQuery, Airbyte, Airflow, dbt, Docker, Python (pandas, scikit-learn, XGBoost), Power BI
🔗 [Voir le projet]()

### [Outil d'aide à la négociation] - Outil d'aide à la négociation des tarifs autocaristes et gestion des stats (Intercruises)

à venir

**Techniques utilisées :** [...]
**Outils :** [...]
🔗 [Voir le projet]()

### 🛫 [Indicateurs.xlsm] - Outil de gestion de données et de reporting (ACA)

Outil métier développé pour l'équipe des opérations aéroportuaires : saisie quotidienne guidée de l'activité côté piste (inspections, péril animalier, FOD, déversements, convoyages et tractages, fret, météo, facturations), surveillance des escales, puis archivage et export automatiques des données pour les rapports PowerPoint mensuels, saisonniers (saisons IATA) et annuels. Il compte 25 formulaires et environ 10 600 lignes de VBA.

**Techniques utilisées :** contrôles de qualité des données à la saisie, Excel utilisé comme base de données (feuilles techniques masquées, plages nommées), automatisation des clôtures de période (archivage mensuel, remise à zéro), exports figés en .xlsx pour le reporting mensuel / saisonnier IATA / annuel, gestion des accès
**Outils :** Microsoft Excel, VBA (UserForms), PowerPoint, olevba pour l'extraction du code
🔗 [Voir le projet](https://github.com/valrqr/data-portfolio/tree/main/Excel_%26_VBA%20-%20Portfolio%20pro/Indicateurs)

### 🔍 [STX.xlsm] - Planning mensuel des surveillances d'escales (ACA)

Outil qui transforme l'export brut des vols du mois (Infopax) en planning de surveillances d'escales tirées au sort, pour vérifier que les assistants en escale respectent les règles d'exploitation de l'aéroport. Chaque assistant a sa semaine : l'outil tire 2 vols sur deux jours différents et prévoit 1 vol de secours, contrôle la validité du tirage, puis remplit automatiquement le modèle Word diffusé à l'équipe. La grille de contrôle utilisée sur le poste avion (une trentaine de points) est intégrée au classeur.

**Techniques utilisées :** nettoyage et filtrage d'un export brut, échantillonnage aléatoire stratifié (par assistant et par semaine) avec contrôle des règles métier, plages nommées, export Word automatisé par signets
**Outils :** Microsoft Excel, VBA (UserForm), Word, olevba pour l'extraction du code, LibreOffice pour les captures sur données fictives
🔗 [Voir le projet](https://github.com/valrqr/data-portfolio/tree/main/Excel_%26_VBA%20-%20Portfolio%20pro/STX)

### 🚗 [Permis pistes] - Examens et suivi des autorisations de conduite côté piste (ACA)

Outils conçus pour encadrer la délivrance des « permis pistes » (autorisations de conduite sur l'aire de trafic) selon le règlement (UE) 2020/2148 : un générateur de QCM aléatoires par type de permis pour l'examen théorique, et des formulaires PDF remplissables pour tracer la formation (quatre visas formateur / examinateur) et noter l'examen pratique (critère éliminatoire au moindre point non satisfaisant).

**Techniques utilisées :** traduction d'une exigence réglementaire en processus et en grilles standardisées, tirage aléatoire de questions, traçabilité pour l'audit
**Outils :** Microsoft Excel, VBA, formulaires PDF
🔗 [Voir le projet](https://github.com/valrqr/data-portfolio/tree/main/Excel_%26_VBA%20-%20Portfolio%20pro/Permis_pistes)

NB: en cours - banque de questions et générateur de QCM à ajouter

---

## 🛠️ Stack technique

| Catégorie | Outils |
|---|---|
| Langages | Python, SQL, VBA |
| Librairies | pandas, numpy, matplotlib, seaborn, OpenAP, pytest, FastAPI, ReportLab |
| Données & BI | Google BigQuery, dbt, Power BI |
| Environnements | Jupyter, Google Colab, Excel, Docker |
| Automatisation | n8n |
| Autres | Git, GitHub, Google Sheets |

---

## 📌 Structure du repo

```
├── Data - Portfolio formation
    ├── Hackaton_Final_ML
    ├── README.md
    └── ...
    ├── Projet_fin_de_formation
    ├── README.md
    └── ...
├── Data - Portfolio perso
    ├── KoPro/  → dépôt dédié : github.com/valrqr/kopro
    ├── eco-flying-nice/  → dépôt dédié : github.com/valrqr/maintain-4000
    ├── Projet Solo-1-trafic25/
        ├── README.md
        ├── notebook_exploration_et_nettoyage.ipynb
        └── data/
    ├── ...
        └── ...
├── Excel_&_VBA - Portfolio pro
  ├── Indicateurs/
    ├── README.md
    ├── src/
    │   ├── forms/
    │   ├── modules/
    │   └── _INDEX.txt
    └── docs/
        └── screenshots/
  ├── STX/
    ├── README.md
    ├── src/
    │   ├── forms/
    │   ├── modules/
    │   └── _INDEX.txt
    └── docs/
        └── screenshots/
  ├── Permis_pistes/
    ├── README.md
    ├── modeles/
    └── docs/
└── README.md
```

---

