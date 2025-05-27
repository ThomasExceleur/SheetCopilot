# Analyse du Projet SheetCopilot

Ce document détaille la structure, le fonctionnement et l'exécution du projet `SheetCopilot`. L'objectif de `SheetCopilot` est d'assister les utilisateurs dans la manipulation de tableurs (Excel, Google Sheets) via des commandes en langage naturel, en s'appuyant sur de Grands Modèles de Langage (LLM).

## 1. Objectif et Fonctionnalités Clés

*   **Assistant pour Tableurs** : `SheetCopilot` agit comme un agent intelligent capable de comprendre des instructions en langage naturel (par exemple, "Calcule la somme de la colonne B et mets le résultat en C1") et de les exécuter sur un tableur.
*   **Support pour Excel et Google Sheets** : Il est conçu pour fonctionner avec Microsoft Excel (principalement sous Windows) et Google Sheets (via un plugin).
*   **Utilisation de LLM** : Le cœur de l'agent repose sur des Grands Modèles de Langage (comme ceux d'OpenAI) pour interpréter les commandes utilisateur et générer des plans d'action.
*   **Chain-of-Thoughts (CoT)** : L'agent utilise une approche de "chaîne de pensées", où le LLM décompose le raisonnement en étapes intermédiaires avant de produire une solution. Cela améliore la transparence et la robustesse.
*   **Actions Atomiques** : Les tâches complexes sont décomposées en une série d'actions élémentaires (ex: écrire dans une cellule, appliquer un format, créer un graphique) que l'agent peut exécuter. Pour Excel, ces actions sont souvent implémentées via la bibliothèque `xlwings` (qui utilise `pywin32`).
*   **Mécanisme de Feedback** : L'agent peut observer l'état du tableur et ajuster ses actions en fonction des retours d'erreur ou de documents externes, créant une boucle de contrôle.
*   **Machine à États** : La logique de planification de l'agent est implémentée comme une machine à états, guidant le processus de traitement d'une instruction.

## 2. Environnement d'Exécution et Installation

### Pour la version Excel :

*   **Système d'exploitation** : Windows uniquement.
*   **Python** : Version 3.10 requise.
*   **Gestionnaire d'environnement** : L'utilisation de Conda est recommandée.
    ```bash
    conda create -n sheetcopilot python=3.10
    conda activate sheetcopilot
    ```
*   **Dépendances Python** : Installées via `pip` à partir du fichier `requirements.txt` (et `agent/requirement.txt`).
    ```bash
    pip install -r requirements.txt
    ```
    Les dépendances clés incluent :
    *   `openai`: Pour interagir avec les API d'OpenAI.
    *   `pandas`: Pour la manipulation de données.
    *   `openpyxl`, `xlwings`, `pywin32`: Pour lire, écrire et contrôler les fichiers Excel.
    *   `PyYAML`: Pour la gestion des fichiers de configuration.
    *   `tiktoken`: Pour compter les tokens pour les modèles OpenAI.
*   **Microsoft Excel** : Doit être installé et un classeur doit être ouvert pour le mode interactif.
*   **Configuration** : Nécessite de configurer les clés API d'OpenAI et d'autres paramètres dans `agent/config/config.yaml`.

### Pour la version Google Sheets :

*   Disponible sous forme de **plugin** sur le Google Workspace Marketplace.
*   Permet l'annulation des opérations (contrairement à la version Excel).

## 3. Structure du Projet

Le projet est organisé comme suit :

*   **`.git/`**: Dossier pour le contrôle de version Git.
*   **`agent/`**: Contient le code source principal de l'agent.
    *   **`Agent/`**: Logique centrale de l'agent.
        *   `agent.py`: Probablement la classe principale de l'agent.
        *   `xwAPI.py`: Implémentation des actions atomiques pour Excel en utilisant `xlwings`.
    *   **`config/`**:
        *   `config.yaml`: Fichier de configuration principal (clés API, modèles LLM, chemins, etc.).
    *   **`utils/`**: Fonctions utilitaires.
    *   **`SheetCopilot_example_logs/`**: Exemples de fichiers journaux (logs).
    *   **`main.py`**: Point d'entrée pour l'exécution automatisée de tâches sur un jeu de données (batch processing). Utilisé pour l'évaluation.
    *   **`interaction.py`**: Point d'entrée pour le mode interactif avec Excel. L'utilisateur tape des commandes.
    *   **`evaluation.py`**: Script pour évaluer les performances de l'agent sur les tâches du dataset.
    *   `requirement.txt`: Dépendances spécifiques au module `agent`.
*   **`assets/`**: Images, icônes et autres ressources utilisées (par exemple, dans le `README.md`).
*   **`dataset/`**: Jeu de données pour l'entraînement et l'évaluation.
    *   `dataset.xlsx`: Liste des tâches, instructions, contextes.
    *   `task_sheets/`: Classeurs Excel initiaux pour les tâches.
    *   `task_sheet_answers/` (et `_v2`): Classeurs Excel de référence (solutions attendues) et fichiers YAML de vérification.
*   **`.gitignore`**: Fichiers et dossiers à ignorer par Git.
*   **`LICENSE`**: Licence du projet.
*   **`README.md`**: Documentation complète du projet, incluant l'installation, l'utilisation, et les détails de l'évaluation.
*   **`requirements.txt`**: Liste des dépendances Python globales du projet.

## 4. Fonctionnement Détaillé

### A. Flux de travail général (Mode Interactif avec `interaction.py`)

1.  **Initialisation** :
    *   L'utilisateur lance `python agent/interaction.py -c agent/config/config.yaml` après avoir ouvert un fichier Excel.
    *   L'agent se connecte à l'instance Excel active.
2.  **Entrée Utilisateur** :
    *   L'utilisateur saisit une instruction en langage naturel (ex: "Mets en gras les cellules de A1 à A5").
3.  **Interprétation par le LLM** :
    *   L'instruction est envoyée à un Grand Modèle de Langage (configuré dans `config.yaml`, par exemple, GPT-3.5-turbo ou GPT-4).
    *   Le LLM, potentiellement en utilisant une approche "Chain-of-Thoughts", génère un plan d'action. Ce plan peut être une séquence d'actions atomiques.
4.  **Exécution des Actions** :
    *   L'agent traduit le plan du LLM en appels aux fonctions définies dans `Agent/xwAPI.py`.
    *   Ces fonctions utilisent `xlwings` pour manipuler le fichier Excel ouvert.
5.  **Feedback (Optionnel)** :
    *   L'agent peut lire l'état du tableur pour vérifier le résultat de l'action ou pour informer les étapes suivantes.
    *   En cas d'erreur, cette information peut être utilisée pour tenter de corriger l'action.
6.  **Itération** : L'utilisateur peut donner de nouvelles instructions.

### B. Exécution Automatisée (`agent/main.py`)

Ce script est utilisé pour évaluer `SheetCopilot` sur un ensemble de tâches prédéfinies (issues de `dataset/dataset.xlsx`).

1.  **Chargement des Tâches** : Le script lit les tâches (instruction, contexte, fichier source) depuis le fichier `dataset.xlsx`.
2.  **Traitement Asynchrone** :
    *   Il utilise `asyncio` pour gérer un "producteur" de tâches et un (ou plusieurs, bien que `worker=1` soit recommandé pour Excel) "travailleur".
    *   Le producteur prépare l'environnement pour chaque tâche (copie du fichier source) et l'ajoute à une file d'attente.
    *   Le travailleur prend une tâche, instancie l'agent, et appelle `agent.Instruction()` pour la traiter.
3.  **Exécution et Journalisation** :
    *   L'agent exécute la tâche, en interagissant avec le fichier Excel.
    *   Les résultats (fichier Excel modifié) et des logs détaillés (y compris les étapes de la CoT et les communications avec le LLM) sont sauvegardés dans un dossier spécifique à la tâche et à l'itération (car chaque tâche peut être répétée plusieurs fois).
4.  **Suivi et Reprise** : Un fichier `index.txt` garde la trace des tâches complétées, permettant de reprendre une exécution interrompue.
5.  **Évaluation** : Les résultats et logs générés par `main.py` sont structurés pour être utilisés par `evaluation.py` afin de calculer des métriques de performance (par exemple, taux de succès).

### C. Configuration (`agent/config/config.yaml`)

Ce fichier est central et contrôle de nombreux aspects :

*   **Clés API OpenAI** (`OPENAI_API_KEY`).
*   **Modèles LLM à utiliser** (pour la planification et la révision, `model_name` sous `ChatGPT_1` et `ChatGPT_2`).
*   **Paramètres des LLM** (température, `max_tokens`).
*   **Chemins** (`task_path`, `source_path`, `save_path`).
*   **Nombre de répétitions** pour chaque tâche (`repeat`) dans `main.py`.
*   **Nombre de travailleurs** (`worker`) dans `main.py`.
*   Activation de la documentation "oracle" des API (`use_oracle_API_doc`).

## 5. Comment Exécuter

### Mode Interactif (pour tester sur un fichier Excel)

1.  S'assurer que l'environnement est configuré (Python 3.10, Conda, dépendances installées).
2.  Renseigner `agent/config/config.yaml` avec votre clé API OpenAI.
3.  Ouvrir le fichier Excel que vous souhaitez manipuler.
4.  Depuis la racine du projet, exécuter :
    ```bash
    python agent/interaction.py -c agent/config/config.yaml
    ```
5.  Suivre les invites pour entrer vos commandes en langage naturel.

**Attention** : Les opérations effectuées en mode interactif sur Excel **ne sont pas annulables** via `Ctrl+Z` dans Excel. Travaillez sur des copies de vos fichiers importants.

### Exécution de Tâches en Batch (pour évaluation)

1.  S'assurer de la configuration de l'environnement et de `agent/config/config.yaml`.
2.  Vérifier les chemins dans `config.yaml` (notamment `task_path` pointant vers `dataset/dataset.xlsx` ou un sous-ensemble, et `source_path` vers `dataset/task_sheets/`).
3.  Depuis la racine du projet (ou le dossier `agent/`), exécuter :
    ```bash
    python agent/main.py -c agent/config/config.yaml
    ```
4.  Les résultats seront sauvegardés dans le dossier spécifié par `save_path` dans la configuration.
5.  Ensuite, pour évaluer les résultats, depuis le dossier `agent/`, exécuter :
    ```bash
    python evaluation.py -c agent/config/config.yaml 
    ```
    (Assurez-vous que `config.yaml` pointe vers le bon `save_path` contenant les résultats).

## 6. Points Notables

*   **Focus sur Windows pour Excel** : L'implémentation actuelle pour Excel est fortement liée à Windows à cause de `pywin32` et `xlwings`.
*   **Dataset et Évaluation Robustes** : Le projet fournit un dataset conséquent et un cadre d'évaluation pour mesurer les performances des agents manipulateurs de tableurs.
*   **Recherche Active** : Le projet est lié à une publication de recherche (NeurIPS 2023) et continue d'évoluer.

Ce document devrait vous donner une bonne vue d'ensemble pour commencer à explorer et à utiliser `SheetCopilot`. N'hésitez pas à consulter le `README.md` original pour plus de détails et les dernières mises à jour. 