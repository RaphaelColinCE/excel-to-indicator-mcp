# Excel-to-Indicator MCP Server

## Description

Serveur MCP (Model Context Protocol) permettant la conversion de formules Excel en formules d'indicateurs compréhensibles par une application métier.

### Contexte

Ce serveur facilite la conversion de formules Excel complexes (réparties sur plusieurs feuilles, tables, colonnes) en formules d'indicateurs basées sur des concepts, dimensions et indicateurs. Le système utilise un mapping centralisé pour traduire les références Excel vers des concepts applicatifs.

## Fonctionnalités

- ✅ Lecture de fichiers Excel
- ✅ Extraction et parsing de formules
- ✅ Traçage des dépendances (profondeur limitée)
- ✅ Résolution de références entre feuilles et tables
- ✅ Conversion via mapping Excel ↔ Concepts applicatifs
- ✅ Validation des formules générées
- ✅ Exposition via tools MCP pour LLM

## Stack technique

- **Python** >= 3.10
- **openpyxl** : Lecture des fichiers Excel
- **Pydantic** : Validation des modèles de données
- **MCP** : Serveur Model Context Protocol
- **python-dotenv** : Gestion des variables d'environnement

## Installation

### 1. Cloner le repository

```bash
git clone <repo-url>
cd excel-to-indicator-mcp
```

### 2. Créer un environnement virtuel

```bash
python -m venv venv
source venv/bin/activate  # Sur Windows: venv\Scripts\activate
```

### 3. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 4. Configurer les variables d'environnement

```bash
cp .env.example .env
# Éditer .env selon vos besoins
```

## Configuration - Où insérer vos données

### Structure des répertoires

```
data/
├── examples/                          # Fichiers d'exemple fournis
│   ├── sample.xlsx                   # Excel d'exemple
│   └── mapping_example.json          # Mapping d'exemple
│
└── actual/                           # VOS DONNÉES (à créer)
    ├── excel_sources/                # Vos fichiers Excel
    │   ├── data.xlsx
    │   ├── indicators.xlsx
    │   └── ...
    │
    └── mappings/                     # Vos fichiers de mapping
        └── mapping.json              # Mapping principal
```

## Prochaines étapes

1. Fournir vos fichiers Excel
2. Fournir vos exemples de transformation de formules
3. Fournir le format de vos indicateurs
4. Générer le mapping automatiquement

## Support

Pour toute question, consultez la documentation ou créez une issue.
