# Document Conversion Toolkit

Ce toolkit permet de convertir facilement des documents entre différents formats, notamment de feuilles Excel en PDF, de documents Word en PDF, et inversement de PDF en documents Word.

## Pour les Utilisateurs Non Techniques

### Comment Utiliser

1. **Prérequis :** Assurez-vous que Python 3 est installé sur votre ordinateur. Si ce n'est pas le cas, le script d'installation tentera de l'installer pour vous.
2. **Lancement :** Double-cliquez sur le fichier `convert_tool.bat`. Une fenêtre de commande s'ouvrira.
3. **Conversion :** Suivez les instructions à l'écran. Vous serez invité à glisser-déposer les fichiers ou dossiers que vous souhaitez convertir. Ensuite, indiquez le dossier où vous souhaitez sauvegarder les résultats. Si vous ne spécifiez pas de dossier, les fichiers seront sauvegardés dans un nouveau dossier `ResultatConversion` dans votre répertoire actuel.
4. **Résultats :** Une fois la conversion terminée, vous trouverez vos fichiers convertis dans le dossier de sortie spécifié.

### Problèmes Communs

- **Python non installé :** Le script tentera d'installer Python automatiquement. Suivez les instructions à l'écran si une intervention est nécessaire.
- **Dossier de sortie non spécifié :** Si aucun chemin de sortie n'est fourni, les fichiers seront sauvegardés dans un dossier `ResultatConversion` par défaut.

## Pour les Développeurs

### Configuration

- **Python 3.x** est requis. Le script vérifie sa présence et tente une installation automatique si nécessaire.
- **Dépendances Python :** `pywin32`, `comtypes`, `pdf2docx`. Elles sont installées automatiquement par le script.

### Utilisation

Le script principal `desk_tool_converter.py` peut être utilisé indépendamment avec les arguments de ligne de commande suivants :

```sh
python desk_tool_converter.py <source_path> <output_folder>
```

- `<source_path>` : Chemin vers le fichier ou le dossier à convertir.
- `<output_folder>` : Chemin vers le dossier de sortie pour les fichiers convertis.

### Script Batch

`convert_tool.bat` sert d'interface pour faciliter l'utilisation du script Python par les utilisateurs finaux. Il gère l'installation des prérequis, la saisie utilisateur et l'appel au script Python avec les paramètres appropriés.

### Fonctions de Conversion

- `convert_sheet_to_pdf(sheet, output_path)` : Convertit une feuille Excel en PDF.
- `convert_excel_to_pdf(source_path, output_folder)` : Convertit un classeur Excel en PDF (toutes les feuilles).
- `convert_word_to_pdf(source_path, output_path)` : Convertit un document Word en PDF.
- `convert_pdf_to_word(source_path, output_path)` : Convertit un PDF en document Word.

### Traitement des Erreurs

Les erreurs rencontrées pendant la conversion sont enregistrées dans `conversion_errors.log` pour un débogage facile.

---

Ce README offre un guide complet pour les utilisateurs de tous niveaux et fournit les détails nécessaires pour les développeurs souhaitant comprendre ou étendre la fonctionnalité des scripts.
