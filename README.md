# SheetToPDF

Ce projet fournit un ensemble de scripts permettant de convertir des feuilles Excel en documents PDF. Le processus est automatisé à l'aide d'un script batch (`converter.bat`) qui prépare l'environnement et lance le script Python (`SheetToPDF.py`) pour effectuer la conversion.

---

## Licence

Ce projet est sous licence MIT. Cela signifie que vous êtes libre de l'utiliser, de le modifier et de le distribuer, à condition d'inclure l'avis de licence original et les conditions de la licence dans toute copie ou version substantielle du logiciel.

Pour plus de détails sur la licence, veuillez consulter le fichier `LICENSE` inclus dans ce projet.

---

## Pour les utilisateurs non techniques

### Prérequis

- Windows 7 ou supérieur.
- Accès à Internet pour l'installation automatique de Python, si nécessaire.

### Comment utiliser

1. **Lancer le script `converter.bat`** : Double-cliquez sur le fichier `converter.bat`. Une fenêtre de commande s'ouvrira.
2. **Sélection du fichier Excel** : Lorsque le script demande le chemin du fichier Excel à convertir, glissez-déposez le fichier dans la fenêtre de commande, puis appuyez sur Entrée.
3. **Sélection du dossier de sortie** : Lorsque le script demande où sauvegarder les PDFs, glissez-déposez le dossier de destination dans la fenêtre de commande. Si vous souhaitez utiliser le répertoire courant, appuyez simplement sur Entrée sans entrer de chemin.
4. **Attente de la conversion** : Le script effectuera la conversion et affichera un message une fois terminée. Les fichiers PDF seront situés dans le dossier spécifié.

### En cas d'erreur

Si une erreur survient, consultez le fichier `conversion_errors.log` généré dans le même dossier que le script pour plus de détails.

## Pour les développeurs

### Dépendances

- Python 3.x
- Bibliothèques Python : `win32com.client`
- Le script utilise `pip` pour installer automatiquement les dépendances nécessaires.

### Fonctionnement du script

**`converter.bat`** :
- Vérifie la présence de Python et l'installe si nécessaire.
- Installe `pip` et les dépendances requises.
- Demande à l'utilisateur le chemin du fichier Excel et le dossier de sortie.
- Exécute `SheetToPDF.py` avec les chemins fournis comme arguments.

**`SheetToPDF.py`** :
- Utilise `win32com.client` pour ouvrir le fichier Excel spécifié.
- Convertit chaque feuille du classeur en un fichier PDF séparé, en appliquant des paramètres de mise en page prédéfinis.
- Sauvegarde chaque PDF dans le dossier de sortie spécifié.

### Utilisation en ligne de commande

Pour exécuter directement `SheetToPDF.py` sans passer par `converter.bat`, utilisez la commande suivante dans un terminal :

```bash
python SheetToPDF.py "<chemin vers le fichier Excel>" "<dossier de sortie>"
```

Assurez-vous de remplacer `<chemin vers le fichier Excel>` et `<dossier de sortie>` par les chemins appropriés.

### Gestion des erreurs

Le script Python affiche les erreurs dans la console et les enregistre dans `conversion_errors.log` lorsqu'exécuté via `converter.bat`.

---
