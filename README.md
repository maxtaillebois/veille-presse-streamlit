# 📰 Outil de Veille Presse

Application web locale pour analyser des articles de presse au format PDF, rechercher des mots-clés et compiler les articles sélectionnés.

## ✨ Fonctionnalités

- **Analyse de PDF** : Extraction automatique du texte des fichiers PDF
- **Recherche de mots-clés** : Configurable avec mise en évidence des occurrences
- **Extraction de métadonnées** : Titre, média, date de publication
- **Résumé automatique** : Génération intelligente de résumés basés sur les mentions
- **Interface de sélection** : Choisissez les articles pertinents à conserver
- **Compilation PDF** : Fusionnez les articles sélectionnés en un seul document
- **Récapitulatif TXT** : Liste des articles par ordre antéchronologique
- **Correction automatique** : Gestion des PDFs avec problèmes d'encodage

## 🖥️ Capture d'écran

L'interface se présente sous forme d'application web locale accessible via votre navigateur.

## 📋 Prérequis

- Python 3.8 ou supérieur
- pip (gestionnaire de paquets Python)

## 🚀 Installation

### Option 1 : Cloner le dépôt

```bash
git clone https://github.com/votre-utilisateur/veille-presse.git
cd veille-presse
pip install -r requirements.txt
```

### Option 2 : Téléchargement direct

1. Téléchargez les fichiers du projet
2. Ouvrez un terminal dans le dossier
3. Installez les dépendances :
   ```bash
   pip install -r requirements.txt
   ```

## 🎯 Utilisation

### Windows

Double-cliquez sur `LANCER_VEILLE_PRESSE.bat`

### macOS / Linux

```bash
chmod +x LANCER_VEILLE_PRESSE.command
./LANCER_VEILLE_PRESSE.command
```

Ou directement :

```bash
streamlit run veille_presse_app.py
```

L'application s'ouvre automatiquement dans votre navigateur à l'adresse `http://localhost:8501`

## ⚙️ Configuration

Dans la barre latérale, vous pouvez configurer :

- **Source des PDF** : Téléverser des fichiers ou scanner un dossier
- **Mots-clés** : Liste des termes à rechercher (un par ligne)
- **Taille du contexte** : Nombre de caractères affichés autour des mots-clés
- **Dossier de sortie** : Emplacement de la compilation finale

### Mots-clés par défaut

```
Yannick Borde
Procivis
Immo de France
Maisons d'en France
```

## 📁 Structure du projet

```
veille-presse/
├── veille_presse_app.py        # Application principale
├── requirements.txt            # Dépendances Python
├── LANCER_VEILLE_PRESSE.bat    # Script de lancement Windows
├── LANCER_VEILLE_PRESSE.command # Script de lancement macOS/Linux
├── README.md                   # Ce fichier
└── GUIDE_INSTALLATION_WINDOWS.md # Guide détaillé pour Windows
```

## 📤 Fichiers générés

Après compilation, l'application génère :

- `compilation_veille_YYYYMMDD.pdf` : PDF contenant tous les articles sélectionnés
- `compilation_veille_YYYYMMDD_recap.txt` : Récapitulatif au format texte

Format du récapitulatif :
```
Média - "Titre de l'article" - JJ/MM/AAAA
```

## 🔧 Dépendances

| Package | Version | Description |
|---------|---------|-------------|
| streamlit | ≥1.28.0 | Interface web |
| pdfplumber | ≥0.10.0 | Extraction de texte PDF |
| pypdf | ≥3.17.0 | Manipulation de PDF |
| reportlab | ≥4.0.0 | Génération de PDF |

## 🐛 Résolution de problèmes

### "Python n'est pas reconnu..."
Réinstallez Python en cochant **"Add Python to PATH"** lors de l'installation.

### "pip n'est pas reconnu..."
Utilisez : `python -m pip install -r requirements.txt`

### Le navigateur ne s'ouvre pas
Ouvrez manuellement : `http://localhost:8501`

### PDF avec texte illisible (mots collés)
L'application détecte et corrige automatiquement ces problèmes. Un avertissement s'affiche pour les PDFs concernés.

## 📝 Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.

## 👤 Auteur

Développé pour faciliter la veille presse et le suivi médiatique.

---

*Dernière mise à jour : Février 2026*
