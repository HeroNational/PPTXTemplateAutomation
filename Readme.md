# PptxGeneratorApp

## 📝 Description

Application console .NET pour générer automatiquement des présentations PowerPoint et PDF à partir d'un template et de données CSV.

## ✨ Fonctionnalités

- Génération de fichiers PPTX à partir d'un template
- Remplacement dynamique des balises par les données CSV
- Conversion automatique en PDF via LibreOffice
- Traitement par lots

## 🔧 Prérequis

- Windows 10/11
- .NET 6.0 Runtime
- LibreOffice
- Visual Studio 2022 ou VS Code

## 📦 Installation

1. Clonez le repository :

```bash
git clone <url-du-repo>
cd PptxGeneratorApp
```

2. Restaurez les packages :

```bash
dotnet restore
```

## 🚀 Utilisation

### Structure des fichiers

- `template.pptx` : Template avec balises `[[BALISE]]`
- `data.csv` : Données source
- Dossiers de sortie :
  - `generated_Orga_pptx/` : Fichiers PPTX générés
  - `generated_Orga_pdf/` : Fichiers PDF générés

### Exécution

```bash
dotnet run
```

## 🛠️ Configuration

### Format CSV

```csv
NOM_COMPLET,AUTRE
John Doe,Sujet 1
```

### Exemple de Balises Template

- `[[VOTRE_BALISE]]`
- `[[SUJET]]`

## 📦 Dépendances requises

### Packages NuGet

Installez les packages via la Console du Gestionnaire de Packages NuGet :

```powershell
Install-Package DocumentFormat.OpenXml -Version 3.3.0
Install-Package CsvHelper -Version 33.1.0
```

Ou via la CLI .NET :

```bash
dotnet add package DocumentFormat.OpenXml --version 3.3.0
dotnet add package CsvHelper --version 33.1.0
```

### Logiciels requis

#### LibreOffice

1. Téléchargez LibreOffice depuis [le site officiel](https://www.libreoffice.org/download/download-libreoffice/)
2. Installez en gardant les options par défaut
3. Vérifiez que le chemin est correct dans `Program.cs` :

```csharp
private const string LIBREOFFICE_PATH = "soffice";
```

#### .NET 6.0 SDK

1. Téléchargez le SDK .NET 6.0 depuis [le site Microsoft](https://dotnet.microsoft.com/download/dotnet/6.0)
2. Installez-le sur votre machine
3. Vérifiez l'installation :

```bash
dotnet --version
```

### Vérification des dépendances

Pour vérifier que toutes les dépendances sont correctement installées :

```bash
dotnet restore
dotnet build
```

Si vous obtenez des erreurs, vérifiez que :

- Le fichier `PptxGeneratorApp.csproj` contient les bonnes références
- LibreOffice est accessible depuis la ligne de commande
- Le SDK .NET 6.0

## 🔍 Dépannage

### Problèmes courants

1. **LibreOffice non trouvé**

   - Vérifiez l'installation
   - Ajustez le chemin dans `LIBREOFFICE_PATH`
2. **Erreurs de conversion PDF**

   - Vérifiez les permissions
   - Consultez les logs

## 📄 License

MIT License

## 🤝 Support

Pour tout problème, ouvrez une issue sur le repository.

## 📝 Changelog

### v1.0.0

- Version initiale
- Génération PPTX et PDF
- Support du traitement par lot
