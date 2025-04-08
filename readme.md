# 📦 SDNE - Sage Delivery Note Exporter

## 🧾 Présentation

Ce programme est un outil **interne** permettant de :
- Se connecter à une base Sage via SQL Server
- Extraire les bons de livraison non exportés
- Générer automatiquement un fichier `.csv`
- Marquer les BL comme exportés dans la base
- Conserver un historique des logs par mois

Le tout est exécuté **en ligne de commande** via un exécutable CLI (fichier `.exe`).

---

## 🗂️ Arborescence du projet

```
V3 - Version finale\Extraction BL Sage
├── app
│   ├── config
│   │   ├── config.conf           # Configuration de la base de données
│   │   └── requirements.txt      # Dépendances Python
│   ├── controller
│   │   └── main.py               # Script principal du programme
│   ├── extracted                 # CSV générés lors des exports
│   ├── log
│   │   ├── template.log          # Modèle de fichier log
│   │   ├── log<Mois>.log         # Fichier de log du mois courant
│   │   └── Ancienne logs         # Logs des mois précédents
│   └── SQL
│       ├── ExtractAllUnexported.sql
│       ├── ExtractD2D.sql
│       └── UpdateBL.sql
├── install
│   ├── install.bat              # Script d’installation
│   └── Extract.bat              # Lancement de l’extraction
└── readme.md                    # Documentation du projet
```

---

## ⚙️ Dépendances

- Python 3.10+
- `pyodbc`
- Driver ODBC : **ODBC Driver 17 for SQL Server**

### Installation des dépendances

```bash
pip install -r app/config/requirements.txt
```

---

## 🔧 Configuration

Dans `app/config/config.conf` :

```
Serveur   : NOM_DU_SERVEUR
Database  : NOM_DE_LA_BASE
Username  : UTILISATEUR
Password  : MOTDEPASSE
```

Les lignes doivent être dans cet ordre, avec le séparateur ` : ` et sans guillemets autour des valeurs.
_(nb : un exemple est déjà présent)_

---

## ▶️ Utilisation

Il suffit de double-cliquer sur `Extract.exe` 

Le programme :
1. Se connecte à la base de données
2. Extrait les BL non exportés
3. Génère un fichier `.csv` dans `app/extracted`
4. Met à jour le statut dans la base
5. Logue les actions dans `app/log`

---

## 🧠 Notes

- Si aucun bon de livraison n’est détecté, le programme ajoute une entrée de log indiquant qu’aucune extraction n’était nécessaire évitant de faire tourner le programme dans le vide".
- Les fichiers de log sont archivés automatiquement à chaque changement de mois.

---

## 🧑‍💻 Copyrights

_Développé par Fearless-Chicken – KOESIO © 2025_
📧 blachere.theo2005@gmail.com
🔗 [github.com/Fearless-Chicken](https://github.com/Fearless-Chicken)

Toute copie, modification, distribution ou utilisation non autorisée, totale ou partielle, est interdite sans l'accord écrit préalable de l'auteur.
