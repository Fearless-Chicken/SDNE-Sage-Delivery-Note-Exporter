Extraction BL Sage - README

Description

Ce projet contient trois fichiers Python qui permettent d'extraire des données depuis Sage selon différents critères :

1. extract_ALL.py : Ce script extrait toutes les données, quelle que soit la date.
2. extract_DateToDate.py : Ce script demande une date de début et une date de fin, puis extrait les données correspondant à cette période.
3. extract_Menu.py : Ce script propose un menu interactif permettant de :
    - Tout extraire
    - Extraire des données pour une période donnée (date de début et de fin)
    - Extraire toutes les données pour un mois ou une année spécifique

Prérequis

Installation de Python
Assurez-vous d'avoir Python installé sur votre machine. Vous pouvez le télécharger depuis https://www.python.org/.

Modules Python nécessaires
Les modules suivants doivent être installés pour exécuter les scripts :
- pick
- pyodbc
- csv

Vous pouvez les installer en exécutant le script suivant : "InstallDep.bat"

Configuration

Avant d'exécuter les scripts, vous devez configurer le fichier config.conf. Ce fichier contient les paramètres nécessaires pour se connecter à la base de données Sage. Voici les paramètres à configurer :

- server : Adresse du serveur de base de données
- database : Nom de la base de données Sage
- username : Nom d'utilisateur pour la connexion
- password : Mot de passe pour la connexion

Exemple de contenu pour config.conf :
    server : """CPRO21-185\SAGE100"""
    database : 'BIJOU'
    username : 'sa'
    password : 'Koesiosa2019'