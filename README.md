# 📄 RPA - Traitement Automatisé de Contrats pour La Poste

## ✨ Description

Ce projet implémente un système de **RPA (Robotic Process Automation)** pour **La Poste**, capable de traiter automatiquement une volumétrie importante de contrats (3000 à 5000 par semaine). Grâce à cette solution, des tâches récurrentes et à faible valeur ajoutée peuvent être externalisées de manière fiable, tout en réduisant considérablement les coûts opérationnels. Ce RPA est :

- 🛠️ **Modulable** : adapté à différents processus internes de l'entreprise.
- 🔄 **Flexible** : paramétrable pour répondre à de nouvelles exigences métier.
- 🚨 **Connecté** : équipé d’un système de journalisation avancé qui notifie les équipes en cas de dysfonctionnements.

---

## 🏗️ Architecture

- **Frontend** : 🖥️ [Electron.js](https://www.electronjs.org/) pour l'interface utilisateur.
- **Backend** : 🐍 Python avec les bibliothèques suivantes :
  - 🕹️ `Selenium` : Automatisation des interactions avec les interfaces web.
  - 📊 `pandas` : Manipulation et analyse des données.
  - 🧠 `psutil` : Gestion des processus système.
  - ⚡ `ThreadPoolExecutor` : Multi-traitement pour maximiser les performances.

---

## 🚀 Fonctionnalités

1. **🔍 Traitement des contrats :**
   - Extraction des numéros de contrats depuis des fichiers Excel 📂.
   - Soumission automatisée des contrats sur les plateformes web internes 🧾.
   - Sauvegarde des résultats et rapports d’erreurs dans des fichiers CSV et JSON 📑.

2. **⚙️ Gestion des processus :**
   - Optimisation des ressources via un pool de WebDrivers 🚗.
   - Surveillance des processus pour éviter la surcharge CPU 📉.

3. **📡 Système d’alerte :**
   - Journalisation centralisée des événements 📝.
   - Alertes en cas de dysfonctionnement via des fichiers de logs 🚨.

4. **📦 Modularité :**
   - Ajout facile de nouveaux processus grâce à une architecture modulaire 🔧.

---
