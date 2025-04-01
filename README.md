# **Automatisation des emails avec Excel & VBA 📩**

## 📌 **Présentation**
Ce projet montre comment **envoyer des emails automatiquement** depuis **Excel VBA** en utilisant **Microsoft Outlook**. 

✅ **Envoi de rapports de stock ou d’expédition en supply chain**  
✅ **Personnalisation des emails selon les données Excel**  
✅ **Automatisation des relances fournisseurs et clients**  

---

## 📥 **Installation**
### 1️⃣ **Importer le module VBA (`VBA_Email_Auto.bas`)**
1. Ouvrir Excel et accéder à l’éditeur VBA (`ALT + F11`).
2. Cliquer droit sur "Modules" > **Importer un fichier**.
3. Sélectionner `VBA_Email_Auto.bas`.

### 2️⃣ **Configurer Outlook**
- Outlook doit être installé et configuré sur votre PC.
- Autoriser l’exécution de macros VBA.

### 3️⃣ **Préparer votre fichier Excel**
- La colonne **A** doit contenir les emails des destinataires.
- La colonne **B** doit contenir le message personnalisé.
- La colonne **C** peut contenir une pièce jointe (optionnel).

### 4️⃣ **Exécuter la macro**
- Lancer `SendEmailsFromExcel` pour envoyer les emails.

---

## 📊 **Cas d’usage concret : Supply Chain**
🔹 **Avant** : Envoi manuel des alertes de stock faible aux fournisseurs.  
🔹 **Après** : **Automatisation des relances fournisseurs** avec Excel & VBA.  

---

## 📌 **Personnalisation**
- Modifier l’objet et le corps du mail dans le script VBA.
- Ajouter des pièces jointes dynamiques en fonction du produit.

📣 **Des suggestions ou améliorations ?** Ouvrez une **issue** ou proposez une **pull request** ! 🚀
