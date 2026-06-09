# 🎓 ACS — Générateur de Bulletins

Application web pour générer automatiquement les bulletins PDF des élèves d'Akroum College of Sciences.

---

## 📁 Structure du projet

```
acs_bulletins_app/
├── app.py              ← Interface Streamlit (UI)
├── generator.py        ← Moteur de génération (logique)
├── requirements.txt    ← Dépendances Python
└── README.md           ← Ce fichier
```

---

## 🚀 Option A — Hébergement gratuit sur Streamlit Cloud

**Idéal pour partager avec toute l'école via un simple lien.**

### Étapes :

1. **Créer un compte GitHub** sur [github.com](https://github.com) (gratuit)

2. **Créer un dépôt** nommé `acs-bulletins`
   - Cliquer "New repository" → nommer `acs-bulletins` → Public → Create

3. **Uploader les 4 fichiers** dans le dépôt :
   - `app.py`
   - `generator.py`
   - `requirements.txt`
   - `README.md`

4. **Créer un compte Streamlit Cloud** sur [share.streamlit.io](https://share.streamlit.io)
   - Se connecter avec GitHub
   - Cliquer "New app"
   - Sélectionner le dépôt `acs-bulletins`
   - Main file path : `app.py`
   - Cliquer "Deploy"

5. **Partager le lien** généré (ex: `https://acs-bulletins.streamlit.app`)

> ⏱️ Déploiement initial : ~3 minutes  
> 💰 Coût : **Gratuit**

---

## 💻 Option B — Lancement local (sur votre ordinateur)

### Prérequis :
- Python 3.9+ installé
- LibreOffice installé ([libreoffice.org](https://www.libreoffice.org))

### Installation :

```bash
# 1. Installer les dépendances Python
pip install -r requirements.txt

# 2. Lancer l'application
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur sur `http://localhost:8501`

---

## ☁️ Option C — Serveur privé (Render.com)

Pour un usage permanent sans limite de temps.

1. Créer un compte sur [render.com](https://render.com) (gratuit)
2. "New Web Service" → connecter GitHub → sélectionner `acs-bulletins`
3. Build command : `pip install -r requirements.txt`
4. Start command : `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5. Deploy

---

## 📋 Utilisation de l'application

1. **Initialiser** LibreOffice + Calibri (une seule fois par session)
2. **Uploader** le fichier Excel des notes
3. **Uploader** les templates Excel (un par groupe)
4. **Sélectionner** les groupes à générer
5. **Cliquer** "Générer"
6. **Télécharger** le ZIP contenant tous les PDFs

---

## 📊 Format du fichier notes

Le fichier Excel doit contenir des feuilles nommées :
- `S1 Notes EB1-2`
- `S1 Notes EB3-6`
- `S1 Notes EB7`
- `S1 Notes EB8`
- `S1 Notes EB9`
- (même pour S2 et S3)

---

## 🔧 Groupes et templates

| Groupe | Template                  | Colonnes |
|--------|---------------------------|----------|
| EB1-2  | template-eb-1to2.xlsx     | 10       |
| EB3-6  | template-eb-3to6.xlsx     | 12       |
| EB7    | template-eb7.xlsx         | 10       |
| EB8    | template-eb8.xlsx         | 11       |
| EB9    | template-eb9.xlsx         | 10       |

---

## ❓ Questions fréquentes

**Q: L'application tourne-t-elle sur Windows ?**  
R: Oui en local si LibreOffice est installé. Sur Streamlit Cloud c'est Linux (géré automatiquement).

**Q: Les données des élèves sont-elles sécurisées ?**  
R: Sur Streamlit Cloud les fichiers uploadés sont temporaires et supprimés après la session. Pour plus de sécurité utilisez l'Option B (local) ou C (serveur privé).

**Q: Combien de bulletins peut-on générer ?**  
R: Pas de limite. ~2-3 secondes par bulletin.

---

*ACS — Akroum College of Sciences · مدرسة أكروم للعلوم · 2025-2026*
