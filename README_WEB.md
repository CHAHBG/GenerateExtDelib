# 🌍 Procasef Automation Web (v7.0)

Version Web de l'outil d'automatisation, conçue pour être hébergée sans serveur (Netlify/GitHub Pages).

## 🚀 Fonctionnement
Cette application utilise **Pyodide** pour exécuter le moteur Python **directement dans votre navigateur**.
- Pas de backend serveur requis (100% Client-Side).
- Les fichiers ne quittent pas votre ordinateur (confidentialité).
- Performance : Dépend de la puissance de votre machine (CPU/RAM).

## 📦 Déploiement

### Option A : Netlify (Recommandé - Ultra Simple)
1. Allez sur **[Netlify Drop](https://app.netlify.com/drop)**.
2. Glissez-déposez le dossier `dist` situé dans `Web_App/dist`.
3. C'est en ligne ! 🎉

### Option B : GitHub Pages
1. Poussez le dossier `Web_App` sur GitHub.
2. Configurez une Action pour build (ou punsh le contenu de `dist` sur une branche `gh-pages`).
*Note : Si hébergé sous `/mon-repo/`, ajustez la `base` dans `vite.config.ts`.*

## 🛠️ Développement Local
1. `cd Web_App`
2. `npm install`
3. `npm run dev` (Démarre le serveur local)

## 📁 Structure
- `/public/python/generate_web.py` : Le cerveau Python (adapté pour le web).
- `/src/App.tsx` : L'interface React.
- `/src/index.css` : Styles (TailwindCSS).
