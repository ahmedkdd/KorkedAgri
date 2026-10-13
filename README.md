# KorkedAgri

Suivi simple des charges et bénéfices de la ferme d'oranges (Android, hors ligne).
Les données sont enregistrées sur le téléphone.

## Obtenir l'APK sans rien installer (GitHub Actions)
1. Crée un dépôt GitHub et pousse ce dossier.
2. Onglet **Actions** → workflow **Build APK** (il se lance tout seul au push).
3. À la fin (~3 min), télécharge l'artefact **KorkedAgri-apk** → `app-debug.apk`.
4. Envoie-le sur le téléphone et installe-le (autoriser « sources inconnues »).

## Avec Android Studio
Ouvrir le dossier → laisser Gradle synchroniser → **Build > Build APK(s)**.

## Modifier l'appli
Tout l'écran est dans `app/src/main/assets/index.html`.
Pour changer de version, augmente `versionCode` dans `app/build.gradle`
(sinon Android refuse la mise à jour par-dessus l'ancienne).

## iPhone (web app installable)
La version iPhone est dans `docs/` et publiée par GitHub Pages :
https://ahmedkdd.github.io/KorkedAgri/
Ouvrir dans **Safari** → bouton Partager → **Sur l'écran d'accueil**.
Fonctionne hors ligne ; les données restent sur le téléphone.
