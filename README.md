# Les Berthos — enquête généalogique

Ce dépôt contient tout le projet : le tableau de travail et le site que
la famille consulte.

## Comment ça marche (pour vous)

Il n'y a qu'**un seul fichier à faire vivre** : `data/Genealogie_MEME.xlsx`.
Tout le reste (le site) se reconstruit **tout seul**.

1. Vous modifiez `data/Genealogie_MEME.xlsx` sur votre ordinateur
   (ajout d'une personne, d'une piste, d'une petite histoire...).
2. Sur la page du dépôt, sur github.com : bouton **Add file → Upload files**,
   vous glissez le fichier `Genealogie_MEME.xlsx` mis à jour (même nom,
   même emplacement `data/`), puis **Commit changes**.
3. Quelques dizaines de secondes plus tard, le site est mis à jour tout seul
   (onglet **Actions** du dépôt pour suivre la progression si besoin).

Pas de logiciel à installer, pas de ligne de commande.

Si un jour vous préférez travailler avec un outil qui simplifie l'étape 2
(glisser-déposer avec un historique visuel), **GitHub Desktop** fait
exactement la même chose en plus confortable — mais ce n'est pas
obligatoire.

## Les trois onglets du tableau qui comptent

- **Registre des personnes** : une ligne = une personne. Les colonnes
  « Petite histoire » et « Piste de recherche » sont celles qui
  alimentent le site.
- **Chronologie** : les événements datés.
- **Guide** : la méthode, à relire si vous vous perdez dans les colonnes.

## Comment la famille participe

Le site (une fois publié, voir plus bas) affiche en bas de chaque page un
bouton **Envoyer une piste par mail**. Vos tontons et tantes cliquent,
ça ouvre leur messagerie habituelle avec l'adresse `FamBertho@proton.me`
déjà pré-remplie, ils n'ont rien d'autre à faire. Vous recevez leurs
souvenirs par mail comme d'habitude, et c'est vous qui les reportez
ensuite dans le tableau.

## Photos de famille : Postimage, pas GitHub

Pour ne pas alourdir le dépôt, les photos de personnes ne sont **pas**
stockées ici. Le circuit :

1. Vous prenez les tirages photo au téléphone (lumière naturelle, pas de
   scan nécessaire).
2. Vous les déposez sur [postimage.org](https://postimage.org) (gratuit,
   pas de compte requis).
3. Vous collez le(s) lien(s) directs dans la colonne « Photos (liens
   Postimage, séparés par ; ) » du Registre des personnes — le premier
   lien devient le portrait de la fiche, les suivants forment sa galerie.

## Documents officiels : Proton Drive, en privé

Livrets de famille et actes de naissance ne sont **jamais** publiés,
même en scan flou — ils peuvent concerner des personnes encore vivantes.
Ils restent uniquement sur Proton Drive, en privé. La page
`archives-privees.html` du site explique à la famille comment en
demander une copie par mail.

## Activer la publication du site (à faire une seule fois)

Sur github.com, dans ce dépôt : **Settings → Pages → Source : Deploy from
a branch**, branche `main`, dossier `/docs`. Le site sera alors visible à
une adresse du type `https://VOTRE-PSEUDO.github.io/NOM-DU-DEPOT/`.

L'adresse mail (`FamBertho@proton.me`) et le nom du site (« Les Berthos »)
sont déjà réglés en haut de `scripts/generate_site.py` (variables
`CONTACT_EMAIL` et `SITE_TITLE`) — à changer là si besoin.

## Régénérer le site vous-même (facultatif)

Si vous êtes à l'aise et voulez voir le résultat avant de le publier :

```bash
pip install -r scripts/requirements.txt
python3 scripts/generate_site.py
```

Cela met à jour le dossier `docs/`. Mais ce n'est pas nécessaire au
quotidien : l'automatisation GitHub (dossier `.github/workflows/`) le
fait pour vous à chaque envoi du tableau.
