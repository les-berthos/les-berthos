#!/usr/bin/env python3
"""
Génère le site statique (dossier /docs) à partir de data/Genealogie_MEME.xlsx.
Ne modifie jamais le xlsx. À relancer à chaque mise à jour du tableau
(en pratique : automatique via .github/workflows/build-site.yml).
"""
import re
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "Genealogie_MEME.xlsx"
DOCS = ROOT / "docs"
PERSONNES_DIR = DOCS / "personnes"

CONTACT_EMAIL = "FamBertho@proton.me"
SITE_TITLE = "Les Berthos"
SITE_TAGLINE = "Une enquête de famille, page après page — mise à jour au fil des découvertes."


def slug(text: str) -> str:
    text = str(text).strip()
    text = re.sub(r"[^\w\-]+", "-", text, flags=re.UNICODE)
    return re.sub(r"-+", "-", text).strip("-").upper()


def is_blank(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan" or s == "?"


def split_list(v):
    if is_blank(v):
        return []
    return [x.strip() for x in re.split(r"[;\n]", str(v)) if x.strip()]


def load_data():
    personnes = pd.read_excel(XLSX, sheet_name="Registre des personnes", dtype=str)
    personnes = personnes.dropna(how="all")
    if "ID Personne" in personnes.columns:
        personnes = personnes[~personnes["ID Personne"].astype(str).str.contains("EXEMPLE", na=False)]
    personnes = personnes.fillna("")
    return personnes


def person_display_name(row) -> str:
    prenom = row.get("Prénom(s)", "")
    nom = row.get("Nom", "")
    if is_blank(prenom) and is_blank(nom):
        return row.get("ID Personne", "Personne inconnue")
    return f"{'' if is_blank(prenom) else prenom} {'' if is_blank(nom) else nom}".strip()


def by_id_index(personnes: pd.DataFrame):
    idx = {}
    for _, row in personnes.iterrows():
        pid = str(row.get("ID Personne", "")).strip()
        if pid:
            idx[pid] = row
    return idx


def find_children(pid, personnes: pd.DataFrame):
    kids = []
    for _, row in personnes.iterrows():
        pere = str(row.get("Père (ID Personne)", "")).strip()
        mere = str(row.get("Mère (ID Personne)", "")).strip()
        if pere == pid or mere == pid:
            kids.append(row)
    return kids


def compute_leads(personnes: pd.DataFrame):
    leads = []
    key_fields = ["Date de naissance", "Lieu de naissance", "Date de décès", "Lieu de décès"]
    for _, row in personnes.iterrows():
        pid = row.get("ID Personne", "")
        name = person_display_name(row)
        piste = row.get("Piste de recherche", "")
        if not is_blank(piste):
            leads.append({"id": pid, "name": name, "text": piste})
        else:
            manquants = [f for f in key_fields if is_blank(row.get(f, ""))]
            if manquants and row.get("Statut (confirmé / hypothèse)", "") != "confirmé":
                leads.append({
                    "id": pid, "name": name,
                    "text": f"Champs encore inconnus : {', '.join(manquants)}."
                })
    return leads


CSS = """
:root{
  --paper:#efe7d8; --paper-dark:#e4d9c3; --ink:#3a2e22; --ink-soft:#6b5c48;
  --wax:#8c2f39; --brass:#b8863b; --sage:#5b7065; --line:#c9bda3; --cream:#f6efdd;
}
*{box-sizing:border-box}
body{
  margin:0; color:var(--ink);
  font-family:'Public Sans', Arial, sans-serif; font-size:18px; line-height:1.55;
  background-image:
    radial-gradient(var(--paper-dark) 0.6px, transparent 0.6px),
    linear-gradient(rgba(239,231,216,.87), rgba(239,231,216,.87)),
    url('https://i.postimg.cc/vmTKF4ZQ/arriere-plan-Bertho.png');
  background-size: 14px 14px, cover, cover;
  background-position: 0 0, center, center;
  background-repeat: repeat, no-repeat, no-repeat;
  background-attachment: scroll, fixed, fixed;
}
.hero{
  position:relative;
  background-image: linear-gradient(180deg, rgba(20,14,8,.4) 0%, rgba(20,14,8,.68) 100%);
  padding:4.5rem 1.5rem 2.2rem; text-align:center;
}
.hero h1{
  font-family:'Playfair Display', Georgia, serif; font-size:3rem; margin:0 0 .4rem;
  color:var(--cream); letter-spacing:1px; text-shadow:0 2px 10px rgba(0,0,0,.5);
}
.hero p{margin:0; color:var(--cream); opacity:.9; font-size:1.05rem; text-shadow:0 1px 6px rgba(0,0,0,.5)}
nav.site{
  display:flex; justify-content:center; gap:1.6rem; padding:.9rem 0; background:var(--paper-dark);
  border-bottom:1px solid var(--line); font-weight:600; letter-spacing:.3px; flex-wrap:wrap;
}
nav.site a{color:var(--ink); text-decoration:none; border-bottom:2px solid transparent; padding-bottom:2px}
nav.site a:hover, nav.site a:focus{border-bottom-color:var(--wax)}
main{max-width:880px; margin:0 auto; padding:2.2rem 1.5rem 4rem}
h2{font-family:'Playfair Display', Georgia, serif; color:var(--wax); font-size:1.7rem; margin-top:2.2rem}
.card{
  background:#faf6ec; border:1px solid var(--line); border-radius:3px;
  padding:1.4rem 1.6rem; margin:1.1rem 0; box-shadow:0 1px 0 var(--line);
}
.card a.name{font-family:'Playfair Display', Georgia, serif; font-size:1.3rem; color:var(--ink); text-decoration:none}
.card a.name:hover{color:var(--wax)}
.stamp{
  display:inline-block; font-family:'Special Elite', 'Courier New', monospace; font-size:.72rem;
  text-transform:uppercase; letter-spacing:1.5px; padding:.25rem .55rem; border:2px solid; border-radius:2px;
  transform:rotate(-3deg); margin-left:.6rem; vertical-align:middle;
}
.stamp.confirme{color:var(--sage); border-color:var(--sage)}
.stamp.hypothese{color:var(--brass); border-color:var(--brass)}
.stamp.inconnu{color:var(--wax); border-color:var(--wax)}
.meta{color:var(--ink-soft); font-size:.95rem; margin:.3rem 0 0}
.histoire{margin-top:.7rem}
.piste{
  background:#f3e7c9; border-left:4px solid var(--brass); padding:.7rem 1rem; margin-top:.8rem;
  font-size:.98rem;
}
.piste strong{color:var(--brass)}
.idcard{
  display:flex; gap:1.4rem; background:#faf6ec; border:2px solid var(--ink); border-radius:4px;
  padding:1.3rem; box-shadow:3px 3px 0 var(--line); flex-wrap:wrap;
}
.idcard .photo{
  width:140px; height:170px; flex-shrink:0; background:var(--paper-dark); border:1px solid var(--ink-soft);
  object-fit:cover; display:flex; align-items:center; justify-content:center; color:var(--ink-soft);
  font-size:.8rem; text-align:center; padding:.5rem;
}
.idcard .infos{flex:1; min-width:220px}
.idcard .infos h2{margin:0 0 .2rem; font-size:1.5rem}
.idcard .infos .eyebrow{
  font-family:'Special Elite','Courier New',monospace; font-size:.7rem; letter-spacing:2px;
  color:var(--ink-soft); text-transform:uppercase; margin-bottom:.4rem;
}
.idcard dl{display:grid; grid-template-columns:auto 1fr; gap:.25rem .8rem; margin:.7rem 0 0; font-size:.95rem}
.idcard dt{color:var(--ink-soft)}
.idcard dd{margin:0}
.gallery{display:flex; flex-wrap:wrap; gap:.6rem; margin-top:.9rem}
.gallery a{display:block; width:90px; height:90px; overflow:hidden; border:1px solid var(--line); border-radius:3px}
.gallery img{width:100%; height:100%; object-fit:cover; display:block}
footer.contact{
  margin-top:3rem; padding:1.8rem; text-align:center; background:var(--paper-dark);
  border-top:1px solid var(--line);
}
footer.contact a.btn{
  display:inline-block; margin-top:.6rem; background:var(--wax); color:#faf6ec; text-decoration:none;
  padding:.7rem 1.4rem; border-radius:3px; font-weight:700;
}
footer.contact a.btn:hover{background:#732530}
.empty{color:var(--ink-soft); font-style:italic}
.back{display:inline-block; margin-bottom:1.2rem; color:var(--ink-soft); text-decoration:none}
.back:hover{color:var(--wax)}
.archive-box{background:#faf6ec; border:1px solid var(--line); border-left:4px solid var(--wax); padding:1.2rem 1.4rem; margin:1.2rem 0}
@media (max-width:600px){
  .hero h1{font-size:2.1rem} body{font-size:17px} .idcard{flex-direction:column; align-items:center; text-align:center}
  .idcard dl{grid-template-columns:1fr; text-align:center}
}
"""

FONT_LINKS = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">'
    '<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&'
    'family=Public+Sans:wght@400;600;700&family=Special+Elite&display=swap" rel="stylesheet">'
)


def page(title, body, active="", depth=0):
    prefix = "../" * depth

    def navlink(href, label, key):
        cls = ' style="border-bottom-color:var(--wax)"' if key == active else ""
        return f'<a href="{prefix}{href}"{cls}>{label}</a>'

    nav = f"""
    <nav class="site">
      {navlink('index.html', "Accueil", 'accueil')}
      {navlink('pistes.html', "Pistes de recherche", 'pistes')}
      {navlink('archives-privees.html', "Archives privées", 'archives')}
    </nav>"""
    if active == "accueil" and depth == 0:
        header = f"""<header class="hero"><h1>{SITE_TITLE}</h1><p>{SITE_TAGLINE}</p></header>"""
    else:
        header = f"""<header style="padding:1.6rem 1.5rem; text-align:center; border-bottom:1px solid var(--line);">
                   <a href="{prefix}index.html" style="font-family:'Playfair Display',Georgia,serif; font-size:1.5rem; color:var(--wax); text-decoration:none;">{SITE_TITLE}</a>
                 </header>"""
    return f"""<!doctype html>
<html lang="fr"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} — {SITE_TITLE}</title>
{FONT_LINKS}
<style>{CSS}</style>
</head><body>
{header}
{nav}
<main>{body}</main>
<footer class="contact">
  <p>Une photo, une lettre, un souvenir, une correction ? Écrivez-nous, ça complète directement l'enquête.</p>
  <a class="btn" href="mailto:{CONTACT_EMAIL}?subject=Une%20piste%20pour%20l'arbre%20des%20Berthos">
    Envoyer une piste par mail</a>
</footer>
</body></html>"""


def stamp(statut: str) -> str:
    s = (statut or "").strip().lower()
    if s.startswith("conf"):
        return '<span class="stamp confirme">confirmé</span>'
    if s.startswith("hyp"):
        return '<span class="stamp hypothese">hypothèse</span>'
    return '<span class="stamp inconnu">à enquêter</span>'


def build_index(personnes: pd.DataFrame):
    if personnes.empty:
        body = (
            "<p class='empty'>L'arbre est encore vide — dès que le « Registre des personnes » du "
            "tableau Excel contient une première ligne, cette page affichera la première fiche.</p>"
        )
    else:
        cards = []
        for _, row in personnes.iterrows():
            pid = row.get("ID Personne", "")
            name = person_display_name(row)
            lien = row.get("Lien avec la MEME", "")
            naiss = row.get("Date de naissance") or row.get("Date de naissance estimée") or "?"
            cards.append(f"""
            <div class="card">
              <a class="name" href="personnes/{slug(pid)}.html">{name}</a> {stamp(row.get('Statut (confirmé / hypothèse)'))}
              <p class="meta">{'' if is_blank(lien) else lien + ' · '}née/né {naiss}</p>
            </div>""")
        body = "<h2>Les personnes identifiées jusqu'ici</h2>" + "".join(cards)
    (DOCS / "index.html").write_text(page("Accueil", body, active="accueil", depth=0), encoding="utf-8")


def build_person_pages(personnes: pd.DataFrame):
    PERSONNES_DIR.mkdir(parents=True, exist_ok=True)
    idx = by_id_index(personnes)

    def name_or_id(pid):
        pid = str(pid).strip()
        if not pid:
            return None
        row = idx.get(pid)
        return person_display_name(row) if row is not None else pid

    def link_or_text(pid):
        pid = str(pid).strip()
        label = name_or_id(pid)
        if label is None:
            return None
        if pid in idx:
            return f'<a href="{slug(pid)}.html">{label}</a>'
        return label

    for _, row in personnes.iterrows():
        pid = row.get("ID Personne", "")
        name = person_display_name(row)
        histoire = row.get("Petite histoire", "")
        piste = row.get("Piste de recherche", "")
        sources = row.get("Sources (ID Document)", "")
        notes = row.get("Notes / Hypothèses", "")
        lien_meme = row.get("Lien avec la MEME", "")

        photos = split_list(row.get("Photos (liens Postimage, séparés par ;)", ""))
        portrait = photos[0] if photos else None
        gallery_photos = photos[1:] if len(photos) > 1 else []

        pere = link_or_text(row.get("Père (ID Personne)", ""))
        mere = link_or_text(row.get("Mère (ID Personne)", ""))
        parents_bits = [p for p in (pere, mere) if p]
        parents_html = ""
        if parents_bits:
            parents_html = f"<dt>Fils / fille de</dt><dd>{' et '.join(parents_bits)}</dd>"

        kids = find_children(str(pid).strip(), personnes)
        kids_html = ""
        if kids:
            kid_links = [f'<a href="{slug(k.get("ID Personne",""))}.html">{person_display_name(k)}</a>' for k in kids]
            kids_html = f"<dt>Père / mère de</dt><dd>{', '.join(kid_links)}</dd>"

        rows_dl = []
        for label, key in [
            ("Née/né le", "Date de naissance"), ("Vers (estimée)", "Date de naissance estimée"),
            ("Lieu de naissance", "Lieu de naissance"), ("Décédée/décédé le", "Date de décès"),
            ("Lieu de décès", "Lieu de décès"),
        ]:
            v = row.get(key, "")
            if not is_blank(v):
                rows_dl.append(f"<dt>{label}</dt><dd>{v}</dd>")
        dl_html = "".join(rows_dl) + parents_html + kids_html
        if not dl_html:
            dl_html = "<dt>État civil</dt><dd class='empty'>encore inconnu</dd>"

        photo_html = (
            f'<img class="photo" src="{portrait}" alt="Portrait de {name}">' if portrait
            else '<div class="photo">Pas encore de photo</div>'
        )

        gallery_html = ""
        if gallery_photos:
            thumbs = "".join(
                f'<a href="{u}" target="_blank" rel="noopener"><img src="{u}" alt="Photo de {name}" loading="lazy"></a>'
                for u in gallery_photos
            )
            gallery_html = f"<h2>Galerie</h2><div class='gallery'>{thumbs}</div>"

        histoire_html = f"<div class='histoire'><p>{histoire}</p></div>" if not is_blank(histoire) else ""
        piste_html = f"<div class='piste'><strong>Piste de recherche :</strong> {piste}</div>" if not is_blank(piste) else ""
        notes_html = f"<p class='meta'>{notes}</p>" if not is_blank(notes) else ""
        sources_html = f"<p class='meta'>Sources : {sources}</p>" if not is_blank(sources) else ""

        body = f"""
        <a class="back" href="../index.html">&larr; Retour à l'arbre</a>
        <div class="idcard">
          {photo_html}
          <div class="infos">
            <div class="eyebrow">{'' if is_blank(lien_meme) else lien_meme}</div>
            <h2>{name} {stamp(row.get('Statut (confirmé / hypothèse)'))}</h2>
            <dl>{dl_html}</dl>
          </div>
        </div>
        {histoire_html}
        {piste_html}
        {notes_html}
        {sources_html}
        {gallery_html}
        """
        (PERSONNES_DIR / f"{slug(pid)}.html").write_text(
            page(name, body, depth=1), encoding="utf-8"
        )


def build_pistes(personnes: pd.DataFrame):
    leads = compute_leads(personnes)
    if not leads:
        body = "<p class='empty'>Aucune piste ouverte pour l'instant : soit tout est confirmé, soit l'arbre est encore vide.</p>"
    else:
        cards = []
        for lead in leads:
            link = f"personnes/{slug(lead['id'])}.html"
            cards.append(f"""
            <div class="card">
              <a class="name" href="{link}">{lead['name']}</a>
              <div class="piste"><strong>À chercher :</strong> {lead['text']}</div>
            </div>""")
        body = "<h2>Ce qu'il nous manque encore</h2>" + "".join(cards)
    (DOCS / "pistes.html").write_text(page("Pistes de recherche", body, active="pistes", depth=0), encoding="utf-8")


def build_archives_privees():
    body = f"""
    <h2>Pourquoi certains documents ne sont pas sur ce site</h2>
    <div class="archive-box">
      <p>Les livrets de famille et actes de naissance contiennent des informations
      personnelles qui peuvent concerner des membres de la famille encore vivants
      aujourd'hui (parents, enfants, mentions de mariage ou de reconnaissance).
      Les publier sur un site public ferait courir un vrai risque
      (usurpation d'identité notamment) — on a donc fait le choix de les garder
      strictement privés.</p>
      <p>Ces documents sont conservés sur un espace privé et chiffré
      (Proton Drive), accessible uniquement sur demande.</p>
    </div>
    <h2>Comment en obtenir une copie</h2>
    <p>Écrivez-nous à l'adresse ci-dessous en précisant le document qui vous
    intéresse (ou la personne concernée) : on vous envoie un lien ou un code
    d'accès personnel pour le télécharger.</p>
    """
    (DOCS / "archives-privees.html").write_text(
        page("Archives privées", body, active="archives", depth=0), encoding="utf-8"
    )


def main():
    DOCS.mkdir(exist_ok=True)
    personnes = load_data()
    build_index(personnes)
    build_person_pages(personnes)
    build_pistes(personnes)
    build_archives_privees()
    print(f"Site généré dans {DOCS} — {len(personnes)} personne(s).")


if __name__ == "__main__":
    main()