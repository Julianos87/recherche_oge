# -*- coding: utf-8 -*-
"""
Génère un DOSSIER DE RECHERCHE COMPLET par cabinet de Géomètre-Expert, pour
faciliter la personnalisation de la lettre de motivation.

Pour chaque cabinet (regroupé par email = siège + toutes ses implantations) :
  - identité : nom, forme juridique, SIREN/SIRET, date de création ;
  - dirigeant(s) avec leur qualité ;
  - implantations (toutes les villes du cabinet) ;
  - finances : CA et résultat net par exercice disponible ;
  - effectif (tranche) ;
  - site internet (déduit du domaine email) : titre, description, spécialités
    détectées et un extrait de contenu ;
  - (option) offres de recrutement via France Travail.

Sorties :
  - sortie/dossiers/<cabinet>.md   : dossier lisible, un par cabinet
  - sortie/dossiers_cabinets.xlsx  : tableau récapitulatif
  - sortie/lettres/<cabinet>.docx  : lettre personnalisée (option --lettres)

Usage :
  python dossier.py --departements 87
  python dossier.py --region nouvelle-aquitaine --lettres
  python dossier.py --tout --pause 0.3
"""

import argparse
import re
import time
import json
import urllib.parse
import urllib.request
from collections import OrderedDict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

import pipeline as P  # réutilise les briques déjà validées

SORTIE = P.SORTIE
DOSSIERS = SORTIE / "dossiers"
LETTRES = SORTIE / "lettres"

# Pages internes à explorer sur le site du cabinet
PAGES = ["", "/nos-prestations", "/prestations", "/nos-services", "/services",
         "/competences", "/savoir-faire", "/le-cabinet", "/qui-sommes-nous",
         "/a-propos", "/notre-cabinet"]


# --------------------------------------------------------------------------- #
# Regroupement en cabinets (siège + implantations)
# --------------------------------------------------------------------------- #
def nettoyer_nom(label):
    """Retire le suffixe d'implantation type ' - BS (VILLE)' / ' - PE (VILLE)'."""
    return re.sub(r"\s*-\s*[A-Z]{2}\s*\(.*?\)\s*$", "", label).strip()


def slugify(nom):
    return re.sub(r"[^a-z0-9]+", "-", P.sans_accent(nom).lower()).strip("-")[:60]


def grouper(cabinets):
    groupes = OrderedDict()
    for c in cabinets:
        cle = (c.get("email") or "").lower().strip() or c["link"]
        groupes.setdefault(cle, []).append(c)
    cabs = []
    for cle, membres in groupes.items():
        siege = next((m for m in membres if not m.get("isbureau")), membres[0])
        villes = []
        for m in membres:
            v = m.get("fullCity", "").strip()
            if v and v not in villes:
                villes.append(v)
        cabs.append({"siege": siege, "membres": membres,
                     "nom": nettoyer_nom(siege["label"]), "villes": villes})
    return cabs


# --------------------------------------------------------------------------- #
# Enrichissement approfondi (API Recherche d'entreprises)
# --------------------------------------------------------------------------- #
def enrichir_complet(siege, pause=0.3):
    cp = str(siege.get("zipcode", "")).strip()
    terme = P.patronyme_recherche(siege["label"])
    base = {"siren": "", "siret": "", "naf": "", "forme": "", "creation": "",
            "effectif": "", "dirigeants": [], "finances": {},
            "nb_etablissements": "", "nom_legal": ""}
    try:
        params = {"q": terme, "per_page": 10}
        if len(cp) == 5:
            params["code_postal"] = cp
        rep = P.api_entreprises(params)
    except Exception:
        return base
    finally:
        time.sleep(pause)

    cands = [e for e in rep.get("results", [])
             if (e.get("activite_principale") or "").startswith("71.12")]
    if not cands:
        return base
    e = max(cands, key=lambda x: P.similarite(x.get("nom_complet", ""), siege["label"]))

    base["nom_legal"] = e.get("nom_complet", "")
    base["siren"] = e.get("siren", "")
    base["siret"] = (e.get("siege") or {}).get("siret", "")
    base["naf"] = e.get("activite_principale", "")
    base["forme"] = e.get("nature_juridique", "")
    base["creation"] = (e.get("date_creation") or "")[:4]
    base["effectif"] = P.libelle_effectif(e.get("tranche_effectif_salarie"))
    base["nb_etablissements"] = e.get("nombre_etablissements", "")
    for d in e.get("dirigeants", []):
        nom = d.get("denomination") or " ".join(
            x for x in [(d.get("prenoms") or "").title(),
                        (d.get("nom") or "").title()] if x).strip()
        if nom:
            base["dirigeants"].append({"nom": nom, "qualite": d.get("qualite", "")})
    base["finances"] = e.get("finances") or {}
    return base


def noms_dirigeants(ent):
    return " ; ".join(d["nom"] for d in ent.get("dirigeants", [])[:4]) or "Non disponible"


# --------------------------------------------------------------------------- #
# Scraping approfondi du site propre du cabinet
# --------------------------------------------------------------------------- #
def nettoyer_html(h):
    import html as _html
    h = re.sub(r"(?is)<(script|style|noscript).*?</\1>", " ", h)
    h = re.sub(r"(?s)<[^>]+>", " ", h)
    h = _html.unescape(h)
    h = re.sub(r"\s+", " ", h)
    return h.strip()


def scraper_site(cabinet_siege, pages_max=6):
    site = P.decouvrir_site(cabinet_siege)
    res = {"url": site or "", "titre": "", "description": "",
           "specialites": "Non disponible", "extrait": ""}
    if not site:
        return res
    textes, brut_total = [], ""
    for chemin in PAGES[:pages_max]:
        try:
            h = P.http_get(site + chemin, timeout=12).decode("utf-8", "replace")
        except Exception:
            continue
        brut_total += " " + h
        if not res["titre"]:
            m = re.search(r"(?is)<title>(.*?)</title>", h)
            if m:
                res["titre"] = nettoyer_html(m.group(1))[:160]
        if not res["description"]:
            m = re.search(r'(?is)<meta[^>]+name=["\']description["\'][^>]+'
                          r'content=["\'](.*?)["\']', h)
            if m:
                res["description"] = nettoyer_html(m.group(1))[:300]
        textes.append(nettoyer_html(h))
    if not brut_total.strip():
        return res
    bas = P.sans_accent(brut_total).lower()
    found = [m for m in P.MOTS_SPECIALITES if P.sans_accent(m).lower() in bas]
    if found:
        res["specialites"] = ", ".join(found)
    # extrait : le plus long bloc de texte récolté (souvent la page d'accueil)
    res["extrait"] = (max(textes, key=len)[:700] + "…") if textes else ""
    return res


# --------------------------------------------------------------------------- #
# Construction et rendu du dossier
# --------------------------------------------------------------------------- #
def construire(cab, ent, web, recrutement):
    s = cab["siege"]
    return {
        "nom": cab["nom"],
        "nom_legal": ent.get("nom_legal", ""),
        "dirigeants": ent.get("dirigeants", []),
        "dirigeant_principal": (ent["dirigeants"][0]["nom"]
                                if ent.get("dirigeants") else "Non disponible"),
        "adresse": ", ".join(x for x in [s.get("fullAddress", "").strip(),
                                         s.get("fullCity", "").strip()] if x),
        "villes": cab["villes"],
        "email": s.get("email", ""), "tel": s.get("phone", ""),
        "site": web.get("url", ""), "site_titre": web.get("titre", ""),
        "site_description": web.get("description", ""),
        "site_extrait": web.get("extrait", ""),
        "specialites": web.get("specialites", "Non disponible"),
        "forme": ent.get("forme", ""), "siren": ent.get("siren", ""),
        "siret": ent.get("siret", ""), "creation": ent.get("creation", ""),
        "effectif": ent.get("effectif", ""),
        "nb_etablissements": ent.get("nb_etablissements", ""),
        "finances": ent.get("finances", {}),
        "recrutement": recrutement,
        "fiche_oge": s.get("link", ""),
    }


def finances_lignes(fin):
    out = []
    for annee in sorted(fin.keys()):
        d = fin[annee]
        ca = d.get("ca")
        rn = d.get("resultat_net")
        bout = f"{annee} : "
        bout += (f"CA {int(ca):,} €".replace(",", " ") if ca else "CA n.c.")
        if rn is not None:
            bout += f", résultat net {int(rn):,} €".replace(",", " ")
        out.append(bout)
    return out


def rendre_md(d):
    L = [f"# {d['nom']}", ""]
    if d["nom_legal"] and d["nom_legal"].upper() != d["nom"].upper():
        L.append(f"*Raison sociale : {d['nom_legal']}*\n")
    L += ["## Contact", f"- Adresse (siège) : {d['adresse'] or 'Non disponible'}"]
    if len(d["villes"]) > 1:
        L.append(f"- Implantations ({len(d['villes'])}) : " + " · ".join(d["villes"]))
    L.append(f"- Téléphone : {d['tel'] or 'Non disponible'}")
    L.append(f"- Email : {d['email'] or 'Non disponible'}")
    L.append(f"- Site internet : {d['site'] or 'Non disponible'}")
    L.append(f"- Fiche OGE : {d['fiche_oge']}")

    L += ["", "## Dirigeant(s) (Géomètre(s)-Expert(s))"]
    if d["dirigeants"]:
        for dd in d["dirigeants"]:
            q = f" — {dd['qualite']}" if dd.get("qualite") else ""
            L.append(f"- {dd['nom']}{q}")
    else:
        L.append("- Non disponible")

    L += ["", "## Identité & structure",
          f"- Forme juridique : {d['forme'] or 'Non disponible'}",
          f"- SIREN : {d['siren'] or 'Non disponible'}"
          f"  |  SIRET (siège) : {d['siret'] or 'Non disponible'}",
          f"- Création : {d['creation'] or 'Non disponible'}",
          f"- Effectif : {d['effectif'] or 'Non disponible'}",
          f"- Nombre d'établissements : {d['nb_etablissements'] or 'Non disponible'}"]

    L += ["", "## Finances (source : API entreprises / comptes publiés)"]
    fl = finances_lignes(d["finances"])
    L += [f"- {x}" for x in fl] if fl else ["- Non disponible (comptes non publiés)"]

    L += ["", "## Spécialités & activité",
          f"- Spécialités détectées : {d['specialites']}"]
    if d["site_titre"]:
        L.append(f"- Titre du site : {d['site_titre']}")
    if d["site_description"]:
        L.append(f"- Description : {d['site_description']}")

    L += ["", "## Recrutement", f"- {d['recrutement']}"]

    if d["site_extrait"]:
        L += ["", "## Extrait du site (brut)", "> " + d["site_extrait"].replace("\n", " ")]
    return "\n".join(L) + "\n"


def dossier_texte_pour_llm(d):
    """Version condensée injectée dans le prompt de génération de lettre."""
    parts = [f"Nom : {d['nom']}",
             f"Dirigeant(s) : {d['dirigeant_principal']}",
             f"Adresse siège : {d['adresse'] or 'Non disponible'}"]
    if len(d["villes"]) > 1:
        parts.append("Implantations : " + ", ".join(d["villes"]))
    bilan = []
    if d["creation"]:
        bilan.append(f"créé en {d['creation']}")
    if d["forme"]:
        bilan.append(d["forme"])
    if d["effectif"]:
        bilan.append(d["effectif"])
    fl = finances_lignes(d["finances"])
    if fl:
        bilan.append("finances — " + " ; ".join(fl))
    if bilan:
        parts.append("Structure : " + ", ".join(bilan))
    parts.append(f"Spécialités : {d['specialites']}")
    if d["site_description"]:
        parts.append(f"Présentation (site) : {d['site_description']}")
    parts.append(f"Recrutement : {d['recrutement']}")
    return "\n".join(parts)


# --------------------------------------------------------------------------- #
# Génération de lettre (modèle de style + dossier)
# --------------------------------------------------------------------------- #
def appel_llm_local(prompt):
    """Génère via un serveur LOCAL compatible OpenAI (LM Studio, llama.cpp, vLLM…).

    Configuration par variables d'environnement :
      LOCAL_LLM_URL   (défaut http://localhost:1234/v1/chat/completions  — LM Studio)
      LOCAL_LLM_MODEL (défaut "local-model" ; mets l'id du modèle chargé si besoin)
    """
    import os
    url = os.environ.get("LOCAL_LLM_URL", "http://localhost:1234/v1/chat/completions")
    modele = os.environ.get("LOCAL_LLM_MODEL", "local-model")
    corps = json.dumps({
        "model": modele,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 1800,
        "stream": False,
    }).encode("utf-8")
    req = urllib.request.Request(url, data=corps,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        rep = json.loads(r.read())
    return rep["choices"][0]["message"]["content"].strip()


def appel_llm_anthropic(prompt):
    import os
    try:
        import anthropic
    except ImportError:
        return None
    cle = os.environ.get("ANTHROPIC_API_KEY")
    if not cle:
        return None
    client = anthropic.Anthropic(api_key=cle)
    msg = client.messages.create(model="claude-opus-4-8", max_tokens=1800,
                                 messages=[{"role": "user", "content": prompt}])
    return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")


def generer_lettre(d, profil, lettre_modele, gabarit, backend="anthropic"):
    prompt = gabarit.format(profil_candidat=profil, lettre_modele=lettre_modele,
                            dossier_cabinet=dossier_texte_pour_llm(d))
    try:
        if backend == "local":
            return appel_llm_local(prompt)
        return appel_llm_anthropic(prompt)
    except Exception as e:
        print(f"    [lettre] échec génération ({backend}) : {e}")
        return None


# --------------------------------------------------------------------------- #
# Export récapitulatif
# --------------------------------------------------------------------------- #
COLS = [("nom", "Cabinet"), ("dirigeant_principal", "Dirigeant"),
        ("adresse", "Siège"), ("_implantations", "Implantations"),
        ("_finances", "Finances"), ("effectif", "Effectif"),
        ("specialites", "Spécialités"), ("recrutement", "Recrutement"),
        ("site", "Site"), ("email", "Email"), ("tel", "Tél"), ("siren", "SIREN")]


def exporter_xlsx(dossiers, chemin):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cabinets"
    for j, (_, t) in enumerate(COLS, 1):
        ws.cell(1, j, t).font = Font(bold=True)
    for i, d in enumerate(dossiers, 2):
        d = dict(d)
        d["_implantations"] = " · ".join(d["villes"])
        d["_finances"] = " ; ".join(finances_lignes(d["finances"])) or ""
        for j, (k, _) in enumerate(COLS, 1):
            ws.cell(i, j, d.get(k, ""))
    larg = [30, 24, 38, 30, 40, 16, 34, 30, 28, 26, 14, 12]
    for j, w in enumerate(larg, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    wb.save(chemin)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Dossiers de recherche par cabinet GE.")
    ap.add_argument("--departements", nargs="+")
    ap.add_argument("--region", choices=sorted(P.REGIONS))
    ap.add_argument("--tout", action="store_true")
    ap.add_argument("--lettres", action="store_true",
                    help="générer les lettres .docx")
    ap.add_argument("--local", action="store_true",
                    help="générer les lettres via un LLM LOCAL (LM Studio / OpenAI-compatible)")
    ap.add_argument("--lettres-seules", dest="lettres_seules", action="store_true",
                    help="ne génère QUE les lettres depuis sortie/dossiers.json "
                         "(aucun appel API/scraping ; idéal en local GPU)")
    ap.add_argument("--recrutement", action="store_true",
                    help="détecter les offres France Travail (FT_CLIENT_ID/SECRET)")
    ap.add_argument("--pause", type=float, default=0.3)
    ap.add_argument("--max", type=int, default=0)
    args = ap.parse_args()

    SORTIE.mkdir(exist_ok=True)
    DOSSIERS.mkdir(exist_ok=True)
    LETTRES.mkdir(exist_ok=True)

    racine = Path(__file__).resolve().parent
    profil = (racine / "profil_candidat.txt").read_text(encoding="utf-8")
    gabarit = (racine / "prompt_lettre.txt").read_text(encoding="utf-8")
    modele = SORTIE / "_lettre_julian.txt"
    lettre_modele = modele.read_text(encoding="utf-8") if modele.exists() else ""
    chemin_json = SORTIE / "dossiers.json"

    backend = "local" if args.local else "anthropic"

    # ----- Mode "lettres seules" : pas de réseau, on lit le cache JSON -------
    if args.lettres_seules:
        if not chemin_json.exists():
            ap.error(f"{chemin_json} introuvable : lance d'abord une collecte.")
        dossiers = json.loads(chemin_json.read_text(encoding="utf-8"))
        if args.max:
            dossiers = dossiers[:args.max]
        print(f"{len(dossiers)} dossiers -> génération des lettres ({backend})...")
        for n, d in enumerate(dossiers, 1):
            print(f"  [{n}/{len(dossiers)}] {d['nom']}")
            txt = generer_lettre(d, profil, lettre_modele, gabarit, backend)
            if txt:
                P.ecrire_docx(txt, LETTRES / f"{slugify(d['nom'])}.docx")
        print(f"\nTerminé. Lettres : {LETTRES}")
        return

    # ----- Mode collecte ----------------------------------------------------
    deps = None
    if args.region:
        deps = P.REGIONS[args.region]
    elif args.departements:
        deps = [d.zfill(2) for d in args.departements]
    elif not args.tout:
        ap.error("Précise --departements, --region, --tout ou --lettres-seules.")

    # on garde tous les membres (sièges + bureaux) pour reconstituer les implantations
    bruts = P.filtrer(P.charger_cabinets(), deps, sieges_seuls=False)
    cabs = grouper(bruts)
    if args.max:
        cabs = cabs[:args.max]
    print(f"{len(cabs)} cabinets (regroupés) à traiter.")

    token_ft = P.ft_token() if args.recrutement else None
    if args.recrutement and not token_ft:
        print("  [France Travail] identifiants absents -> recrutement ignoré.")

    dossiers = []
    for n, cab in enumerate(cabs, 1):
        print(f"  [{n}/{len(cabs)}] {cab['nom']}")
        ent = enrichir_complet(cab["siege"], pause=args.pause)
        web = scraper_site(cab["siege"])
        recr = (P.offres_recrutement(cab["siege"], token_ft)
                if token_ft else "Non disponible")
        d = construire(cab, ent, web, recr)
        dossiers.append(d)

        slug = slugify(cab["nom"])
        (DOSSIERS / f"{slug}.md").write_text(rendre_md(d), encoding="utf-8")
        # sauvegarde incrémentale du cache JSON (résiste à une interruption)
        chemin_json.write_text(json.dumps(dossiers, ensure_ascii=False, indent=1),
                               encoding="utf-8")

        if args.lettres and lettre_modele:
            txt = generer_lettre(d, profil, lettre_modele, gabarit, backend)
            if txt:
                P.ecrire_docx(txt, LETTRES / f"{slug}.docx")

    exporter_xlsx(dossiers, SORTIE / "dossiers_cabinets.xlsx")
    print(f"\nTerminé. Dossiers : {DOSSIERS}\nRécapitulatif : {SORTIE/'dossiers_cabinets.xlsx'}"
          f"\nCache : {chemin_json}")


if __name__ == "__main__":
    main()
