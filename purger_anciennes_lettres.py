import json
from pathlib import Path
import sys
import re

# Ajouter le répertoire de travail pour importer pipeline et dossier
sys.path.append(str(Path(__file__).resolve().parent))
import pipeline as P
from dossier import slugify

# Chemin vers le cache
fichier_cache = Path(r"C:\Users\Julian\Documents\recherche_oge\sortie\dossiers.json")
dossiers_dir = Path(r"C:\Users\Julian\Documents\recherche_oge\sortie\dossiers")
lettres_dir = Path(r"C:\Users\Julian\Documents\recherche_oge\sortie\lettres")

if not fichier_cache.exists():
    print("Le fichier dossiers.json n'existe pas.")
    exit()

# Charger le cache
with open(fichier_cache, "r", encoding="utf-8") as f:
    data = json.load(f)

total_avant = len(data)

# On vérifie qu'il y a bien au moins 267 cabinets
if total_avant < 267:
    print(f"Il n'y a que {total_avant} cabinets dans le cache. Rien à faire.")
    exit()

# Extraire les 267 premiers cabinets à purger
cabs_a_purger = data[:267]

# Supprimer leurs fichiers associés (.md et .docx)
supprimes_md = 0
supprimes_docx = 0

for cab in cabs_a_purger:
    nom = cab.get("nom", "")
    if not nom:
        continue
    slug = slugify(nom)
    cp = str(cab.get("zipcode", "")).strip().zfill(5)
    dep = cp[:2]
    if cp.startswith("20"): dep = "20"
    
    region_nom = "Autre"
    for r_name, r_deps in P.REGIONS.items():
        if dep in r_deps:
            region_nom = r_name.title().replace("-", " ")
            break
            
    # Chemins des fichiers
    chemin_md = dossiers_dir / region_nom / dep / f"{slug}.md"
    chemin_docx = lettres_dir / region_nom / dep / f"lettre_{slug}.docx"
    
    if chemin_md.exists():
        try:
            chemin_md.unlink()
            supprimes_md += 1
        except Exception as e:
            print(f"Erreur lors de la suppression de {chemin_md} : {e}")
            
    if chemin_docx.exists():
        try:
            chemin_docx.unlink()
            supprimes_docx += 1
        except Exception as e:
            print(f"Erreur lors de la suppression de {chemin_docx} : {e}")

# Supprimer les 267 premiers éléments de la liste
del data[:267]
total_apres = len(data)

# Sauvegarder le nouveau cache nettoyé
with open(fichier_cache, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

print("Nettoyage terminé avec succès !")
print(f"Nombre de cabinets dans le cache avant : {total_avant}")
print(f"Nombre de cabinets dans le cache après suppression : {total_apres}")
print(f"Fichiers MD supprimés : {supprimes_md}")
print(f"Fichiers DOCX supprimés : {supprimes_docx}")
