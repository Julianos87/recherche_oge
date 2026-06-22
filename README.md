# Recherche de cabinets de Géomètres-Experts → lettres de motivation

Automatise la **recherche de cabinets** (stage DPLG) et le **pré-remplissage des variables**
des lettres de motivation, pour ne garder à ta main que la relecture/finalisation.

## Deux modes

- **`dossier.py` (recommandé)** — produit un **dossier de recherche complet par cabinet**
  (`sortie/dossiers/<cabinet>.md`) : implantations regroupées, dirigeants, finances,
  spécialités, contenu du site, recrutement. C'est ce que lance `lancer.bat`.
- **`pipeline.py`** — version « tableur » plus simple : une ligne de variables par cabinet
  dans `sortie/cabinets_enrichis.xlsx`.

Les deux partagent les mêmes briques (sourcing OGE, API entreprises, scraping site,
France Travail, génération de lettres).

> ⚠️ **Cadence API** : l'API `recherche-entreprises.api.gouv.fr` limite à ~7 requêtes/s
> par IP. Garde `--pause 0.3` (défaut) ; une rafale plus rapide fait bannir l'IP quelques
> minutes. Les appels sont déjà protégés par un retry/backoff.

## Ce que fait le pipeline

| Étape | Source | Champs obtenus |
|-------|--------|----------------|
| 1. Sourcing | Annuaire officiel OGE (`geometre-expert.fr`) | nom du cabinet, adresse, ville, CP, email, téléphone, lien fiche |
| 2. Enrichissement | API gouv. `recherche-entreprises.api.gouv.fr` (filtre NAF 71.12A) | **dirigeant(s)**, SIREN, date de création, tranche d'effectif, CA, résultat net |
| 3. Spécialités | site **propre** du cabinet (domaine de l'email) | spécialités déclarées (bornage, 3D, photogrammétrie, BIM, VRD…) — option `--specialites` |
| 3 bis. Recrutement | API **France Travail** (Offres d'emploi v2) | offres en cours du cabinet — option `--recrutement` |
| 4. Génération | `prompt_lettre.txt` + API Anthropic | une lettre `.docx` par cabinet (option `--lettres`) |

Sourcing + enrichissement + spécialités sont **publics, gratuits et sans clé**.
France Travail et la génération de lettres nécessitent des identifiants (voir ci-dessous).

## Installation

```powershell
python -m pip install --user openpyxl python-docx anthropic
```

## Identifiants (optionnels selon les options)

```powershell
# Génération des lettres (option --lettres)
$env:ANTHROPIC_API_KEY = "sk-ant-..."

# Détection des offres (option --recrutement)
#   1. Créer un compte sur https://francetravail.io
#   2. Créer une application, souscrire à l'API "Offres d'emploi v2"
#   3. Récupérer l'identifiant client + clé secrète :
$env:FT_CLIENT_ID     = "PAR_xxx"
$env:FT_CLIENT_SECRET = "..."
```

### Notes sur les deux options ajoutées

- **`--specialites`** : on déduit le site du cabinet du **domaine de son email**
  (les adresses `wanadoo.fr`, `orange.fr`, `geometre-expert.fr`… sont ignorées car
  ce ne sont pas des sites de cabinet). On scrape la page d'accueil + quelques pages
  probables (`/prestations`, `/services`…) et on relève les spécialités déclarées.
  Sans domaine propre → « Non disponible » (aucun faux positif).
- **`--recrutement`** : recherche les offres par patronyme + département, puis ne
  conserve que celles dont l'entreprise correspond au cabinet. Sans identifiants
  France Travail, l'option est ignorée proprement.

## Utilisation — dossiers complets (recommandé)

```powershell
python dossier.py --departements 87                 # un dossier .md par cabinet
python dossier.py --region nouvelle-aquitaine
python dossier.py --tout --pause 0.35               # France entière (long)
python dossier.py --departements 87 --lettres       # + lettres (API Anthropic)
python dossier.py --departements 87 --lettres --local  # + lettres via LLM LOCAL
python dossier.py --departements 87 --recrutement   # + offres France Travail
```

### Génération des lettres en LOCAL (LM Studio / GPU)

```powershell
# 1. Lance LM Studio, charge un modele, demarre le "Local Server" (onglet Developer).
# 2. (optionnel) precise l'URL et le modele si differents des defauts :
$env:LOCAL_LLM_URL   = "http://localhost:1234/v1/chat/completions"
$env:LOCAL_LLM_MODEL = "nom-du-modele-charge"   # souvent ignore par LM Studio
# 3. Lance avec --local :
python dossier.py --departements 87 --lettres --local
```

Le backend local parle le protocole **OpenAI-compatible** : il marche aussi avec
llama.cpp (`--server`), vLLM, text-generation-webui, ou Ollama via son endpoint
`/v1/chat/completions`. Aucun paquet supplementaire requis.

#### Workflow recommandé pour le GPU local (rapide à itérer)

La collecte (API + scraping) et la génération de lettres sont **découplées**. Chaque
collecte sauvegarde un cache `sortie/dossiers.json`. Tu peux donc :

```powershell
# 1. Collecte une fois (remplit les dossiers + le cache JSON)
python dossier.py --tout --pause 0.35

# 2. Corrige les lettres existantes en LOCAL depuis le cache, SANS rappeler les API
#    Source : sortie/lettres_corrigees ; destination : sortie/lettres_corrigees_V2
python dossier.py --lettres-seules --local
python dossier.py --lettres-seules --local --max 10   # test sur 10 cabinets
python dossier.py --lettres-seules --local --region nouvelle-aquitaine
python dossier.py --lettres-seules --local --region bretagne
python dossier.py --lettres-seules --local --region pays-de-la-loire
```

Avant une correction massive, lance l'echantillon d'audit. Il selectionne 20 cabinets
repartis entre dossiers riches, limites et minimaux, puis enregistre les resultats dans
`sortie/lettres_test_audit` sans melanger ces essais avec les lettres V2 :

```powershell
python dossier.py --lettres-seules --local --audit-echantillon
```

Chaque reponse du modele est controlee avant la creation du DOCX : objet, formule d'appel,
formulation du DPLG, signature, longueur, vocabulaire interdit et niveau reel des acquis.
Une sortie refusee est automatiquement soumise une seconde fois au modele avec la liste
des erreurs. Si elle reste non conforme, aucun fichier n'est cree pour ce cabinet.

Ainsi, si une lettre ne te plait pas, tu ajustes `prompt_lettre.txt` ou le modele
LM Studio et tu relances l'étape 2 seule — la collecte n'est pas refaite.

En mode `--lettres-seules`, chaque lettre V2 est une correction conservatrice de la
lettre `.docx` correspondante dans `sortie/lettres_corrigees`. Une lettre source
introuvable est signalee puis ignoree : le programme ne fabrique pas de remplacement
generique.

Sorties : `sortie/dossiers/<cabinet>.md`, `sortie/dossiers_cabinets.xlsx`,
et `sortie/lettres_corrigees_V2/<Region>/<Departement>/lettre_<cabinet>.docx`.

Le dossier V2 est séparé des sorties précédentes : aucune lettre existante dans
`sortie/lettres` ou `sortie/lettres_corrigees` n'est écrasée.

Les lettres sont rédigées en imitant le **style de ta lettre-modèle**, lue dans
`sortie/_lettre_julian.txt` (générée à partir de ton .docx).

## Utilisation — mode tableur simple

```powershell
# Haute-Vienne (tableur seul)
python pipeline.py --departements 87

# Plusieurs départements
python pipeline.py --departements 87 19 23 24

# Région prédéfinie (voir dict REGIONS dans pipeline.py)
python pipeline.py --region nouvelle-aquitaine

# France entière (long : ~2150 cabinets, ménage l'API avec --pause)
python pipeline.py --tout --pause 0.5

# Tout activer : spécialités (site propre) + recrutement (France Travail) + lettres
python pipeline.py --departements 87 --specialites --recrutement --lettres

# Test rapide : limiter le nombre de cabinets
python pipeline.py --region occitanie --max 5
```

## Sorties

- `sortie/cabinets_enrichis.xlsx` — une ligne par cabinet, variables prêtes.
- `sortie/lettres/<cabinet>.docx` — lettres générées (avec `--lettres`).
- `sortie/_oge_annuaire.html` — cache de l'annuaire (supprimer pour rafraîchir).

## Fichiers à personnaliser

- `profil_candidat.txt` — ton profil et tes coordonnées (l'en-tête des lettres).
- `prompt_lettre.txt` — le gabarit d'instruction envoyé au modèle.

## Pistes d'extension

- **`autres_infos`** : avis Google (API Places), actualités locales.
- **CA / finances détaillées** : l'API gouv. ne donne qu'un exercice ; Pappers/INPI
  (avec clé) fournissent l'historique et la tendance.
- **Dédoublonnage** : un même GE peut diriger plusieurs sociétés ; regrouper par SIREN.

## Notes techniques

- L'annuaire OGE embarque la liste en JSON dans l'attribut `data-cabinets` de `#maps` :
  une seule requête HTTP récupère les ~2150 cabinets.
- Les bureaux secondaires (`isbureau: true`) sont écartés par défaut (sièges seuls).
- Le rapprochement avec l'API entreprises se fait par patronyme + code postal, puis
  meilleur score de similarité de nom, en ne gardant que le NAF 71.12A.
  Vérifier les cas ambigus (cabinets homonymes) dans le tableur avant envoi.
