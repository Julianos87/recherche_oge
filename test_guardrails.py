import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

import dossier
import pipeline


def dossier_minimal(nom="Cabinet Test", siren="1"):
    return {
        "nom": nom,
        "nom_legal": nom,
        "dirigeant_principal": "Jean Test",
        "adresse": "1 rue du Test, 87000 LIMOGES",
        "zipcode": "87000",
        "villes": ["87000 LIMOGES"],
        "site": "",
        "site_description": "",
        "site_extrait": "",
        "specialites": "Non disponible",
        "creation": "",
        "recrutement": "Non disponible",
        "analyse_ia": "Valeurs inventees par une ancienne analyse IA.",
        "siren": siren,
    }


def lettre_valide():
    remplissage = " ".join(
        ["Je souhaite progresser par la pratique et contribuer avec rigueur aux missions du cabinet."]
        * 20
    )
    return "\n".join([
        "Julian Brouet",
        "31 RUE PIERRE ET NATHALIE MARTROU",
        "87000 LIMOGES",
        "06.20.70.01.99",
        "julian.brouet@gmail.com",
        "A l'attention de Monsieur Jean Test",
        "Cabinet Test",
        "1 rue du Test",
        "87000 LIMOGES",
        "Objet : Candidature pour un poste de collaborateur en vue du DPLG de geometre-expert",
        "Monsieur,",
        "Actuellement engage dans une reconversion vers le metier de geometre-expert, "
        "je prepare ma candidature a la session DPLG d'octobre 2026. L'Ordre ayant "
        "confirme l'eligibilite de mon parcours, je recherche un poste de collaborateur "
        "me permettant d'effectuer le stage professionnel de deux ans.",
        remplissage,
        "Je serais heureux de pouvoir echanger avec vous au sujet de ma candidature.",
        "Je vous prie d'agreer, Monsieur, l'expression de mes salutations distinguees.",
        "Julian Brouet",
    ])


class GuardrailTests(unittest.TestCase):
    def test_analyse_ia_is_not_injected_as_fact(self):
        texte = dossier.dossier_texte_pour_llm(dossier_minimal())
        self.assertNotIn("Valeurs inventees", texte)
        self.assertIn("NIVEAU DE PERSONNALISATION : MINIMAL", texte)

    def test_personalization_levels_use_verified_data(self):
        limite = dossier_minimal("Limite", "2")
        limite["creation"] = "2020"
        riche = dossier_minimal("Riche", "3")
        riche["site_description"] = "Bornage, division et topographie."
        self.assertEqual(dossier.niveau_personnalisation(limite), "LIMITE")
        self.assertEqual(dossier.niveau_personnalisation(riche), "RICHE")

    def test_generic_oge_page_is_not_used_for_personalization(self):
        donnees = dossier_minimal("RGE", "4")
        donnees["site_extrait"] = (
            "Le géomètre-expert est le garant d'un cadre vie durable. "
            "Un géomètre-expert référent par région a été désigné."
        )
        donnees["specialites"] = "bornage, copropriété, urbanisme"
        texte = dossier.dossier_texte_pour_llm(donnees)
        self.assertEqual(dossier.niveau_personnalisation(donnees), "MINIMAL")
        self.assertNotIn("Spécialités documentées", texte)
        self.assertIn("page generale de la profession", texte)

    def test_valid_letter_passes(self):
        self.assertEqual(dossier.valider_lettre_generee(
            lettre_valide(), dossier_minimal()), [])

    def test_missing_object_is_rejected(self):
        texte = lettre_valide().replace("Objet :", "Candidature :")
        erreurs = dossier.valider_lettre_generee(texte, dossier_minimal())
        self.assertTrue(any("objet absent" in erreur for erreur in erreurs))

    def test_forbidden_wording_is_rejected(self):
        texte = lettre_valide().replace(
            "progresser par la pratique",
            "valoriser ma dimension intellectuelle",
            1,
        )
        erreurs = dossier.valider_lettre_generee(texte, dossier_minimal())
        self.assertTrue(any("dimension intellectuelle" in erreur for erreur in erreurs))

    def test_autonomous_vous_is_rejected_for_minimal_data(self):
        texte = lettre_valide().replace(
            "Je serais heureux",
            "Votre cabinet propose un environnement professionnel particulier dans lequel "
            "je souhaite developper toutes mes competences et approfondir durablement ma "
            "connaissance du metier.\nJe serais heureux",
        )
        erreurs = dossier.valider_lettre_generee(texte, dossier_minimal())
        self.assertTrue(any("VOUS autonome" in erreur for erreur in erreurs))

    def test_audit_sample_is_stratified(self):
        dossiers = []
        for index in range(30):
            minimal = dossier_minimal(f"Minimal {index}", f"M{index}")
            limite = dossier_minimal(f"Limite {index}", f"L{index}")
            limite["creation"] = "2020"
            riche = dossier_minimal(f"Riche {index}", f"R{index}")
            riche["site_description"] = "Bornage et division."
            dossiers.extend([minimal, limite, riche])
        selection = dossier.selectionner_echantillon_audit(dossiers, 20)
        niveaux = [dossier.niveau_personnalisation(item) for item in selection]
        self.assertEqual(len(selection), 20)
        self.assertEqual(niveaux.count("RICHE"), 7)
        self.assertEqual(niveaux.count("LIMITE"), 7)
        self.assertEqual(niveaux.count("MINIMAL"), 6)

    def test_generation_retries_once_after_rejection(self):
        gabarit = "{profil_candidat}\n{lettre_modele}\n{dossier_cabinet}\n{lettre_existante}"
        with patch("dossier.appeler_modele", side_effect=["Texte refuse", lettre_valide()]) as modele:
            resultat = dossier.generer_lettre(
                dossier_minimal(), "profil", "modele", gabarit, "source", "local"
            )
        self.assertEqual(resultat, lettre_valide())
        self.assertEqual(modele.call_count, 2)

    def test_docx_is_a4_and_missing_object_is_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            chemin = Path(tmp) / "lettre.docx"
            pipeline.ecrire_docx(lettre_valide(), chemin)
            document = Document(chemin)
            section = document.sections[0]
            self.assertAlmostEqual(section.page_width.cm, 21.0, places=1)
            self.assertAlmostEqual(section.page_height.cm, 29.7, places=1)

            sans_objet = lettre_valide().replace("Objet :", "Candidature :")
            with self.assertRaisesRegex(ValueError, "objet absent"):
                pipeline.ecrire_docx(sans_objet, Path(tmp) / "refusee.docx")


if __name__ == "__main__":
    unittest.main()
