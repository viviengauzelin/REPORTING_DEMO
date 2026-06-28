"""Test_utils.py - Tests unitaires et de robustesse du moteur d'ingestion.

Principes (cf. standards projet) :

- **Isolation** : tout est simulé en mémoire (DataFrames, dataclasses, manifeste
  factice). Les rares tests de loaders écrivent des fichiers minuscules dans le
  ``tmp_path`` fourni par pytest — jamais de dépendance à des fichiers réels du dépôt.
- **Edge cases** : on teste la réaction du code face à un DataFrame vide, des colonnes
  manquantes et des types erronés (texte dans une colonne de prix), pas seulement les
  cas nominaux.
- **Validation du checkpoint** : un test prouve que ``reconcile_against_manifest``
  *détecte* une perte d'intégrité (CA faussé, clé étrangère orpheline).

Exécution : ``pytest Test_utils.py -v``
"""

from __future__ import annotations

import json

import numpy as np
import pandas as pd
import pytest

import utils
from config import canonical_category, canonical_segment
from utils import (
    BudgetCleaningReport,
    CatalogueCleaningReport,
    CrmCleaningReport,
    ReconciliationCheck,
    ReconciliationReport,
    SalesCleaningReport,
)


# ===========================================================================
# parse_fr_amount - montants au format français
# ===========================================================================


class TestParseFrAmount:
    """Conversion des montants FR (espace insécable, virgule décimale)."""

    def test_format_fr_avec_separateur_milliers(self) -> None:
        """Un montant ``"3\\u00a0000,52"`` doit donner ``3000.52``."""
        out = utils.parse_fr_amount(pd.Series(["3\u00a0000,52", "17,75"]))
        assert out.iloc[0] == pytest.approx(3000.52)
        assert out.iloc[1] == pytest.approx(17.75)

    def test_format_point_inchange(self) -> None:
        """Un montant déjà en décimale point est laissé intact (idempotence)."""
        out = utils.parse_fr_amount(pd.Series(["11.73", "5"]))
        assert out.iloc[0] == pytest.approx(11.73)
        assert out.iloc[1] == pytest.approx(5.0)

    def test_texte_dans_colonne_prix_devient_nan(self) -> None:
        """EDGE CASE : du texte dans une colonne de prix devient NaN, sans lever."""
        out = utils.parse_fr_amount(pd.Series(["abc", "12,5", "N/A", None]))
        assert pd.isna(out.iloc[0])
        assert out.iloc[1] == pytest.approx(12.5)
        assert pd.isna(out.iloc[2])
        assert pd.isna(out.iloc[3])

    def test_dtype_float_garanti(self) -> None:
        """Le dtype de sortie est un float nullable, même sur des entiers ronds."""
        out = utils.parse_fr_amount(pd.Series(["100022", "81177"]))
        assert out.dtype == "Float64"

    def test_serie_vide(self) -> None:
        """EDGE CASE : une série vide ne lève pas et reste vide."""
        out = utils.parse_fr_amount(pd.Series([], dtype="object"))
        assert len(out) == 0


# ===========================================================================
# parse_date - multi-format strict
# ===========================================================================


class TestParseDate:
    """Conversion de dates multi-format sans heuristique de devinette."""

    def test_format_francais(self) -> None:
        out = utils.parse_date(pd.Series(["20/01/2024"]))
        assert out.iloc[0] == pd.Timestamp("2024-01-20")

    def test_format_iso(self) -> None:
        """Le format ISO ``AAAA-MM-JJ`` (référentiel propre) est reconnu."""
        out = utils.parse_date(pd.Series(["2019-06-12"]))
        assert out.iloc[0] == pd.Timestamp("2019-06-12")

    def test_date_impossible_devient_nat(self) -> None:
        """EDGE CASE : ``32/13/2024`` échoue à tous les formats -> NaT (pas de devinette)."""
        out = utils.parse_date(pd.Series(["32/13/2024", "20/01/2024"]))
        assert pd.isna(out.iloc[0])
        assert out.iloc[1] == pd.Timestamp("2024-01-20")

    def test_format_unique_restreint(self) -> None:
        """En imposant un seul format, une date ISO doit échouer (mono-source)."""
        out = utils.parse_date(pd.Series(["2019-06-12"]), formats=("%d/%m/%Y",))
        assert pd.isna(out.iloc[0])


# ===========================================================================
# normalize_headers & check_columns - schéma piloté par config
# ===========================================================================


class TestNormalizeHeaders:
    """Normalisation des en-têtes bruts vers les noms canoniques."""

    def test_entetes_a_espaces_et_casse(self) -> None:
        """``"PU Vente HT "`` et ``"N° Commande"`` -> noms canoniques."""
        df = pd.DataFrame(
            columns=["Date Commande", "N° Commande", "N° Ligne", "Quantité ",
                     "PU Vente HT ", "Remise %", "Code Client", "Code Produit",
                     "Code Commercial", "Statut"]
        )
        out = utils.normalize_headers(df, "ventes")
        assert "prix_unitaire_vente" in out.columns
        assert "commande_id" in out.columns
        assert "quantite" in out.columns


class TestCheckColumns:
    """Validation des colonnes requises."""

    def test_colonne_requise_manquante_leve(self) -> None:
        """EDGE CASE : une colonne requise absente lève ValueError."""
        df = pd.DataFrame(columns=["commande_id", "ligne_id"])  # incomplet
        with pytest.raises(ValueError, match="manquantes"):
            utils.check_columns(df, "ventes")

    def test_dataframe_vide_leve(self) -> None:
        """EDGE CASE : un DataFrame totalement vide lève (toutes colonnes manquent)."""
        with pytest.raises(ValueError):
            utils.check_columns(pd.DataFrame(), "ventes")


# ===========================================================================
# coerce_types - typage piloté par le schéma
# ===========================================================================


class TestCoerceTypes:
    """Coercition des types selon le dtype déclaré dans config."""

    def test_types_appliques(self) -> None:
        df = pd.DataFrame({
            "code_produit": [" P1 "], "libelle": ["x"], "categorie": ["Consommables"],
            "sous_categorie": ["y"], "cout_unitaire": ["11,73"],
            "prix_catalogue": ["26.91"], "fournisseur": ["z"],
        })
        out = utils.coerce_types(df, "catalogue")
        assert out["cout_unitaire"].iloc[0] == pytest.approx(11.73)
        assert out["code_produit"].iloc[0] == "P1"  # strip appliqué

    def test_texte_dans_float_devient_nan(self) -> None:
        """EDGE CASE : texte dans une colonne float -> NaN sans crash."""
        df = pd.DataFrame({
            "code_produit": ["P1"], "libelle": ["x"], "categorie": ["Consommables"],
            "sous_categorie": ["y"], "cout_unitaire": ["erreur"],
            "prix_catalogue": ["10"], "fournisseur": ["z"],
        })
        out = utils.coerce_types(df, "catalogue")
        assert pd.isna(out["cout_unitaire"].iloc[0])

    def test_dataframe_vide_ne_crash_pas(self) -> None:
        """EDGE CASE : coerce_types sur un DataFrame vide (bonnes colonnes) ne lève pas."""
        df = pd.DataFrame(columns=list(
            __import__("config").SOURCES["catalogue"].columns.keys()
        ))
        out = utils.coerce_types(df, "catalogue")
        assert len(out) == 0


# ===========================================================================
# Normalisation métier (config) - catégories & segments
# ===========================================================================


class TestNormalisationMetier:
    """Catégories et segments : correspondance insensible à la casse."""

    @pytest.mark.parametrize("variant,attendu", [
        ("consommables", "Consommables"),
        ("Conso.", "Consommables"),
        ("hygiene et entretien", "Hygiène & Entretien"),
        ("Mobilier", "Mobilier"),  # déjà canonique
    ])
    def test_categories(self, variant: str, attendu: str) -> None:
        assert canonical_category(variant) == attendu

    @pytest.mark.parametrize("variant,attendu", [
        ("pme", "PME"),
        ("PME ", "PME"),  # espace parasite
        ("INDÉPENDANTS", "Indépendants"),
        ("grands comptes", "Grands Comptes"),
    ])
    def test_segments(self, variant: str, attendu: str) -> None:
        assert canonical_segment(variant) == attendu

    def test_segment_acronyme_non_casse(self) -> None:
        """Régression : "PME" ne doit JAMAIS devenir "Pme" (piège du Title Case)."""
        assert canonical_segment("pme") == "PME"
        assert canonical_segment("pme") != "Pme"


# ===========================================================================
# load_ventes - intégration légère (fichiers minuscules en tmp_path)
# ===========================================================================


def _make_ventes_file(path, rows: list[dict]) -> None:
    """Écrit un fichier ventes minuscule avec les en-têtes bruts réels."""
    cols = ["Date Commande", "N° Commande", "N° Ligne", "Code Client",
            "Code Produit", "Quantité ", "PU Vente HT ", "Remise %",
            "Code Commercial", "Statut"]
    pd.DataFrame(rows, columns=cols).to_excel(path, index=False)


class TestLoadVentes:
    """Pipeline ventes : dédup, parsing FR, drop dates invalides."""

    def test_dedup_et_compteurs(self, tmp_path) -> None:
        """Doublon (même clé) retiré ; date et montant invalides comptés."""
        base = {"Code Client": "CL1", "Code Produit": "P1", "Quantité ": "2",
                "Remise %": "0", "Code Commercial": "C01", "Statut": "Livrée"}
        rows = [
            {**base, "Date Commande": "20/01/2024", "N° Commande": "CMD1",
             "N° Ligne": "L1", "PU Vente HT ": "10,00"},
            # doublon exact de L1 (clé commande+ligne) -> à dédupliquer
            {**base, "Date Commande": "20/01/2024", "N° Commande": "CMD1",
             "N° Ligne": "L1", "PU Vente HT ": "10,00"},
            # date impossible -> NaT, ligne droppée
            {**base, "Date Commande": "32/13/2024", "N° Commande": "CMD1",
             "N° Ligne": "L2", "PU Vente HT ": "5,00"},
            # montant invalide -> NaN, ligne conservée
            {**base, "Date Commande": "21/01/2024", "N° Commande": "CMD2",
             "N° Ligne": "L1", "PU Vente HT ": "erreur"},
        ]
        _make_ventes_file(tmp_path / "export_ventes_2024-01.xlsx", rows)

        df, report = utils.load_ventes(tmp_path)
        assert report.rows_raw == 4
        assert report.rows_after_dedup == 3       # 1 doublon retiré
        assert report.duplicates_removed == 1
        assert report.invalid_dates == 1          # 32/13/2024
        assert report.invalid_amounts == 1        # "erreur"
        assert report.rows_kept == 2              # 3 dédupliquées - 1 date NaT

    def test_aucun_fichier_leve(self, tmp_path) -> None:
        """EDGE CASE : répertoire sans fichier ventes -> FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            utils.load_ventes(tmp_path)


# ===========================================================================
# ReconciliationReport - logique du verdict
# ===========================================================================


class TestReconciliationReport:
    """Verdict global d'intégrité."""

    def test_tout_ok(self) -> None:
        r = ReconciliationReport(checks=[
            ReconciliationCheck("a", 1, 1, ok=True),
            ReconciliationCheck("b", 2, 2, ok=True),
        ])
        assert r.integrity_ok is True

    def test_un_echec_suffit(self) -> None:
        r = ReconciliationReport(checks=[
            ReconciliationCheck("a", 1, 1, ok=True),
            ReconciliationCheck("b", 2, 3, ok=False),
        ])
        assert r.integrity_ok is False


# ===========================================================================
# reconcile_against_manifest - VALIDATION DU CHECKPOINT
# ===========================================================================


@pytest.fixture
def tiny_pipeline():
    """Construit un star schema minuscule cohérent + rapports + manifeste factice.

    CA Livrée attendu = 25.0 (CMD1 : 2×10 + 1×5 ; CMD2 Annulée exclue).
    """
    star = {
        "dim_region": pd.DataFrame({"code_region": ["R01"], "nom_region": ["Île-de-France"]}),
        "dim_commercial": pd.DataFrame({
            "code_commercial": ["C01"], "nom": ["Alice"], "code_region": ["R01"],
            "date_embauche": [pd.Timestamp("2020-01-01")],
        }),
        "dim_produit": pd.DataFrame({
            "code_produit": ["P1", "P2"], "libelle": ["a", "b"],
            "categorie": ["Consommables", "Mobilier"], "sous_categorie": ["x", "y"],
            "cout_unitaire": [5.0, 2.0], "prix_catalogue": [10.0, 5.0],
            "fournisseur": ["f", "f"],
        }),
        "dim_client": pd.DataFrame({
            "code_client": ["CL1"], "raison_sociale": ["Client 1"], "segment": ["PME"],
            "type_etablissement": ["Restaurant"], "code_region": ["R01"],
            "ville": ["Paris"], "code_commercial": ["C01"],
            "date_premiere_commande": [pd.Timestamp("2023-01-01")],
        }),
        "fait_commande": pd.DataFrame({
            "commande_id": ["CMD1", "CMD2"],
            "date_commande": [pd.Timestamp("2024-01-20"), pd.Timestamp("2024-01-21")],
            "code_client": ["CL1", "CL1"], "code_commercial": ["C01", "C01"],
            "statut": ["Livrée", "Annulée"],
        }),
        "fait_ligne_commande": pd.DataFrame({
            "ligne_id": ["L1", "L2", "L3"],
            "commande_id": ["CMD1", "CMD1", "CMD2"],
            "code_produit": ["P1", "P2", "P1"],
            "quantite": pd.array([2, 1, 10], dtype="Int64"),
            "prix_unitaire_vente": pd.array([10.0, 5.0, 100.0], dtype="Float64"),
            "remise_pct": [0.0, 0.0, 0.0],
        }),
        "fait_budget": pd.DataFrame({
            "periode": ["2024-01"], "code_region": ["R01"],
            "categorie": ["Consommables"], "ca_budgete": [1000.0],
            "marge_budgetee": [400.0],
        }),
    }
    reports = {
        "ventes": SalesCleaningReport(
            files_read=1, rows_raw=4, rows_after_dedup=3, duplicates_removed=1,
            invalid_dates=1, invalid_amounts=1, rows_kept=3),
        "catalogue": CatalogueCleaningReport(
            rows=2, n_categories_raw=2, n_categories_canonical=2,
            n_products_recategorized=0),
        "crm": CrmCleaningReport(
            rows=1, invalid_first_order_dates=0, segments_normalized=0,
            invalid_segments=0, distinct_segments=1),
        "budget": BudgetCleaningReport(
            files_read=1, rows=1, n_regions=1, n_categories=1, n_periods=1,
            subtotal_rows_dropped=0, leaked_subtotals=0, missing_amounts=0,
            duplicate_keys=0),
    }
    manifest = {"attendus_apres_nettoyage": {
        "ca_reconcilie_attendu_eur": 25.0,
        "nb_lignes_apres_dedup": 3,
        "nb_dates_invalides_attendu": 1,
        "nb_montants_invalides_attendu": 1,
        "nb_categories_apres_normalisation": 2,
    }}
    return star, reports, manifest


class TestReconcileAgainstManifest:
    """Le checkpoint doit valider un pipeline sain ET détecter toute corruption."""

    def test_pipeline_sain_passe(self, tiny_pipeline) -> None:
        star, reports, manifest = tiny_pipeline
        recon = utils.reconcile_against_manifest(star, reports, manifest)
        assert recon.integrity_ok is True

    def test_detecte_ca_fausse(self, tiny_pipeline) -> None:
        """Perte d'intégrité : on supprime une ligne livrée -> le CA chute -> détecté."""
        star, reports, manifest = tiny_pipeline
        # On retire L1 (2×10 = 20) : CA passe de 25 à 5.
        star["fait_ligne_commande"] = star["fait_ligne_commande"].iloc[1:].copy()
        recon = utils.reconcile_against_manifest(star, reports, manifest)
        assert recon.integrity_ok is False
        ca_check = next(c for c in recon.checks if c.label.startswith("CA"))
        assert ca_check.ok is False

    def test_detecte_fk_orpheline(self, tiny_pipeline) -> None:
        """Perte d'intégrité : une ligne pointe vers un produit inexistant -> détecté."""
        star, reports, manifest = tiny_pipeline
        lignes = star["fait_ligne_commande"].copy()
        lignes.loc[lignes["ligne_id"] == "L2", "code_produit"] = "P_INEXISTANT"
        star["fait_ligne_commande"] = lignes
        recon = utils.reconcile_against_manifest(star, reports, manifest)
        assert recon.integrity_ok is False
        fk_check = next(c for c in recon.checks if "Produits" in c.label)
        assert fk_check.ok is False
        assert fk_check.obtained == 1  # exactement 1 orphelin

    def test_detecte_compteur_dedup_fausse(self, tiny_pipeline) -> None:
        """Perte d'intégrité : un compteur de dédup erroné est détecté."""
        star, reports, manifest = tiny_pipeline
        manifest["attendus_apres_nettoyage"]["nb_lignes_apres_dedup"] = 999
        recon = utils.reconcile_against_manifest(star, reports, manifest)
        assert recon.integrity_ok is False


# ===========================================================================
# build_reporting_views - pont star schema -> vues de reporting
# ===========================================================================


def _bridge_star(
    commandes: list[dict],
    lignes: list[dict],
    commerciaux: list[dict],
) -> dict[str, pd.DataFrame]:
    """Construit un star schema minimal (3 tables) pour tester le pont.

    Seules ``fait_commande``, ``fait_ligne_commande`` et ``dim_commercial`` sont
    lues par ``build_reporting_views`` ; un star partiel suffit donc et isole le test.
    """
    return {
        "fait_commande": pd.DataFrame(commandes),
        "fait_ligne_commande": pd.DataFrame(lignes),
        "dim_commercial": pd.DataFrame(commerciaux),
    }


class TestBuildReportingViews:
    """Le pont doit recalculer le CA sans perte et exposer les cas KO sans masquer."""

    def test_cas_nominal_recolle_au_ca_livre(self, tiny_pipeline) -> None:
        """Le total des 3 vues égale le CA livré (25.0) ; la commande Annulée est exclue."""
        star, _, _ = tiny_pipeline
        df, months, salespeople = utils.build_reporting_views(star)

        # Colonnes attendues par workbook.py
        assert list(df.columns) == ["date", "mois", "commercial", "montant"]
        # CMD1 livrée -> 2 lignes (L1, L2) ; CMD2 Annulée (L3) exclue.
        assert len(df) == 2
        assert df["montant"].sum() == pytest.approx(25.0)
        assert months["montant"].sum() == pytest.approx(25.0)
        assert salespeople["montant"].sum() == pytest.approx(25.0)
        # Agrégats : un seul mois, un seul commercial.
        assert months.loc[0, "mois"] == "2024-01"
        assert months.loc[0, "montant"] == pytest.approx(25.0)
        assert salespeople.loc[0, "commercial"] == "Alice"

    def test_montant_avec_remise(self) -> None:
        """La remise est appliquée : 10 × 100 € × (1 - 0.20) = 800 €."""
        star = _bridge_star(
            commandes=[{"commande_id": "C1", "date_commande": pd.Timestamp("2025-03-10"),
                        "code_client": "CL1", "code_commercial": "S1", "statut": "Livrée"}],
            lignes=[{"ligne_id": "L1", "commande_id": "C1", "code_produit": "P1",
                     "quantite": 10, "prix_unitaire_vente": 100.0, "remise_pct": 0.20}],
            commerciaux=[{"code_commercial": "S1", "nom": "Bob", "code_region": "R1",
                          "date_embauche": pd.Timestamp("2020-01-01")}],
        )
        df, _, _ = utils.build_reporting_views(star)
        assert df["montant"].sum() == pytest.approx(800.0)

    def test_prix_invalide_devient_nan_mais_ligne_conservee(self) -> None:
        """Edge : prix NULL -> montant NaN, exclu de la somme, mais ligne gardée dans df."""
        star = _bridge_star(
            commandes=[{"commande_id": "C1", "date_commande": pd.Timestamp("2025-01-05"),
                        "code_client": "CL1", "code_commercial": "S1", "statut": "Livrée"}],
            lignes=[
                {"ligne_id": "L1", "commande_id": "C1", "code_produit": "P1",
                 "quantite": 2, "prix_unitaire_vente": 50.0, "remise_pct": 0.0},
                {"ligne_id": "L2", "commande_id": "C1", "code_produit": "P2",
                 "quantite": 3, "prix_unitaire_vente": None, "remise_pct": 0.0},
            ],
            commerciaux=[{"code_commercial": "S1", "nom": "Bob", "code_region": "R1",
                          "date_embauche": pd.Timestamp("2020-01-01")}],
        )
        df, months, _ = utils.build_reporting_views(star)
        assert len(df) == 2                              # la ligne KO reste présente
        assert int(df["montant"].isna().sum()) == 1      # un montant NaN
        assert df["montant"].sum() == pytest.approx(100.0)  # 2×50, la ligne KO ignorée
        assert months["montant"].sum() == pytest.approx(100.0)

    def test_date_invalide_exclue_des_mois_mais_ligne_conservee(self) -> None:
        """Edge : date NaT -> mois nul, exclu de report_months, mais ligne gardée dans df."""
        star = _bridge_star(
            commandes=[{"commande_id": "C1", "date_commande": pd.NaT,
                        "code_client": "CL1", "code_commercial": "S1", "statut": "Livrée"}],
            lignes=[{"ligne_id": "L1", "commande_id": "C1", "code_produit": "P1",
                     "quantite": 1, "prix_unitaire_vente": 42.0, "remise_pct": 0.0}],
            commerciaux=[{"code_commercial": "S1", "nom": "Bob", "code_region": "R1",
                          "date_embauche": pd.Timestamp("2020-01-01")}],
        )
        df, months, _ = utils.build_reporting_views(star)
        assert len(df) == 1                  # ligne conservée
        assert df["mois"].isna().all()       # mois nul (NaT)
        assert months.empty                  # aucun mois exploitable -> agrégat vide

    def test_commercial_orphelin_visible_jamais_perdu(self) -> None:
        """Détection : un code_commercial absent du référentiel -> commercial=None (LEFT JOIN).

        Preuve défensive : la ligne n'est PAS perdue silencieusement (le CA reste
        complet), mais l'anomalie devient visible (nom nul) au lieu de disparaître.
        """
        star = _bridge_star(
            commandes=[{"commande_id": "C1", "date_commande": pd.Timestamp("2025-02-01"),
                        "code_client": "CL1", "code_commercial": "ZZZ", "statut": "Livrée"}],
            lignes=[{"ligne_id": "L1", "commande_id": "C1", "code_produit": "P1",
                     "quantite": 1, "prix_unitaire_vente": 99.0, "remise_pct": 0.0}],
            commerciaux=[{"code_commercial": "S1", "nom": "Bob", "code_region": "R1",
                          "date_embauche": pd.Timestamp("2020-01-01")}],
        )
        df, _, salespeople = utils.build_reporting_views(star)
        assert len(df) == 1                              # ligne conservée (pas de perte)
        assert df["montant"].sum() == pytest.approx(99.0)  # CA complet préservé
        assert df["commercial"].isna().all()             # anomalie rendue visible
        assert salespeople.empty                         # nom nul -> hors agrégat commercial

    def test_statut_non_livre_exclu(self) -> None:
        """Seules les commandes livrées comptent : Retour et Annulée sont exclues."""
        star = _bridge_star(
            commandes=[
                {"commande_id": "C1", "date_commande": pd.Timestamp("2025-01-01"),
                 "code_client": "CL1", "code_commercial": "S1", "statut": "Livrée"},
                {"commande_id": "C2", "date_commande": pd.Timestamp("2025-01-02"),
                 "code_client": "CL1", "code_commercial": "S1", "statut": "Retour"},
                {"commande_id": "C3", "date_commande": pd.Timestamp("2025-01-03"),
                 "code_client": "CL1", "code_commercial": "S1", "statut": "Annulée"},
            ],
            lignes=[
                {"ligne_id": "L1", "commande_id": "C1", "code_produit": "P1",
                 "quantite": 1, "prix_unitaire_vente": 10.0, "remise_pct": 0.0},
                {"ligne_id": "L2", "commande_id": "C2", "code_produit": "P1",
                 "quantite": 1, "prix_unitaire_vente": 999.0, "remise_pct": 0.0},
                {"ligne_id": "L3", "commande_id": "C3", "code_produit": "P1",
                 "quantite": 1, "prix_unitaire_vente": 999.0, "remise_pct": 0.0},
            ],
            commerciaux=[{"code_commercial": "S1", "nom": "Bob", "code_region": "R1",
                          "date_embauche": pd.Timestamp("2020-01-01")}],
        )
        df, _, _ = utils.build_reporting_views(star)
        assert len(df) == 1                               # seule la ligne livrée
        assert df["montant"].sum() == pytest.approx(10.0)  # Retour/Annulée exclus


# ===========================================================================
# load_manifest - chargement de l'oracle
# ===========================================================================


class TestLoadManifest:
    """Le manifeste est obligatoire : on le charge, ou on échoue explicitement."""

    def test_charge_manifeste_present(self, tmp_path) -> None:
        """Un manifeste écrit sur disque est relu fidèlement."""
        from config import SETTINGS
        payload = {"attendus_apres_nettoyage": {"ca_reconcilie_attendu_eur": 1234.5}}
        (tmp_path / SETTINGS.manifest_name).write_text(
            json.dumps(payload), encoding="utf-8"
        )
        loaded = utils.load_manifest(tmp_path)
        assert loaded["attendus_apres_nettoyage"]["ca_reconcilie_attendu_eur"] == 1234.5

    def test_manifeste_absent_leve_erreur(self, tmp_path) -> None:
        """Sans oracle, l'audit n'a pas de sens : FileNotFoundError explicite."""
        with pytest.raises(FileNotFoundError):
            utils.load_manifest(tmp_path)


# ===========================================================================
# build_analytics_views - couche analytique DAF (marge, budget, YoY, axes)
# ===========================================================================


class TestBuildAnalyticsViews:
    """La couche analytique recalcule la marge et l'écart budget, sans rien stocker."""

    def test_cles_presentes(self, tiny_pipeline) -> None:
        """Les 12 vues attendues sont présentes."""
        star, _, _ = tiny_pipeline
        views = utils.build_analytics_views(star)
        assert set(views) == {
            "summary", "by_year", "monthly", "by_category",
            "budget_monthly", "budget_by_region",
            "category_mix", "by_salesperson", "by_segment",
            "client_pareto", "client_top", "by_discount_band",
        }

    def test_by_salesperson(self, tiny_pipeline) -> None:
        """Un seul commercial (Alice) : CA 25, marge 13, taux 52 %, part 100 %."""
        star, _, _ = tiny_pipeline
        sp = utils.build_analytics_views(star)["by_salesperson"]
        assert len(sp) == 1
        row = sp.iloc[0]
        assert row["commercial"] == "Alice"
        assert row["ca"] == pytest.approx(25.0)
        assert row["taux_marge"] == pytest.approx(52.0)
        assert row["part_ca_pct"] == pytest.approx(100.0)

    def test_by_segment(self, tiny_pipeline) -> None:
        """Un seul segment (PME) : 1 client, CA 25, part 100 %."""
        star, _, _ = tiny_pipeline
        seg = utils.build_analytics_views(star)["by_segment"]
        assert len(seg) == 1
        row = seg.iloc[0]
        assert row["segment"] == "PME"
        assert row["nb_clients"] == 1
        assert row["part_ca_pct"] == pytest.approx(100.0)

    def test_client_pareto_et_top(self, tiny_pipeline) -> None:
        """Un seul client : rang 1, CA cumulé 100 %, présent dans le top."""
        star, _, _ = tiny_pipeline
        views = utils.build_analytics_views(star)
        pareto = views["client_pareto"]
        assert len(pareto) == 1
        assert pareto.iloc[0]["rang"] == 1
        assert pareto.iloc[0]["ca_cumule_pct"] == pytest.approx(100.0)
        top = views["client_top"]
        assert top.iloc[0]["raison_sociale"] == "Client 1"
        assert top.iloc[0]["taux_marge"] == pytest.approx(52.0)

    def test_category_mix_part_par_annee(self, tiny_pipeline) -> None:
        """Mix 2024 : Consommables 80 %, Mobilier 20 % (une seule année dans la fixture)."""
        star, _, _ = tiny_pipeline
        mix = utils.build_analytics_views(star)["category_mix"].set_index("categorie")
        assert mix.loc["Consommables", "part_2024"] == pytest.approx(80.0)
        assert mix.loc["Mobilier", "part_2024"] == pytest.approx(20.0)

    def test_by_discount_band(self, tiny_pipeline) -> None:
        """Les deux lignes de la fixture sont sans remise : une seule tranche '0 %'."""
        star, _, _ = tiny_pipeline
        band = utils.build_analytics_views(star)["by_discount_band"]
        assert len(band) == 1
        row = band.iloc[0]
        assert row["tranche"] == "0 %"
        assert row["lignes"] == 2
        assert row["taux_marge"] == pytest.approx(52.0)
        assert row["part_ca_pct"] == pytest.approx(100.0)

    def test_summary_marge_recalculee(self, tiny_pipeline) -> None:
        """CA=25, coût=2×5+1×2=12, marge=13, taux=52 % (calcul à la main)."""
        star, _, _ = tiny_pipeline
        s = utils.build_analytics_views(star)["summary"].iloc[0]
        assert s["ca"] == pytest.approx(25.0)
        assert s["marge"] == pytest.approx(13.0)
        assert s["taux_marge"] == pytest.approx(52.0)
        assert s["clients_actifs"] == 1

    def test_coherence_ca_avec_reconciliation(self, tiny_pipeline) -> None:
        """Le CA de la vue mensuelle égale le CA recalculé par compute_ca_livree."""
        star, _, _ = tiny_pipeline
        views = utils.build_analytics_views(star)
        ca_recon = utils.compute_ca_livree(
            star["fait_ligne_commande"], star["fait_commande"]
        )
        assert views["monthly"]["ca"].sum() == pytest.approx(ca_recon)

    def test_by_category_part_et_marge(self, tiny_pipeline) -> None:
        """Consommables : ca=20, marge=10, tm=50 %, part=80 % ; Mobilier : tm=60 %."""
        star, _, _ = tiny_pipeline
        cat = utils.build_analytics_views(star)["by_category"].set_index("categorie")
        assert cat.loc["Consommables", "ca"] == pytest.approx(20.0)
        assert cat.loc["Consommables", "taux_marge"] == pytest.approx(50.0)
        assert cat.loc["Consommables", "part_ca_pct"] == pytest.approx(80.0)
        assert cat.loc["Mobilier", "taux_marge"] == pytest.approx(60.0)

    def test_ecart_budget_signe_defavorable(self, tiny_pipeline) -> None:
        """Écart = réel - budget : négatif quand on est sous le plan (réel 25 < budget 1000)."""
        star, _, _ = tiny_pipeline
        bm = utils.build_analytics_views(star)["budget_monthly"].iloc[0]
        assert bm["mois"] == "2024-01"
        assert bm["ca_reel"] == pytest.approx(25.0)
        assert bm["ca_budget"] == pytest.approx(1000.0)
        assert bm["ecart_ca"] == pytest.approx(-975.0)       # défavorable
        assert bm["ecart_marge"] == pytest.approx(13.0 - 400.0)

    def test_yoy_indefini_sur_une_seule_annee(self, tiny_pipeline) -> None:
        """Le YoY n'a pas de sens avec une seule année : renvoyé à NaN, pas à 0."""
        star, _, _ = tiny_pipeline
        s = utils.build_analytics_views(star)["summary"].iloc[0]
        assert pd.isna(s["ca_yoy_pct"])
        assert pd.isna(s["taux_marge_delta_pts"])


@pytest.fixture
def multi_axis_star():
    """Star schema multi-axes pour tester les filtres de présentation.

    Deux régions, deux commerciaux, deux segments, deux années — chaque commande
    isole un axe, ce qui rend les sommes faciles à vérifier à la main :

    - CMD1 (2024-01, Île-de-France, Alice, PME)      : P1 ×2 @10  -> CA 20, coût 10
    - CMD2 (2025-01, Sud, Bob, Grand Compte)         : P2 ×4 @5   -> CA 20, coût 8

    CA total livré = 40. Toutes commandes « Livrée », aucune remise.
    """
    star = {
        "dim_region": pd.DataFrame({
            "code_region": ["R01", "R02"],
            "nom_region": ["Île-de-France", "Sud"],
        }),
        "dim_commercial": pd.DataFrame({
            "code_commercial": ["C01", "C02"], "nom": ["Alice", "Bob"],
            "code_region": ["R01", "R02"],
            "date_embauche": [pd.Timestamp("2020-01-01"), pd.Timestamp("2021-01-01")],
        }),
        "dim_produit": pd.DataFrame({
            "code_produit": ["P1", "P2"], "libelle": ["a", "b"],
            "categorie": ["Consommables", "Mobilier"], "sous_categorie": ["x", "y"],
            "cout_unitaire": [5.0, 2.0], "prix_catalogue": [10.0, 5.0],
            "fournisseur": ["f", "f"],
        }),
        "dim_client": pd.DataFrame({
            "code_client": ["CL1", "CL2"],
            "raison_sociale": ["Client 1", "Client 2"],
            "segment": ["PME", "Grand Compte"],
            "type_etablissement": ["Restaurant", "Hôtel"],
            "code_region": ["R01", "R02"], "ville": ["Paris", "Nice"],
            "code_commercial": ["C01", "C02"],
            "date_premiere_commande": [pd.Timestamp("2023-01-01"),
                                       pd.Timestamp("2023-06-01")],
        }),
        "fait_commande": pd.DataFrame({
            "commande_id": ["CMD1", "CMD2"],
            "date_commande": [pd.Timestamp("2024-01-20"), pd.Timestamp("2025-01-15")],
            "code_client": ["CL1", "CL2"], "code_commercial": ["C01", "C02"],
            "statut": ["Livrée", "Livrée"],
        }),
        "fait_ligne_commande": pd.DataFrame({
            "ligne_id": ["L1", "L2"], "commande_id": ["CMD1", "CMD2"],
            "code_produit": ["P1", "P2"],
            "quantite": pd.array([2, 4], dtype="Int64"),
            "prix_unitaire_vente": pd.array([10.0, 5.0], dtype="Float64"),
            "remise_pct": [0.0, 0.0],
        }),
        "fait_budget": pd.DataFrame({
            "periode": ["2024-01", "2025-01"], "code_region": ["R01", "R02"],
            "categorie": ["Consommables", "Mobilier"],
            "ca_budgete": [1000.0, 500.0], "marge_budgetee": [400.0, 200.0],
        }),
    }
    return star


class TestAnalyticsViewsFilters:
    """Les filtres de présentation découpent le réel sans jamais altérer l'audit.

    Invariant central (checkpoint) : la réunion de partitions complémentaires
    redonne exactement le total non filtré — preuve que 100 % du périmètre est
    traité, simplement redécoupé pour l'affichage.
    """

    def test_sans_filtre_retrocompatible(self, multi_axis_star) -> None:
        """Sans argument = avec tous les filtres à None = CA total 40 (comportement antérieur intact)."""
        sans = utils.build_analytics_views(multi_axis_star)
        none = utils.build_analytics_views(
            multi_axis_star, months=None, regions=None,
            segments=None, salespeople=None,
        )
        assert set(sans) == set(none)
        assert sans["summary"].iloc[0]["ca"] == pytest.approx(40.0)
        assert none["summary"].iloc[0]["ca"] == pytest.approx(40.0)

    def test_filtre_periode_complementarite(self, multi_axis_star) -> None:
        """CA(2024) + CA(2025) == CA(total) et chaque vue ne contient que ses mois."""
        full = float(
            utils.build_analytics_views(multi_axis_star)["monthly"]["ca"].sum()
        )
        v24 = utils.build_analytics_views(multi_axis_star, months=["2024-01"])
        v25 = utils.build_analytics_views(multi_axis_star, months=["2025-01"])
        assert list(v24["monthly"]["mois"]) == ["2024-01"]
        assert list(v25["monthly"]["mois"]) == ["2025-01"]
        ca24 = float(v24["monthly"]["ca"].sum())
        ca25 = float(v25["monthly"]["ca"].sum())
        assert ca24 == pytest.approx(20.0)
        assert ca25 == pytest.approx(20.0)
        assert ca24 + ca25 == pytest.approx(full)  # 100 % traité, redécoupé

    def test_filtre_region_complementarite(self, multi_axis_star) -> None:
        """La somme des CA par région isolée égale le total non filtré."""
        full = float(
            utils.build_analytics_views(multi_axis_star)["monthly"]["ca"].sum()
        )
        total = 0.0
        for region in ["Île-de-France", "Sud"]:
            v = utils.build_analytics_views(multi_axis_star, regions=[region])
            total += float(v["monthly"]["ca"].sum())
        assert total == pytest.approx(full)

    def test_filtre_segment_et_commercial(self, multi_axis_star) -> None:
        """Segment PME et commercial Alice pointent tous deux la seule CMD1 (CA 20)."""
        seg = utils.build_analytics_views(multi_axis_star, segments=["PME"])
        com = utils.build_analytics_views(multi_axis_star, salespeople=["Alice"])
        assert float(seg["monthly"]["ca"].sum()) == pytest.approx(20.0)
        assert float(com["monthly"]["ca"].sum()) == pytest.approx(20.0)
        # L'unique segment / commercial restant représente 100 % du réel filtré.
        assert seg["by_segment"].iloc[0]["part_ca_pct"] == pytest.approx(100.0)
        assert com["by_salesperson"].iloc[0]["part_ca_pct"] == pytest.approx(100.0)

    def test_filtre_mono_annee_yoy_none(self, multi_axis_star) -> None:
        """Avec les deux années le YoY est défini ; restreint à une année il vaut None."""
        s_full = utils.build_analytics_views(multi_axis_star)["summary"].iloc[0]
        assert not pd.isna(s_full["ca_yoy_pct"])  # 2 années -> calculable
        s_one = utils.build_analytics_views(
            multi_axis_star, months=["2024-01"]
        )["summary"].iloc[0]
        assert pd.isna(s_one["ca_yoy_pct"])
        assert pd.isna(s_one["taux_marge_delta_pts"])

    def test_budget_suit_filtre_mois_et_region(self, multi_axis_star) -> None:
        """Le budget est filtré par mois et région (ses seules dimensions valides)."""
        bm = utils.build_analytics_views(multi_axis_star, months=["2024-01"])
        assert list(bm["budget_monthly"]["mois"]) == ["2024-01"]
        assert bm["budget_monthly"]["ca_budget"].sum() == pytest.approx(1000.0)
        br = utils.build_analytics_views(
            multi_axis_star, regions=["Île-de-France"]
        )["budget_by_region"]
        assert list(br["region"]) == ["Île-de-France"]
        assert br.iloc[0]["ca_budget"] == pytest.approx(1000.0)

    def test_filtre_vide_ne_casse_pas(self, multi_axis_star) -> None:
        """Cas limite : un filtre sans correspondance renvoie des vues vides, sans crash."""
        v = utils.build_analytics_views(multi_axis_star, months=["2099-12"])
        assert v["monthly"].empty
        s = v["summary"].iloc[0]
        assert s["ca"] == pytest.approx(0.0)
        assert pd.isna(s["taux_marge"])      # division par CA nul -> None
        assert s["clients_actifs"] == 0


# ===========================================================================
# _synthese_narrative - narratif factuel auto-généré (logique pure de workbook)
# ===========================================================================


class TestSyntheseNarrative:
    """Le narratif doit refléter fidèlement les signes (croissance, marge, budget)."""

    def test_croissance_avec_erosion_declenche_point_attention(self) -> None:
        """CA en hausse + marge en recul -> mention explicite du point d'attention."""
        import workbook
        s = pd.Series({
            "ca": 1_000_000.0, "ca_yoy_pct": 16.9, "taux_marge": 26.6,
            "taux_marge_delta_pts": -2.1, "ecart_ca_pct": -3.8, "ecart_marge_pct": -3.5,
        })
        txt = workbook._synthese_narrative(s)
        assert "hausse de 16,9 %" in txt
        assert "recul de 2,1 pts" in txt
        assert "sous le budget de 3,8 %" in txt
        assert "Point d'attention" in txt and "érosion" in txt

    def test_une_seule_annee_pas_de_yoy_ni_point_attention(self) -> None:
        """Sans YoY (une seule année), pas de phrase de croissance ni de point d'attention."""
        import workbook
        s = pd.Series({
            "ca": 1_000_000.0, "ca_yoy_pct": None, "taux_marge": 26.6,
            "taux_marge_delta_pts": None, "ecart_ca_pct": -3.8, "ecart_marge_pct": -3.5,
        })
        txt = workbook._synthese_narrative(s)
        assert "sur la période" in txt          # branche sans YoY
        assert "Point d'attention" not in txt
        assert "vs N-1" not in txt

    def test_marge_en_progression(self) -> None:
        """Une marge en hausse est décrite comme une progression, sans point d'attention."""
        import workbook
        s = pd.Series({
            "ca": 1_000_000.0, "ca_yoy_pct": 5.0, "taux_marge": 30.0,
            "taux_marge_delta_pts": 1.5, "ecart_ca_pct": 2.0, "ecart_marge_pct": 3.0,
        })
        txt = workbook._synthese_narrative(s)
        assert "progression de 1,5 pts" in txt
        assert "au-dessus du budget" in txt
        assert "Point d'attention" not in txt