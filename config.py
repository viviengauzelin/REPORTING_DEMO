"""
config.py - Source de vérité unique du projet (Projet 1 : ingestion & consolidation CHR).

Ce module centralise tout ce qui décrit le modèle de données et son traitement :
- ``SETTINGS`` : paramètres techniques (chemins, encodages, tolérance de réconciliation).
- ``SOURCES`` : schéma attendu de CHAQUE fichier source (en-têtes bruts -> noms canoniques,
  types, colonnes requises). Pilote la validation ``check_columns`` source par source.
- Maps de normalisation métier (catégories, segments, statuts) : corrigent les anomalies
  de référentiel de façon centralisée et auditable.
- ``STAR_SCHEMA`` : modèle en étoile canonique que la consolidation doit produire
  (consommé ensuite par le Projet 2 : PostgreSQL + Power BI).

Philosophie : « convention over configuration ». Les défauts fonctionnent partout ;
les surcharges locales (chemins, DSN PostgreSQL) passent par variables d'environnement / `.env`.

Note de migration : ce schéma relationnel remplace l'ancien ``DATA_DICTIONARY`` plat.
Le pipeline (``utils.py``) et l'UI (``app.py``) sont adaptés en conséquence.

Usage :
    from config import SETTINGS, SOURCES, STAR_SCHEMA, CATEGORY_ALIASES
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

# ---------------------------------------------------------------------------
# Support .env (optionnel - pip install python-dotenv)
# ---------------------------------------------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv non installé -> valeurs par défaut utilisées


# ===========================================================================
# SETTINGS - Paramètres techniques de l'application
# ===========================================================================

@dataclass(frozen=True)
class AppSettings:
    """Paramètres globaux, immuables après instanciation (traçabilité d'audit).

    Attributes:
        raw_dir: Répertoire des fichiers sources « sales » (mode Batch).
        output_dir: Répertoire des livrables et logs.
        manifest_name: Nom du manifeste de vérité (oracle de réconciliation/tests).
        app_name: Nom de l'application (affiché dans les rapports).
        app_version: Version sémantique (tracée dans les logs).
        log_level: Niveau de log Python.
        hash_algorithm: Algorithme d'empreinte des fichiers sources (intégrité).
        reconciliation_tolerance_pct: Écart relatif toléré à la réconciliation
            (0.001 = 0,1 % ; couvre les arrondis IEEE 754, au-delà = alerte).
        csv_encodings_fallback: Encodages testés en lecture CSV, du plus courant
            (contexte FR) au plus rare. Justifie la lecture robuste du CRM cp1252.
        source_extensions: Extensions de sources acceptées en Batch.
    """

    # --- Chemins (surchargeables via .env) ---
    raw_dir: Path = field(default_factory=lambda: Path(os.getenv("RAW_DIR", "data_raw")))
    output_dir: Path = field(default_factory=lambda: Path(os.getenv("OUTPUT_DIR", "output")))
    manifest_name: str = "_manifest_anomalies.json"

    # --- Application ---
    app_name: str = "Reporting CHR — Projet 1 (ingestion & consolidation)"
    app_version: str = "3.2.0"
    log_level: str = field(default_factory=lambda: os.getenv("LOG_LEVEL", "INFO"))

    # --- Audit / réconciliation ---
    hash_algorithm: str = "sha256"
    reconciliation_tolerance_pct: float = field(
        default_factory=lambda: float(os.getenv("RECON_TOLERANCE_PCT", "0.001"))
    )

    # --- Lecture des sources ---
    # Ordre : utf-8-sig (Excel "CSV UTF-8" avec BOM) -> utf-8 -> cp1252/latin-1 (CRM Windows).
    csv_encodings_fallback: tuple[str, ...] = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
    source_extensions: tuple[str, ...] = ("xlsx", "csv")


SETTINGS: AppSettings = AppSettings()


# ===========================================================================
# SCHÉMA DES COLONNES - briques de description
# ===========================================================================

DtypeLiteral = Literal["datetime64[ns]", "float64", "int64", "str", "object"]


@dataclass(frozen=True)
class ColSpec:
    """Spécification d'une colonne canonique d'une source.

    Attributes:
        description: Description métier (documentation + messages d'erreur).
        dtype: Type pandas attendu après nettoyage.
        is_required: Si ``True``, l'absence de la colonne bloque le traitement.
        nullable: Si ``True``, des valeurs manquantes sont tolérées après nettoyage
            (ex: une date de première commande non saisie n'invalide pas le client).
        aliases: En-têtes bruts reconnus, en forme NORMALISÉE (strip + minuscules).
            La normalisation côté loader compare l'en-tête nettoyé à ces alias.
    """

    description: str
    dtype: DtypeLiteral
    is_required: bool
    nullable: bool = False
    aliases: tuple[str, ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class SourceSpec:
    """Description complète d'une famille de fichiers source.

    Attributes:
        description: Rôle métier de la source.
        filename_glob: Motif de nom de fichier (mode Batch).
        file_format: ``xlsx`` ou ``csv``.
        grain: Grain d'une ligne (documentation).
        columns: Schéma canonique (nom interne -> ColSpec).
        encoding: Encodage attendu (CSV uniquement ; None = auto via fallback).
        separator: Séparateur CSV (None pour xlsx).
        is_wide: ``True`` si mise en page « large » à remodeler (budget).
        sheet_to_metric: Pour les fichiers multi-feuilles, feuille -> mesure canonique.
        layout: Indices de mise en page pour le remodelage (budget large).
    """

    description: str
    filename_glob: str
    file_format: Literal["xlsx", "csv"]
    grain: str
    columns: dict[str, ColSpec]
    encoding: str | None = None
    separator: str | None = None
    is_wide: bool = False
    sheet_to_metric: dict[str, str] = field(default_factory=dict)
    layout: dict[str, object] = field(default_factory=dict)


# ===========================================================================
# SOURCES - schéma attendu de chaque fichier (pilote check_columns)
# ===========================================================================

SOURCES: dict[str, SourceSpec] = {
    # -- 1. Export ERP des ventes (24 fichiers mensuels, grain ligne) --
    "ventes": SourceSpec(
        description="Export ERP des ventes, dénormalisé au grain ligne de commande.",
        filename_glob="export_ventes_*.xlsx",
        file_format="xlsx",
        grain="ligne de commande",
        columns={
            "date_commande": ColSpec(
                "Date de la commande. Formats sources : JJ/MM/AAAA. "
                "Ligne sans date valide = supprimée (une transaction non datée "
                "n'est pas comptabilisable).",
                "datetime64[ns]", is_required=True, aliases=("date commande",),
            ),
            "commande_id": ColSpec(
                "Identifiant de commande. Avec N° Ligne, forme la clé de déduplication.",
                "str", is_required=True, aliases=("n° commande", "n commande"),
            ),
            "ligne_id": ColSpec(
                "Identifiant de ligne de commande. Clé de dédup (commande_id, ligne_id).",
                "str", is_required=True, aliases=("n° ligne", "n ligne"),
            ),
            "code_client": ColSpec(
                "Code client (jointure vers le CRM).",
                "str", is_required=True, aliases=("code client",),
            ),
            "code_produit": ColSpec(
                "Code produit (jointure vers le catalogue).",
                "str", is_required=True, aliases=("code produit",),
            ),
            "quantite": ColSpec(
                "Quantité commandée.",
                "int64", is_required=True, aliases=("quantité", "quantite"),
            ),
            "prix_unitaire_vente": ColSpec(
                "Prix unitaire de vente HT. Format FR (virgule, espace milliers) à convertir. "
                "Montant non convertible -> NaN, exclu du CA et tracé.",
                "float64", is_required=True,
                aliases=("pu vente ht", "pu vente", "prix unitaire"),
            ),
            "remise_pct": ColSpec(
                "Remise en pourcentage (ex: 5.0 = 5 %). Divisée par 100 au nettoyage.",
                "float64", is_required=True, aliases=("remise %", "remise"),
            ),
            "code_commercial": ColSpec(
                "Code commercial (jointure vers le référentiel commerciaux).",
                "str", is_required=True, aliases=("code commercial",),
            ),
            "statut": ColSpec(
                "Statut de la commande (Livrée / Retour / Annulée). Pilote l'inclusion au CA.",
                "str", is_required=True, aliases=("statut", "status"),
            ),
        },
    ),

    # -- 2. Référentiel client (CRM, cp1252) --
    "crm": SourceSpec(
        description="Référentiel clients issu du CRM (saisie commerciale, encodage cp1252).",
        filename_glob="export_clients_crm.csv",
        file_format="csv",
        grain="client",
        encoding="cp1252",
        separator=";",
        columns={
            "code_client": ColSpec(
                "Code client (clé primaire CRM).",
                "str", is_required=True, aliases=("code_client",),
            ),
            "raison_sociale": ColSpec(
                "Raison sociale du client.",
                "str", is_required=False, aliases=("raison_sociale",),
            ),
            "segment": ColSpec(
                "Segment commercial. Normalisé (casse/espaces) vers le référentiel canonique.",
                "str", is_required=True, aliases=("segment",),
            ),
            "type_etablissement": ColSpec(
                "Type d'établissement (Restaurant / Hôtel / Bar / Collectivité).",
                "str", is_required=False, aliases=("type_etablissement",),
            ),
            "code_region": ColSpec(
                "Code région (jointure vers le référentiel régions).",
                "str", is_required=True, aliases=("code_region",),
            ),
            "ville": ColSpec(
                "Ville du client.",
                "str", is_required=False, aliases=("ville",),
            ),
            "code_commercial": ColSpec(
                "Code du commercial en charge du compte.",
                "str", is_required=False, aliases=("code_commercial",),
            ),
            "date_premiere_commande": ColSpec(
                "Date de première commande (cohortes). Saisie main : valeurs invalides "
                "tolérées (-> NaT), sans bloquer le client.",
                "datetime64[ns]", is_required=True, nullable=True,
                aliases=("date_premiere_commande",),
            ),
        },
    ),

    # -- 3. Catalogue produit (maintenu à la main) --
    "catalogue": SourceSpec(
        description="Catalogue produit maintenu à la main (variantes orthographiques de catégorie).",
        filename_glob="catalogue_produits.xlsx",
        file_format="xlsx",
        grain="produit",
        columns={
            "code_produit": ColSpec(
                "Code produit (clé primaire catalogue).",
                "str", is_required=True, aliases=("code_produit",),
            ),
            "libelle": ColSpec(
                "Libellé produit.",
                "str", is_required=False, aliases=("libelle", "libellé"),
            ),
            "categorie": ColSpec(
                "Catégorie produit. Variantes orthographiques à normaliser via "
                "CATEGORY_ALIASES (sinon l'effet mix est faussé).",
                "str", is_required=True, aliases=("categorie", "catégorie"),
            ),
            "sous_categorie": ColSpec(
                "Sous-catégorie produit.",
                "str", is_required=False, aliases=("sous_categorie", "sous_catégorie"),
            ),
            "cout_unitaire": ColSpec(
                "Coût unitaire HT (base du calcul de marge en SQL).",
                "float64", is_required=True, aliases=("cout_unitaire", "coût_unitaire"),
            ),
            "prix_catalogue": ColSpec(
                "Prix catalogue HT (prix de référence).",
                "float64", is_required=True, aliases=("prix_catalogue",),
            ),
            "fournisseur": ColSpec(
                "Fournisseur du produit.",
                "str", is_required=False, aliases=("fournisseur",),
            ),
        },
    ),

    # -- 4. Référentiel commerciaux (sales-ops, propre) --
    "commerciaux": SourceSpec(
        description="Référentiel commerciaux (export sales-ops, propre - source de contraste).",
        filename_glob="commerciaux.csv",
        file_format="csv",
        grain="commercial",
        encoding="utf-8",
        separator=";",
        columns={
            "code_commercial": ColSpec(
                "Code commercial (clé primaire).",
                "str", is_required=True, aliases=("code_commercial",),
            ),
            "nom": ColSpec(
                "Nom du commercial.",
                "str", is_required=True, aliases=("nom",),
            ),
            "code_region": ColSpec(
                "Code région de rattachement.",
                "str", is_required=True, aliases=("code_region",),
            ),
            "date_embauche": ColSpec(
                "Date d'embauche (ancienneté, trajectoires de performance).",
                "datetime64[ns]", is_required=True, aliases=("date_embauche",),
            ),
        },
    ),

    # -- 5. Budget (2 fichiers annuels, format LARGE à remodeler) --
    "budget": SourceSpec(
        description="Budget du contrôle de gestion, format large (mois en colonnes, sous-totaux).",
        filename_glob="budget_*.xlsx",
        file_format="xlsx",
        grain="région × catégorie × mois (après remodelage)",
        is_wide=True,
        sheet_to_metric={"CA": "ca_budgete", "Marge": "marge_budgetee"},
        # Indices de mise en page pour le remodelage large -> long.
        layout={
            "title_rows": 1,          # ligne de titre à ignorer
            "header_row": 3,          # ligne d'en-tête (Région, Catégorie, mois...)
            "id_columns": ("Région", "Catégorie"),
            "subtotal_marker": "Sous-total",  # lignes de sous-total à écarter
            "month_labels": ("Janv", "Févr", "Mars", "Avr", "Mai", "Juin",
                             "Juil", "Août", "Sept", "Oct", "Nov", "Déc"),
        },
        # Schéma canonique APRÈS remodelage (melt) - pour la validation post-reshape.
        columns={
            "periode": ColSpec("Période AAAA-MM.", "str", is_required=True),
            "nom_region": ColSpec("Région.", "str", is_required=True),
            "categorie": ColSpec("Catégorie produit.", "str", is_required=True),
            "ca_budgete": ColSpec("CA budgété (€).", "float64", is_required=True),
            "marge_budgetee": ColSpec("Marge budgétée (€).", "float64", is_required=True),
        },
    ),
}


def required_columns(source_name: str) -> list[str]:
    """Retourne les colonnes obligatoires d'une source (dérivé du schéma).

    Avantage : passer une colonne en ``is_required`` la propage automatiquement
    dans toute la chaîne de validation, sans toucher au code de nettoyage.

    Args:
        source_name: Clé de ``SOURCES``.

    Returns:
        Liste des noms canoniques requis.
    """
    return [n for n, c in SOURCES[source_name].columns.items() if c.is_required]


def alias_index(source_name: str) -> dict[str, str]:
    """Construit l'index {en-tête normalisé -> nom canonique} d'une source.

    Sert au loader : après normalisation d'un en-tête brut (strip + minuscules),
    on retrouve la colonne canonique correspondante.

    Args:
        source_name: Clé de ``SOURCES``.

    Returns:
        Dictionnaire alias normalisé -> nom canonique.
    """
    index: dict[str, str] = {}
    for canonical, spec in SOURCES[source_name].columns.items():
        index[canonical] = canonical  # le nom canonique se reconnaît lui-même
        for alias in spec.aliases:
            index[alias] = canonical
    return index


# ===========================================================================
# NORMALISATION MÉTIER - référentiels canoniques + maps de correction
# ===========================================================================

# Statuts de commande et règle d'inclusion au chiffre d'affaires.
# Règle d'audit : seules les commandes LIVRÉES comptent au CA. Les Annulées sont
# exclues ; les Retours sont conservés en drapeau (analyse "brut vs net" en Projet 2).
STATUS_VALUES: tuple[str, ...] = ("Livrée", "Retour", "Annulée")
STATUS_INCLUDED_IN_CA: tuple[str, ...] = ("Livrée",)

# Segments clients canoniques. Normalisation : strip + correspondance INSENSIBLE
# À LA CASSE vers cette liste (voir canonical_segment). NB : un Title Case naïf
# corromprait l'acronyme "PME" -> "Pme" ; on passe donc par un index lower -> canonique.
SEGMENTS_CANONICAL: tuple[str, ...] = ("Grands Comptes", "PME", "Indépendants")

# Catégories produit canoniques (6) et famille "équipement" (pour l'effet mix).
CATEGORIES_CANONICAL: tuple[str, ...] = (
    "Consommables", "Ingrédients secs", "Hygiène & Entretien",
    "Petit équipement", "Gros équipement", "Mobilier",
)
EQUIPMENT_CATEGORIES: frozenset[str] = frozenset(
    {"Petit équipement", "Gros équipement", "Mobilier"}
)

# Map de normalisation des catégories : variante (en minuscules) -> catégorie canonique.
# L'anomalie de catalogue la plus dangereuse : non corrigée, elle éclate une
# catégorie en plusieurs et fausse silencieusement l'analyse de mix.
CATEGORY_ALIASES: dict[str, str] = {
    "consommables": "Consommables",
    "conso.": "Consommables",
    "hygiene et entretien": "Hygiène & Entretien",
}


def canonical_category(raw_value: str) -> str:
    """Normalise une valeur de catégorie vers sa forme canonique.

    Args:
        raw_value: Valeur de catégorie issue du catalogue (potentiellement variante).

    Returns:
        Catégorie canonique si une variante est reconnue, sinon la valeur d'origine
        nettoyée (strip). Une valeur déjà canonique est renvoyée telle quelle.
    """
    cleaned = str(raw_value).strip()
    return CATEGORY_ALIASES.get(cleaned.lower(), cleaned)


# Index de normalisation des segments : forme minuscule -> forme canonique.
# Construit depuis SEGMENTS_CANONICAL : aucune liste à maintenir en double.
_SEGMENT_INDEX: dict[str, str] = {s.lower(): s for s in SEGMENTS_CANONICAL}


def canonical_segment(raw_value: str) -> str:
    """Normalise un segment client vers sa forme canonique (insensible à la casse).

    Le segment CRM est un champ libre : il arrive en casses variées (``"pme"``,
    ``"PME "``, ``"INDÉPENDANTS"``) selon la saisie commerciale. On normalise par
    strip + correspondance sur la forme minuscule, ce qui préserve correctement les
    acronymes (``"PME"``) là où un Title Case échouerait.

    Args:
        raw_value: Valeur de segment issue du CRM (potentiellement variante de casse).

    Returns:
        Le segment canonique si reconnu, sinon la valeur d'origine nettoyée (strip).
    """
    cleaned = str(raw_value).strip()
    return _SEGMENT_INDEX.get(cleaned.lower(), cleaned)


# ===========================================================================
# REGIONS - référentiel de dimension (code <-> nom)
# ===========================================================================
#
# Donnée de dimension canonique. Aucune source transactionnelle ne porte le couple
# (code, nom) : le CRM et les commerciaux n'ont que le code (R0x), le budget n'a que
# le nom. Ce mapping est donc la source de vérité qui relie les deux, et permet de
# construire dim_region et de rattacher fait_budget à un code_region.

REGIONS: dict[str, str] = {
    "R01": "Île-de-France",
    "R02": "Nord",
    "R03": "Est",
    "R04": "Sud",
    "R05": "Ouest",
}


# ===========================================================================
# STAR_SCHEMA - modèle en étoile canonique produit par la consolidation
# ===========================================================================
#
# Sortie du Projet 1, entrée du Projet 2 (PostgreSQL + Power BI).
# IMPORTANT : la marge n'est JAMAIS stockée. La table de faits ne porte que les
# composantes (quantité, prix, remise, coût) ; CA et marge se recalculent en SQL.

STAR_SCHEMA: dict[str, tuple[str, ...]] = {
    "dim_region": ("code_region", "nom_region"),
    "dim_commercial": ("code_commercial", "nom", "code_region", "date_embauche"),
    "dim_produit": ("code_produit", "libelle", "categorie", "sous_categorie",
                    "cout_unitaire", "prix_catalogue", "fournisseur"),
    "dim_client": ("code_client", "raison_sociale", "segment", "type_etablissement",
                   "code_region", "ville", "code_commercial", "date_premiere_commande"),
    "fait_commande": ("commande_id", "date_commande", "code_client",
                      "code_commercial", "statut"),
    "fait_ligne_commande": ("ligne_id", "commande_id", "code_produit", "quantite",
                            "prix_unitaire_vente", "remise_pct"),
    "fait_budget": ("periode", "code_region", "categorie",
                    "ca_budgete", "marge_budgetee"),
}