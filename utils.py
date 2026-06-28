"""utils.py - Moteur d'ingestion & de consolidation (Projet 1 : distributeur B2B CHR).

Ce module porte la **logique métier pure**, indépendante de toute interface
(Streamlit ou Batch). Il est piloté par ``config.SOURCES`` : chaque source est lue,
normalisée et validée à partir de son schéma déclaratif, ce qui garantit qu'une
évolution du modèle (ajout d'une colonne requise, nouvel alias d'en-tête) se
propage sans toucher au code de nettoyage.

Conception :

- **Auditable** : empreinte SHA-256 de chaque fichier source (intégrité) et
  réconciliation qui prouve *chiffre à chiffre* que le nettoyage retombe sur
  l'oracle (``_manifest_anomalies.json``).
- **Robuste** : gestion explicite des en-têtes sales, du format FR des montants,
  des dates invalides, des doublons d'export et des encodages CSV.
- **Testable** : aucune dépendance à Streamlit ; les loaders acceptent un
  répertoire, les fonctions de bas niveau acceptent des ``pd.Series``.
- **Frontière pandas / DuckDB** : pandas nettoie *intra-source* (format physique
  des fichiers) ; DuckDB consolide *inter-sources* et recalcule le CA/la marge en
  SQL (jamais stockés).

Dépendances :
    pandas, openpyxl, duckdb  (voir requirements.txt)
    config  (ce projet)

Périmètre courant : BLOC 1 (loader des ventes + réconciliation CA). Les loaders
CRM / catalogue / commerciaux / budget et la construction complète du STAR_SCHEMA
viennent dans les blocs suivants.
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from functools import reduce
from pathlib import Path
from typing import Optional, Sequence

import duckdb
import pandas as pd

from config import (
    REGIONS,
    SEGMENTS_CANONICAL,
    SETTINGS,
    SOURCES,
    STAR_SCHEMA,
    STATUS_INCLUDED_IN_CA,
    alias_index,
    canonical_category,
    canonical_segment,
    required_columns,
)

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------


def setup_logging(base_dir: str = "output", run_ts: Optional[str] = None) -> Path:
    """Configure le logging Python vers un fichier horodaté + console.

    Le nom du fichier log inclut l'heure et la minute du run (``run_ts``) afin de
    distinguer plusieurs exécutions le même jour (relances manuelles, planificateur
    multi-horaires). ``force=True`` réinitialise les handlers existants (évite les
    doublons lors des relances en mode Batch ou en tests).

    Args:
        base_dir: Répertoire de sortie des logs (créé s'il n'existe pas).
        run_ts: Horodatage court du run au format ``HHhMM`` (ex: ``"14h30"``),
            calculé une seule fois au démarrage du run pour cohérence entre tous
            les fichiers produits. Si ``None``, calculé ici depuis ``datetime.now()``.

    Returns:
        Le chemin du fichier log créé (ex: ``output/log_2025-01-15_14h30.txt``).
    """
    Path(base_dir).mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    effective_ts = run_ts if run_ts is not None else now.strftime("%Hh%M")
    log_path = Path(base_dir) / f"log_{today}_{effective_ts}.txt"

    logging.basicConfig(
        level=SETTINGS.log_level,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )
    return log_path


# ---------------------------------------------------------------------------
# HACHAGE - Intégrité des fichiers sources (audit)
# ---------------------------------------------------------------------------


def hash_file_on_disk(path: Path) -> str:
    """Calcule l'empreinte SHA-256 d'un fichier sur le disque.

    Le hash garantit qu'un fichier n'a pas été altéré entre sa réception et son
    traitement : toute modification (même un octet) change l'empreinte. La lecture
    se fait par blocs de 64 Ko pour ne jamais charger tout le fichier en RAM
    (indispensable à la scalabilité). Le hash porte sur les octets bruts, donc
    fonctionne à l'identique pour les ``.xlsx`` et les ``.csv``.

    Args:
        path: Chemin du fichier à hacher.

    Returns:
        L'empreinte hexadécimale SHA-256 (64 caractères).

    Raises:
        FileNotFoundError: Si le fichier n'existe pas.
        PermissionError: Si l'accès au fichier est refusé.
    """
    h = hashlib.new(SETTINGS.hash_algorithm)
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65_536), b""):
            h.update(block)
    return h.hexdigest()


def hash_bytes(data: bytes) -> str:
    """Calcule l'empreinte SHA-256 d'un contenu binaire en mémoire.

    Utilisé pour les fichiers uploadés via Streamlit (déjà chargés en RAM).

    Args:
        data: Contenu binaire du fichier.

    Returns:
        L'empreinte hexadécimale SHA-256 (64 caractères).
    """
    return hashlib.new(SETTINGS.hash_algorithm, data).hexdigest()


# ---------------------------------------------------------------------------
# GESTION DES COUPURES - Nommage atomique des fichiers de sortie
# ---------------------------------------------------------------------------

#: Préfixe apposé sur les fichiers en cours d'écriture. Sa présence sur le disque
#: signale qu'une coupure est survenue pendant l'écriture : le fichier est
#: potentiellement incomplet et ne doit pas être considéré comme fiable.
_CORRUPT_PREFIX: str = "CORROMPU_"


def build_output_paths(
    base_stem: str,
    out_dir: Path,
    run_ts: str,
    extension: str,
) -> tuple[Path, Path]:
    """Calcule les chemins temporaire et définitif d'un fichier de sortie.

    Stratégie anti-coupure : un fichier est d'abord écrit sous un nom portant le
    préfixe ``CORROMPU_``. Si le run est interrompu, le fichier incomplet reste sur
    le disque avec ce préfixe et ne peut pas être confondu avec un livrable finalisé.
    Le renommage vers le nom définitif (``commit_output_file``) n'intervient qu'une
    fois toutes les écritures terminées avec succès.

    Args:
        base_stem: Racine du nom sans extension ni horodatage
            (ex: ``"reporting_2024_to_2025"``).
        out_dir: Répertoire de destination.
        run_ts: Horodatage court ``HHhMM`` (ex: ``"14h30"``), calculé une seule fois
            au démarrage du run pour cohérence entre tous les fichiers produits.
        extension: Extension avec le point (ex: ``".xlsx"``).

    Returns:
        Un tuple ``(temp_path, final_path)`` : le chemin temporaire (préfixé
        ``CORROMPU_``) et le chemin définitif (sans préfixe).
    """
    final_name = f"{base_stem}_{run_ts}{extension}"
    temp_name = f"{_CORRUPT_PREFIX}{final_name}"
    return out_dir / temp_name, out_dir / final_name


def commit_output_file(temp_path: Path, final_path: Path) -> None:
    """Renomme un fichier temporaire en son nom définitif (commit atomique).

    ``Path.rename()`` est atomique sur un même système de fichiers (un seul appel
    ``rename(2)``) : le fichier bascule instantanément du nom temporaire au nom
    définitif, sans état intermédiaire visible.

    Args:
        temp_path: Chemin du fichier temporaire (préfixe ``CORROMPU_``).
        final_path: Chemin de destination définitif.

    Raises:
        FileNotFoundError: Si ``temp_path`` n'existe pas (écriture avortée).
        OSError: Si le renommage échoue (permissions, disque différent, etc.).
    """
    if not temp_path.exists():
        logging.error(
            "[Commit] Fichier temporaire introuvable : %s - "
            "le run précédent a peut-être été interrompu avant l'écriture.",
            temp_path,
        )
        raise FileNotFoundError(
            f"Impossible de finaliser le fichier : {temp_path} introuvable."
        )
    temp_path.rename(final_path)
    logging.info("[Commit] Fichier finalisé : %s -> %s", temp_path.name, final_path.name)


# ---------------------------------------------------------------------------
# NORMALISATION D'EN-TÊTES & VALIDATION DE SCHÉMA (piloté par SOURCES)
# ---------------------------------------------------------------------------


def _coalesce_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Fusionne les colonnes de même nom (peut apparaître après mapping d'alias).

    Si deux en-têtes bruts pointent vers le même nom canonique, on conserve la
    première valeur non nulle par ligne (``bfill`` horizontal). Cas rare, mais on
    le traite pour ne jamais perdre silencieusement une colonne.

    Args:
        df: DataFrame potentiellement porteur de colonnes dupliquées.

    Returns:
        Un DataFrame sans colonnes dupliquées.
    """
    if not df.columns.duplicated().any():
        return df
    out = df.copy()
    for dup in out.columns[out.columns.duplicated()].unique():
        block = out.loc[:, out.columns == dup]
        out = out.drop(columns=dup)
        out[dup] = block.bfill(axis=1).iloc[:, 0]
    return out


def normalize_headers(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Renomme les en-têtes bruts vers les noms canoniques de la source.

    Chaque en-tête est normalisé (``strip`` + minuscules) puis confronté à
    ``alias_index(source_name)``. Un en-tête inconnu est conservé sous sa forme
    normalisée (il sera ignoré en aval s'il n'appartient pas au schéma). C'est ce
    mécanisme qui absorbe l'anomalie « en-têtes à espaces » des exports ERP
    (ex: ``"PU Vente HT "`` -> ``prix_unitaire_vente``).

    Args:
        df: DataFrame brut issu de la lecture du fichier.
        source_name: Clé de ``config.SOURCES`` (ex: ``"ventes"``).

    Returns:
        Un DataFrame dont les colonnes portent les noms canoniques.
    """
    index = alias_index(source_name)
    out = df.copy()
    out.columns = [index.get(str(c).strip().lower(), str(c).strip().lower()) for c in df.columns]
    return _coalesce_duplicate_columns(out)


def check_columns(df: pd.DataFrame, source_name: str) -> None:
    """Vérifie la présence des colonnes requises d'une source ; lève si manquantes.

    S'appuie sur ``required_columns(source_name)`` : passer une colonne en
    ``is_required`` dans ``config.SOURCES`` suffit à la rendre bloquante ici, sans
    modifier cette fonction.

    Args:
        df: DataFrame **déjà passé par** ``normalize_headers``.
        source_name: Clé de ``config.SOURCES``.

    Raises:
        ValueError: Si au moins une colonne requise est absente. Le message liste
            les colonnes manquantes, attendues et présentes (diagnostic immédiat).
    """
    required = required_columns(source_name)
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Source '{source_name}' : colonnes requises manquantes {missing}. "
            f"Attendues : {required}. Présentes : {list(df.columns)}."
        )


# ---------------------------------------------------------------------------
# PARSING TYPÉ - helpers vectorisés réutilisables
# ---------------------------------------------------------------------------


def parse_fr_amount(series: pd.Series) -> pd.Series:
    """Convertit des montants au format français en float64.

    Format source : séparateur de milliers en espace insécable (``\\u00a0``) ou
    espace, décimale en virgule (ex: ``"3\\u00a0000,52"`` -> ``3000.52``). Une valeur
    non convertible (texte, cellule vide) devient ``NaN`` (``errors="coerce"``) :
    elle est tracée comme « montant invalide » et exclue du CA, sans bloquer le run.

    Vectorisé : O(n), sans boucle Python.

    Args:
        series: Série de montants bruts (typiquement en ``str``).

    Returns:
        Une série ``float64`` ; les valeurs non convertibles sont ``NaN``.
    """
    cleaned = (
        series.astype("string")
        .str.replace("\u00a0", "", regex=False)  # espace insécable (milliers)
        .str.replace(" ", "", regex=False)  # espace classique éventuel
        .str.replace(",", ".", regex=False)  # décimale FR -> décimale point
    )
    # astype("Float64") : garantit un float nullable même quand les valeurs sont
    # des entiers ronds (sinon to_numeric infère Int64 et trahit le dtype déclaré).
    return pd.to_numeric(cleaned, errors="coerce").astype("Float64")


def parse_date(
    series: pd.Series,
    formats: tuple[str, ...] = ("%d/%m/%Y", "%Y-%m-%d"),
) -> pd.Series:
    """Convertit des dates en datetime64 en essayant plusieurs formats **stricts**.

    Le jeu de données mêle deux formats : français ``JJ/MM/AAAA`` (ERP, CRM) et ISO
    ``AAAA-MM-JJ`` (référentiel sales-ops, bien tenu). On essaie les formats dans
    l'ordre, chacun en mode strict (``errors="coerce"``), sur les seules valeurs
    encore non résolues. Aucune heuristique de devinette (``dayfirst``) : une date
    impossible (ex: ``"32/13/2024"``) échoue à tous les formats et devient ``NaT``,
    au lieu d'être réinterprétée silencieusement.

    Comme les séparateurs diffèrent (``/`` vs ``-``), il n'y a pas d'ambiguïté de
    correspondance croisée entre les deux formats.

    Args:
        series: Série de dates brutes (typiquement en ``str``).
        formats: Formats ``strftime`` essayés dans l'ordre. Restreindre à un seul
            format (ex: ``("%d/%m/%Y",)``) impose une source mono-format.

    Returns:
        Une série ``datetime64`` ; les dates qu'aucun format ne reconnaît sont ``NaT``.
    """
    s = series.astype("string")
    out = pd.to_datetime(pd.Series(pd.NA, index=s.index), errors="coerce")
    for fmt in formats:
        unresolved = out.isna()
        if not unresolved.any():
            break
        out.loc[unresolved] = pd.to_datetime(
            s[unresolved], format=fmt, errors="coerce"
        )
    return out


def coerce_types(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Applique le type cible de chaque colonne du schéma (piloté par ``config``).

    La conversion est dérivée du ``dtype`` déclaré dans ``ColSpec`` : c'est le
    schéma qui décide, pas le code. Une colonne du schéma absente du DataFrame est
    ignorée (le caractère bloquant des colonnes requises est déjà tranché par
    ``check_columns``). Les colonnes hors schéma sont laissées telles quelles.

    Règles :
        - ``datetime64[ns]`` -> ``parse_date`` (multi-format strict ; ``NaT`` si invalide) ;
        - ``float64``        -> ``parse_fr_amount`` (robuste FR et point ; ``NaN`` si KO) ;
        - ``int64``          -> entier nullable ``Int64`` (``<NA>`` si non convertible) ;
        - sinon (``str``/``object``) -> chaîne nettoyée (``strip``).

    Ce helper sert les sources « dimension » (catalogue, CRM, commerciaux, budget).
    La source de faits « ventes » reste traitée à la main dans ``load_ventes`` :
    ses compteurs d'anomalies doivent être relevés à des instants précis du pipeline
    qu'une coercition générique ne saurait exprimer.

    Args:
        df: DataFrame **déjà passé par** ``normalize_headers``.
        source_name: Clé de ``config.SOURCES``.

    Returns:
        Un DataFrame typé selon le schéma de la source.
    """
    spec = SOURCES[source_name]
    out = df.copy()
    for name, col in spec.columns.items():
        if name not in out.columns:
            continue
        if col.dtype == "datetime64[ns]":
            out[name] = parse_date(out[name])
        elif col.dtype == "float64":
            out[name] = parse_fr_amount(out[name])
        elif col.dtype == "int64":
            out[name] = pd.to_numeric(out[name], errors="coerce").astype("Int64")
        else:
            out[name] = out[name].astype("string").str.strip()
    return out


# ---------------------------------------------------------------------------
# LOADER VENTES (BLOC 1)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SalesCleaningReport:
    """Compteurs de nettoyage de la source « ventes » (preuve d'audit).

    Ces compteurs sont confrontés à l'oracle ``_manifest_anomalies.json`` par
    ``reconcile_against_manifest``. L'ordre des opérations est crucial : la dédup précède le
    filtrage des dates, et les montants invalides sont comptés sur l'ensemble
    dédupliqué (avant tout drop de ligne).

    Attributes:
        files_read: Nombre de fichiers ``export_ventes_*.xlsx`` lus.
        rows_raw: Lignes après concaténation brute (avant dédup).
        rows_after_dedup: Lignes après dédup ``(commande_id, ligne_id)`` —
            checkpoint oracle ``nb_lignes_apres_dedup``.
        duplicates_removed: Doublons d'export retirés (``rows_raw - rows_after_dedup``).
        invalid_dates: Lignes à date non parseable (``NaT``) — checkpoint oracle.
        invalid_amounts: Lignes à montant non convertible (``NaN``) — checkpoint oracle.
        rows_kept: Lignes conservées après drop des dates ``NaT`` (grain de
            ``fait_ligne_commande``). Les lignes à montant ``NaN`` sont conservées.
    """

    files_read: int
    rows_raw: int
    rows_after_dedup: int
    duplicates_removed: int
    invalid_dates: int
    invalid_amounts: int
    rows_kept: int


def load_ventes(raw_dir: Path) -> tuple[pd.DataFrame, SalesCleaningReport]:
    """Lit, normalise, déduplique et type la source « ventes ».

    Pipeline (ordre aligné sur l'oracle) :

    1. Glob ``export_ventes_*.xlsx`` (24 fichiers mensuels), hash de chaque fichier
       (trace d'audit), lecture en ``dtype=str`` (aucun auto-parsing : on contrôle
       chaque conversion et on préserve les zéros initiaux des codes).
    2. ``normalize_headers`` (absorbe les en-têtes à espaces) puis ``check_columns``.
    3. **Dédup** sur ``(commande_id, ligne_id)`` — *avant* tout filtrage.
    4. Parsing : prix FR -> float (``NaN`` si KO, compté), date -> datetime
       (``NaT`` si KO, compté), quantité -> ``Int64``, remise -> /100.
    5. Drop des lignes à date ``NaT`` (transaction non datable, hors périmètre CA).
       Les lignes à montant ``NaN`` sont **conservées** (drapeau) et naturellement
       exclues du ``SUM`` de réconciliation.

    Note de scalabilité (×100, ~6M lignes) : la lecture reste O(n) ; au-delà,
    paralléliser la lecture des fichiers et/ou enregistrer chaque fichier
    directement dans DuckDB (``read_xlsx``/Parquet) pour repousser l'agrégation
    hors de la RAM Python.

    Args:
        raw_dir: Répertoire contenant les fichiers sources.

    Returns:
        Un tuple ``(df, report)`` : le DataFrame propre au grain ligne de commande
        (colonnes canoniques typées) et le ``SalesCleaningReport`` associé.

    Raises:
        FileNotFoundError: Si aucun fichier ``export_ventes_*.xlsx`` n'est trouvé.
        ValueError: Si une colonne requise est absente (via ``check_columns``).
    """
    spec = SOURCES["ventes"]
    files = sorted(Path(raw_dir).glob(spec.filename_glob))
    if not files:
        raise FileNotFoundError(
            f"Aucun fichier '{spec.filename_glob}' dans {raw_dir}. "
            "Générer le dataset (generate_demo_data.py) ou vérifier RAW_DIR."
        )

    frames: list[pd.DataFrame] = []
    for f in files:
        digest = hash_file_on_disk(f)
        logging.info("[ventes] lu %s (sha256=%s…)", f.name, digest[:12])
        frames.append(pd.read_excel(f, dtype=str))

    raw = pd.concat(frames, ignore_index=True)
    rows_raw = len(raw)

    df = normalize_headers(raw, "ventes")
    check_columns(df, "ventes")

    # (1) Dédup d'abord : on retire les doublons d'export sur la clé naturelle.
    df = df.drop_duplicates(subset=["commande_id", "ligne_id"]).copy()
    rows_after_dedup = len(df)
    duplicates_removed = rows_raw - rows_after_dedup

    # (2) Parsing typé. Les compteurs sont relevés sur l'ensemble dédupliqué,
    #     AVANT tout drop de ligne (cohérence avec l'oracle).
    df["prix_unitaire_vente"] = parse_fr_amount(df["prix_unitaire_vente"])
    invalid_amounts = int(df["prix_unitaire_vente"].isna().sum())

    df["date_commande"] = parse_date(df["date_commande"], formats=("%d/%m/%Y",))
    invalid_dates = int(df["date_commande"].isna().sum())

    df["quantite"] = pd.to_numeric(df["quantite"], errors="coerce").astype("Int64")
    # Remise saisie en points de % (ex: 6.1) -> taux décimal (0.061) pour le calcul.
    df["remise_pct"] = pd.to_numeric(df["remise_pct"], errors="coerce") / 100.0

    # (3) Drop des dates invalides (non datable = hors CA). Montants NaN conservés.
    df = df[df["date_commande"].notna()].copy()
    rows_kept = len(df)

    report = SalesCleaningReport(
        files_read=len(files),
        rows_raw=rows_raw,
        rows_after_dedup=rows_after_dedup,
        duplicates_removed=duplicates_removed,
        invalid_dates=invalid_dates,
        invalid_amounts=invalid_amounts,
        rows_kept=rows_kept,
    )
    logging.info(
        "[ventes] %d fichiers | %d lignes brutes -> %d après dédup "
        "(%d doublons) | %d dates KO, %d montants KO | %d lignes conservées",
        report.files_read, report.rows_raw, report.rows_after_dedup,
        report.duplicates_removed, report.invalid_dates, report.invalid_amounts,
        report.rows_kept,
    )
    return df, report


# ---------------------------------------------------------------------------
# RÉCONCILIATION - Checkpoint d'intégrité contre l'oracle (DuckDB)
# ---------------------------------------------------------------------------


def compute_ca_livree(
    fait_ligne_commande: pd.DataFrame,
    fait_commande: pd.DataFrame,
) -> float:
    """Recalcule le chiffre d'affaires livré en SQL, depuis les faits du star schema.

    Le CA n'est jamais stocké : il se recalcule à partir des composantes de la ligne
    (``quantite``, ``prix_unitaire_vente``, ``remise_pct``), jointes au statut porté
    par la commande. Seuls les statuts de ``config.STATUS_INCLUDED_IN_CA``
    (``"Livrée"``) comptent. Les montants ``NaN`` (irrécupérables) sont explicitement
    exclus pour ne pas « empoisonner » la somme.

    Comme ``statut`` est constant par commande (vérifié), ce calcul au grain ligne ⋈
    commande est strictement équivalent à un filtre au grain ligne : c'est l'intérêt
    de modéliser le statut sur ``fait_commande`` sans perte.

    Args:
        fait_ligne_commande: Table de faits au grain ligne de commande.
        fait_commande: Table de faits au grain commande (porte ``statut``).

    Returns:
        Le chiffre d'affaires livré en euros.
    """
    con = duckdb.connect()
    try:
        con.register("fait_ligne", fait_ligne_commande)
        con.register("fait_cmd", fait_commande)
        placeholders = ", ".join("?" for _ in STATUS_INCLUDED_IN_CA)
        query = f"""
            SELECT SUM(l.quantite * l.prix_unitaire_vente * (1 - l.remise_pct)) AS ca
            FROM fait_ligne AS l
            JOIN fait_cmd AS c ON l.commande_id = c.commande_id
            WHERE c.statut IN ({placeholders})
              AND l.prix_unitaire_vente IS NOT NULL
              AND NOT isnan(l.prix_unitaire_vente)
        """
        result = con.execute(query, list(STATUS_INCLUDED_IN_CA)).fetchone()
    finally:
        con.close()
    return float(result[0]) if result and result[0] is not None else 0.0


@dataclass(frozen=True)
class ReconciliationCheck:
    """Résultat d'un point de contrôle élémentaire (une métrique de l'oracle).

    Attributes:
        label: Libellé métier du contrôle (affiché dans le rapport).
        obtained: Valeur obtenue par le pipeline.
        expected: Valeur attendue (oracle).
        ok: ``True`` si l'écart est dans la tolérance.
        detail: Précision optionnelle (écart, tolérance appliquée).
    """

    label: str
    obtained: float
    expected: float
    ok: bool
    detail: str = ""


@dataclass(frozen=True)
class ReconciliationReport:
    """Rapport de réconciliation : preuve chiffrée d'intégrité du pipeline.

    Attributes:
        checks: Liste des points de contrôle élémentaires.
        integrity_ok: ``True`` si **tous** les contrôles passent.
    """

    checks: list[ReconciliationCheck] = field(default_factory=list)

    @property
    def integrity_ok(self) -> bool:
        """Indique si tous les points de contrôle sont au vert."""
        return all(c.ok for c in self.checks)

    def render(self) -> str:
        """Rend le rapport en texte aligné (miroir du log / de la feuille Excel).

        Returns:
            Le rapport formaté en tableau lisible.
        """
        lines = ["Réconciliation contre l'oracle :"]
        for c in self.checks:
            status = "OK " if c.ok else "ALERTE"
            extra = f"  ({c.detail})" if c.detail else ""
            lines.append(
                f"  [{status}] {c.label:<26} obtenu={c.obtained:<16} "
                f"attendu={c.expected}{extra}"
            )
        verdict = "INTÉGRITÉ OK" if self.integrity_ok else "ALERTE INTÉGRITÉ"
        lines.append(f"  => {verdict}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# LOADER CATALOGUE (BLOC 2)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CatalogueCleaningReport:
    """Compteurs de nettoyage de la source « catalogue ».

    Attributes:
        rows: Nombre de produits.
        n_categories_raw: Catégories distinctes AVANT normalisation (variantes incluses).
        n_categories_canonical: Catégories distinctes APRÈS normalisation —
            checkpoint oracle ``nb_categories_apres_normalisation``.
        n_products_recategorized: Produits dont la catégorie a été corrigée
            (cross-check vs ``nb_produits_impactes`` du manifeste).
    """

    rows: int
    n_categories_raw: int
    n_categories_canonical: int
    n_products_recategorized: int


def load_catalogue(raw_dir: Path) -> tuple[pd.DataFrame, CatalogueCleaningReport]:
    """Lit, type et normalise les catégories de la source « catalogue ».

    L'anomalie traitée ici est la plus dangereuse du jeu de données : des variantes
    orthographiques de catégorie (``"consommables"``, ``"Conso."``,
    ``"hygiene et entretien"``) qui, non corrigées, éclatent une catégorie en
    plusieurs et faussent **silencieusement** l'analyse de mix (le cœur du récit du
    Projet 2). La correction passe par ``config.canonical_category`` : la table de
    correspondance vit dans ``config``, donc une nouvelle variante se règle sans
    toucher au code.

    Args:
        raw_dir: Répertoire contenant ``catalogue_produits.xlsx``.

    Returns:
        Un tuple ``(df, report)`` : le catalogue propre (catégories canoniques,
        coûts/prix en float) et son ``CatalogueCleaningReport``.

    Raises:
        FileNotFoundError: Si le fichier catalogue est absent.
        ValueError: Si une colonne requise est absente (via ``check_columns``).
    """
    spec = SOURCES["catalogue"]
    files = sorted(Path(raw_dir).glob(spec.filename_glob))
    if not files:
        raise FileNotFoundError(
            f"Aucun fichier '{spec.filename_glob}' dans {raw_dir}."
        )
    path = files[0]
    digest = hash_file_on_disk(path)
    logging.info("[catalogue] lu %s (sha256=%s…)", path.name, digest[:12])

    df = pd.read_excel(path, dtype=str)
    df = normalize_headers(df, "catalogue")
    check_columns(df, "catalogue")
    df = coerce_types(df, "catalogue")

    # --- Normalisation métier des catégories (variante -> canonique) ---
    raw_category = df["categorie"]
    df["categorie"] = raw_category.map(canonical_category)
    n_recategorized = int((df["categorie"] != raw_category).sum())

    report = CatalogueCleaningReport(
        rows=len(df),
        n_categories_raw=int(raw_category.nunique()),
        n_categories_canonical=int(df["categorie"].nunique()),
        n_products_recategorized=n_recategorized,
    )
    logging.info(
        "[catalogue] %d produits | %d catégories brutes -> %d canoniques "
        "(%d produits recatégorisés)",
        report.rows, report.n_categories_raw,
        report.n_categories_canonical, report.n_products_recategorized,
    )
    return df, report


# ---------------------------------------------------------------------------
# LECTURE CSV ROBUSTE (helper privé)
# ---------------------------------------------------------------------------


def _read_csv_robust(path: Path, source_name: str) -> pd.DataFrame:
    """Lit un CSV en respectant l'encodage et le séparateur déclarés dans le schéma.

    L'encodage déclaré dans ``SourceSpec.encoding`` est essayé **en premier** (ex:
    ``cp1252`` pour le CRM Windows), puis les encodages de repli de
    ``SETTINGS.csv_encodings_fallback``. Prioriser l'encodage déclaré évite le piège
    classique : ``latin-1`` décode n'importe quel octet sans lever d'erreur et
    produirait du mojibake silencieux ; on ne l'utilise qu'en dernier recours.

    Lecture en ``dtype=str`` : aucune conversion implicite (les types sont posés
    ensuite par ``coerce_types``, de façon contrôlée et traçable).

    Args:
        path: Chemin du fichier CSV.
        source_name: Clé de ``config.SOURCES`` (fournit encodage et séparateur).

    Returns:
        Le DataFrame brut (toutes colonnes en ``str``).

    Raises:
        ValueError: Si aucun encodage candidat ne permet de lire le fichier.
    """
    spec = SOURCES[source_name]
    sep = spec.separator or ";"
    # Encodage déclaré d'abord, puis repli ; doublons retirés en gardant l'ordre.
    candidates = [spec.encoding, *SETTINGS.csv_encodings_fallback]
    encodings = list(dict.fromkeys(e for e in candidates if e))

    last_error: Exception | None = None
    for enc in encodings:
        try:
            df = pd.read_csv(path, sep=sep, dtype=str, encoding=enc)
            logging.info("[%s] lu %s (encodage=%s)", source_name, path.name, enc)
            return df
        except (UnicodeDecodeError, UnicodeError) as exc:
            last_error = exc
            logging.warning(
                "[%s] échec lecture %s en %s, encodage suivant…",
                source_name, path.name, enc,
            )
    raise ValueError(
        f"Lecture impossible de {path.name} : aucun encodage parmi {encodings} "
        f"n'a fonctionné. Dernière erreur : {last_error}"
    )


# ---------------------------------------------------------------------------
# LOADERS CRM & COMMERCIAUX (BLOC 3)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CrmCleaningReport:
    """Compteurs de nettoyage de la source « crm ».

    Attributes:
        rows: Nombre de clients (préservé : aucune ligne supprimée).
        invalid_first_order_dates: Dates de première commande non parseables
            (``NaT``) — **tolérées**, le client est conservé. Cross-check vs
            ``date_premiere_commande_invalide`` du manifeste.
        segments_normalized: Lignes dont le segment a été corrigé (casse/espaces).
        invalid_segments: Segments restant hors référentiel après normalisation
            (doit être 0 : preuve que la normalisation est exhaustive).
        distinct_segments: Segments canoniques distincts obtenus (attendu : 3).
    """

    rows: int
    invalid_first_order_dates: int
    segments_normalized: int
    invalid_segments: int
    distinct_segments: int


def load_crm(raw_dir: Path) -> tuple[pd.DataFrame, CrmCleaningReport]:
    """Lit le CRM (cp1252), type les colonnes et normalise les segments.

    Trois anomalies de saisie commerciale : encodage ``cp1252`` (accents),
    casse/espaces incohérents sur le segment, et dates de première commande
    invalides. Choix métier clé : une date de première commande invalide ne
    **disqualifie pas** le client (``date_premiere_commande`` est ``nullable``) —
    elle devient ``NaT``, contrairement à une ligne de vente non datée qui, elle,
    sort du périmètre comptable.

    Args:
        raw_dir: Répertoire contenant ``export_clients_crm.csv``.

    Returns:
        Un tuple ``(df, report)`` : le CRM propre (segments canoniques) et son rapport.

    Raises:
        FileNotFoundError: Si le fichier CRM est absent.
        ValueError: Si une colonne requise est absente (via ``check_columns``).
    """
    spec = SOURCES["crm"]
    files = sorted(Path(raw_dir).glob(spec.filename_glob))
    if not files:
        raise FileNotFoundError(f"Aucun fichier '{spec.filename_glob}' dans {raw_dir}.")
    path = files[0]
    digest = hash_file_on_disk(path)
    logging.info("[crm] sha256=%s…", digest[:12])

    df = _read_csv_robust(path, "crm")
    df = normalize_headers(df, "crm")
    check_columns(df, "crm")
    df = coerce_types(df, "crm")  # date_premiere_commande -> NaT si invalide (toléré)

    invalid_dates = int(df["date_premiere_commande"].isna().sum())

    # --- Normalisation métier des segments (casse/espaces -> canonique) ---
    raw_segment = df["segment"]
    df["segment"] = raw_segment.map(canonical_segment)
    segments_normalized = int((df["segment"] != raw_segment).sum())
    invalid_segments = int((~df["segment"].isin(SEGMENTS_CANONICAL)).sum())

    report = CrmCleaningReport(
        rows=len(df),
        invalid_first_order_dates=invalid_dates,
        segments_normalized=segments_normalized,
        invalid_segments=invalid_segments,
        distinct_segments=int(df["segment"].nunique()),
    )
    logging.info(
        "[crm] %d clients | %d dates 1re commande KO (tolérées) | "
        "%d segments normalisés -> %d distincts, %d invalides",
        report.rows, report.invalid_first_order_dates,
        report.segments_normalized, report.distinct_segments, report.invalid_segments,
    )
    return df, report


@dataclass(frozen=True)
class CommerciauxCleaningReport:
    """Compteurs de la source « commerciaux » (référentiel propre, contraste).

    Attributes:
        rows: Nombre de commerciaux.
        invalid_hire_dates: Dates d'embauche non parseables (attendu : 0, source propre).
        duplicate_codes: Codes commerciaux dupliqués (attendu : 0, clé primaire).
    """

    rows: int
    invalid_hire_dates: int
    duplicate_codes: int


def load_commerciaux(raw_dir: Path) -> tuple[pd.DataFrame, CommerciauxCleaningReport]:
    """Lit le référentiel commerciaux (utf-8, propre) et type ses colonnes.

    Cette source est volontairement **sans anomalie** : elle sert de contraste et
    de témoin (si le pipeline introduisait un défaut, il apparaîtrait ici sur des
    données saines). À noter : ``date_embauche`` est au format **ISO** (``AAAA-MM-JJ``),
    d'où l'intérêt du parsing multi-format de ``parse_date``.

    Args:
        raw_dir: Répertoire contenant ``commerciaux.csv``.

    Returns:
        Un tuple ``(df, report)`` : le référentiel propre et son rapport.

    Raises:
        FileNotFoundError: Si le fichier est absent.
        ValueError: Si une colonne requise est absente (via ``check_columns``).
    """
    spec = SOURCES["commerciaux"]
    files = sorted(Path(raw_dir).glob(spec.filename_glob))
    if not files:
        raise FileNotFoundError(f"Aucun fichier '{spec.filename_glob}' dans {raw_dir}.")
    path = files[0]
    digest = hash_file_on_disk(path)
    logging.info("[commerciaux] sha256=%s…", digest[:12])

    df = _read_csv_robust(path, "commerciaux")
    df = normalize_headers(df, "commerciaux")
    check_columns(df, "commerciaux")
    df = coerce_types(df, "commerciaux")  # date_embauche ISO -> datetime

    report = CommerciauxCleaningReport(
        rows=len(df),
        invalid_hire_dates=int(df["date_embauche"].isna().sum()),
        duplicate_codes=int(df["code_commercial"].duplicated().sum()),
    )
    logging.info(
        "[commerciaux] %d commerciaux | %d dates KO | %d codes dupliqués",
        report.rows, report.invalid_hire_dates, report.duplicate_codes,
    )
    return df, report


# ---------------------------------------------------------------------------
# REMODELAGE BUDGET (BLOC 4) - format large -> format long
# ---------------------------------------------------------------------------


def _extract_budget_year(path: Path) -> int:
    """Extrait l'année d'un fichier budget depuis son nom (``budget_AAAA.xlsx``).

    Args:
        path: Chemin du fichier budget.

    Returns:
        L'année (ex: 2024).

    Raises:
        ValueError: Si aucune année à 4 chiffres n'est trouvée dans le nom.
    """
    match = re.search(r"(\d{4})", path.stem)
    if not match:
        raise ValueError(
            f"Impossible d'extraire l'année du fichier budget : {path.name}"
        )
    return int(match.group(1))


def _reshape_budget_sheet(
    path: Path,
    sheet: str,
    metric_col: str,
    year: int,
    layout: dict,
) -> tuple[pd.DataFrame, int]:
    """Remodèle une feuille budget « large » en table longue pour une métrique.

    Étapes :
        1. Lecture à partir de la ligne d'en-tête (``header_row``) : le titre et la
           ligne vide au-dessus sont ignorés par pandas.
        2. ``ffill`` sur la colonne région : les cellules fusionnées ne portent la
           valeur que sur leur première ligne (NaN ensuite).
        3. Suppression des lignes de sous-total (``subtotal_marker``) : ce sont des
           agrégats, pas des données — les garder doublerait le budget.
        4. ``melt`` des 12 colonnes de mois en lignes (large -> long).
        5. Construction de ``periode`` ``AAAA-MM`` : le numéro de mois est la
           **position** du libellé dans ``month_labels`` (pas de dictionnaire de mois
           français à maintenir).

    Args:
        path: Chemin du fichier budget.
        sheet: Nom de la feuille (ex: ``"CA"``).
        metric_col: Nom canonique de la métrique produite (ex: ``"ca_budgete"``).
        year: Année du fichier (pour ``periode``).
        layout: Dictionnaire de mise en page de ``SOURCES["budget"].layout``.

    Returns:
        Un couple ``(df_long, n_subtotals)`` : le DataFrame long
        ``[periode, nom_region, categorie, <metric_col>]`` et le nombre de lignes de
        sous-total écartées (traçabilité, sans relecture du fichier).
    """
    header_idx = int(layout["header_row"]) - 1  # 1-indexé (métier) -> 0-indexé (pandas)
    id_columns = list(layout["id_columns"])
    month_labels = list(layout["month_labels"])
    subtotal_marker = layout["subtotal_marker"]
    region_col, category_col = id_columns

    df = pd.read_excel(path, sheet_name=sheet, header=header_idx, dtype=object)
    df = df[id_columns + month_labels]
    df[region_col] = df[region_col].ffill()  # propage la région des cellules fusionnées
    is_subtotal = df[category_col] == subtotal_marker
    n_subtotals = int(is_subtotal.sum())
    df = df[~is_subtotal].copy()  # écarte les sous-totaux

    long = df.melt(
        id_vars=id_columns,
        value_vars=month_labels,
        var_name="mois_label",
        value_name=metric_col,
    )
    month_number = {label: i + 1 for i, label in enumerate(month_labels)}
    long["periode"] = long["mois_label"].map(month_number).map(
        lambda mm: f"{year}-{mm:02d}"
    )
    long = long.rename(columns={region_col: "nom_region", category_col: "categorie"})
    return long[["periode", "nom_region", "categorie", metric_col]], n_subtotals


@dataclass(frozen=True)
class BudgetCleaningReport:
    """Compteurs du remodelage budget.

    Attributes:
        files_read: Nombre de fichiers budget (un par année).
        rows: Lignes de la table longue finale.
        n_regions: Régions distinctes.
        n_categories: Catégories distinctes.
        n_periods: Périodes distinctes (mois × années).
        subtotal_rows_dropped: Lignes de sous-total écartées (toutes feuilles).
        leaked_subtotals: Sous-totaux ayant échappé au filtre (doit être 0).
        missing_amounts: Montants ``NaN`` après fusion CA/Marge (doit être 0 :
            preuve que les deux feuilles couvrent exactement la même grille).
        duplicate_keys: Clés ``(periode, region, categorie)`` en double (doit être 0).
    """

    files_read: int
    rows: int
    n_regions: int
    n_categories: int
    n_periods: int
    subtotal_rows_dropped: int
    leaked_subtotals: int
    missing_amounts: int
    duplicate_keys: int


def reshape_budget(raw_dir: Path) -> tuple[pd.DataFrame, BudgetCleaningReport]:
    """Lit les budgets « larges » multi-feuilles et produit une table longue propre.

    Chaque fichier (une année) contient deux feuilles (``CA``, ``Marge``) au format
    large : titre, en-têtes en ligne 3, régions en cellules fusionnées, 12 mois en
    colonnes, sous-totaux intercalés. On remodèle chaque feuille puis on fusionne
    les métriques sur la clé ``(periode, nom_region, categorie)`` : chaque ligne
    finale porte donc CA **et** Marge budgétés, prêts à être comparés au réalisé
    (Projet 2).

    Args:
        raw_dir: Répertoire contenant ``budget_*.xlsx``.

    Returns:
        Un tuple ``(df, report)`` : la table budget longue canonique et son rapport.

    Raises:
        FileNotFoundError: Si aucun fichier budget n'est trouvé.
        ValueError: Si une colonne canonique requise est absente (via ``check_columns``).
    """
    spec = SOURCES["budget"]
    layout = spec.layout
    subtotal_marker = layout["subtotal_marker"]
    keys = ["periode", "nom_region", "categorie"]

    files = sorted(Path(raw_dir).glob(spec.filename_glob))
    if not files:
        raise FileNotFoundError(f"Aucun fichier '{spec.filename_glob}' dans {raw_dir}.")

    frames: list[pd.DataFrame] = []
    subtotal_rows_dropped = 0
    for path in files:
        digest = hash_file_on_disk(path)
        year = _extract_budget_year(path)
        logging.info("[budget] lu %s (année=%d, sha256=%s…)", path.name, year, digest[:12])

        per_metric: list[pd.DataFrame] = []
        for sheet, metric_col in spec.sheet_to_metric.items():
            long, n_subtotals = _reshape_budget_sheet(path, sheet, metric_col, year, layout)
            subtotal_rows_dropped += n_subtotals
            per_metric.append(long)
        # Fusion des métriques (CA, Marge) sur la grille commune.
        merged = reduce(lambda left, right: left.merge(right, on=keys, how="outer"), per_metric)
        frames.append(merged)

    budget = pd.concat(frames, ignore_index=True)
    budget = coerce_types(budget, "budget")
    check_columns(budget, "budget")

    metric_cols = list(spec.sheet_to_metric.values())
    report = BudgetCleaningReport(
        files_read=len(files),
        rows=len(budget),
        n_regions=int(budget["nom_region"].nunique()),
        n_categories=int(budget["categorie"].nunique()),
        n_periods=int(budget["periode"].nunique()),
        subtotal_rows_dropped=subtotal_rows_dropped,
        leaked_subtotals=int((budget["categorie"] == subtotal_marker).sum()),
        missing_amounts=int(budget[metric_cols].isna().any(axis=1).sum()),
        duplicate_keys=int(budget.duplicated(subset=keys).sum()),
    )
    logging.info(
        "[budget] %d fichiers | %d lignes | %d régions × %d catégories × %d périodes "
        "| %d sous-totaux écartés | %d montants manquants | %d clés dupliquées",
        report.files_read, report.rows, report.n_regions, report.n_categories,
        report.n_periods, report.subtotal_rows_dropped, report.missing_amounts,
        report.duplicate_keys,
    )
    return budget, report



# ---------------------------------------------------------------------------
# CONSOLIDATION STAR SCHEMA (BLOC 5) - jointures DuckDB
# ---------------------------------------------------------------------------


def build_star_schema(cleaned: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Assemble le modèle en étoile canonique à partir des sources nettoyées (DuckDB).

    On enregistre les DataFrames propres dans une connexion DuckDB in-memory, puis on
    produit chaque table de ``config.STAR_SCHEMA`` par SQL. Choix de modélisation :

    - ``dim_region`` provient du référentiel ``config.REGIONS`` (aucune source
      transactionnelle ne relie code et nom) ; il sert aussi à rattacher le budget.
    - La source « ventes » est dénormalisée au grain ligne : on la scinde en
      ``fait_commande`` (grain commande, porte ``statut`` — constant par commande) et
      ``fait_ligne_commande`` (grain ligne, porte les composantes du CA). La marge et
      le CA ne sont **pas** stockés : ils se recalculent en SQL.
    - ``fait_budget`` rattache le budget (qui porte ``nom_region``) à un
      ``code_region`` via ``dim_region``.

    Frontière assumée : toute la jointure inter-sources se fait en SQL (et non en
    pandas), conformément au principe « pandas nettoie, DuckDB consolide ».

    Args:
        cleaned: Dictionnaire des DataFrames propres, par nom de source
            (``ventes``, ``crm``, ``catalogue``, ``commerciaux``, ``budget``).

    Returns:
        Un dictionnaire ``{nom_table: DataFrame}`` couvrant exactement les 7 tables
        de ``config.STAR_SCHEMA``, colonnes ordonnées selon le schéma.
    """
    regions_df = pd.DataFrame(
        {"code_region": list(REGIONS.keys()), "nom_region": list(REGIONS.values())}
    )

    # Une requête par table de faits/dimension. Les colonnes sont sélectionnées et
    # ordonnées explicitement (la validation finale se fait contre STAR_SCHEMA).
    queries: dict[str, str] = {
        "dim_region": "SELECT code_region, nom_region FROM regions",
        "dim_commercial": """
            SELECT code_commercial, nom, code_region, date_embauche
            FROM commerciaux
        """,
        "dim_produit": """
            SELECT code_produit, libelle, categorie, sous_categorie,
                   cout_unitaire, prix_catalogue, fournisseur
            FROM catalogue
        """,
        "dim_client": """
            SELECT code_client, raison_sociale, segment, type_etablissement,
                   code_region, ville, code_commercial, date_premiere_commande
            FROM crm
        """,
        # Grain commande : une ligne par commande. statut/date/client/commercial
        # étant constants par commande, DISTINCT suffit et reste sans perte.
        "fait_commande": """
            SELECT DISTINCT commande_id, date_commande, code_client,
                   code_commercial, statut
            FROM ventes
        """,
        "fait_ligne_commande": """
            SELECT ligne_id, commande_id, code_produit, quantite,
                   prix_unitaire_vente, remise_pct
            FROM ventes
        """,
        # Rattachement budget -> code_region via le référentiel régions.
        "fait_budget": """
            SELECT b.periode, r.code_region, b.categorie,
                   b.ca_budgete, b.marge_budgetee
            FROM budget AS b
            JOIN regions AS r ON b.nom_region = r.nom_region
        """,
    }

    con = duckdb.connect()
    star: dict[str, pd.DataFrame] = {}
    try:
        con.register("regions", regions_df)
        for name, df in cleaned.items():
            con.register(name, df)
        for table, sql in queries.items():
            result = con.execute(sql).df()
            # Réordonne/restreint aux colonnes canoniques du schéma (garde-fou).
            result = result[list(STAR_SCHEMA[table])]
            star[table] = result
            logging.info("[star] %-20s %d lignes", table, len(result))
    finally:
        con.close()
    return star


# ---------------------------------------------------------------------------
# RÉCONCILIATION UNIFIÉE CONTRE L'ORACLE (BLOC 5)
# ---------------------------------------------------------------------------


def _count_orphans(
    con: "duckdb.DuckDBPyConnection",
    child: str,
    child_key: str,
    parent: str,
    parent_key: str,
) -> int:
    """Compte les clés étrangères orphelines (enfant sans parent correspondant).

    Args:
        con: Connexion DuckDB où ``child`` et ``parent`` sont enregistrés.
        child: Table enfant (porteuse de la clé étrangère).
        child_key: Colonne clé étrangère dans l'enfant.
        parent: Table parente (porteuse de la clé primaire).
        parent_key: Colonne clé primaire dans le parent.

    Returns:
        Le nombre de lignes enfant dont la clé n'existe pas dans le parent.
    """
    query = f"""
        SELECT COUNT(*) FROM {child} AS c
        WHERE c.{child_key} IS NOT NULL
          AND NOT EXISTS (
              SELECT 1 FROM {parent} AS p WHERE p.{parent_key} = c.{child_key}
          )
    """
    return int(con.execute(query).fetchone()[0])


def reconcile_against_manifest(
    star: dict[str, pd.DataFrame],
    reports: dict[str, object],
    manifest: dict,
) -> ReconciliationReport:
    """Checkpoint d'intégrité unifié : confronte le pipeline complet à l'oracle.

    Remplace les réconciliations par bloc. Le rapport agrège :

    - **Les 5 métriques officielles** de ``attendus_apres_nettoyage`` : lignes après
      dédup, dates invalides, montants invalides (depuis le rapport ventes), CA
      réconcilié (recalculé en SQL sur les faits), catégories canoniques.
    - **Des contrôles structurels** complémentaires : segments CRM tous canoniques,
      grille budget complète, et **intégrité référentielle** (aucune clé étrangère
      orpheline entre faits et dimensions) — la preuve que 100 % des transactions se
      rattachent au modèle.

    Args:
        star: Tables du star schema (sortie de ``build_star_schema``).
        reports: Rapports de nettoyage par source (``ventes``, ``catalogue``,
            ``crm``, ``budget``...).
        manifest: Manifeste chargé (doit contenir ``attendus_apres_nettoyage``).

    Returns:
        Le ``ReconciliationReport`` complet (oracle + contrôles structurels).
    """
    expected = manifest["attendus_apres_nettoyage"]
    sales = reports["ventes"]
    catalogue = reports["catalogue"]
    crm = reports["crm"]
    budget = reports["budget"]
    checks: list[ReconciliationCheck] = []

    # --- 5 métriques officielles de l'oracle ---
    oracle_int = [
        ("Lignes après dédup", sales.rows_after_dedup, "nb_lignes_apres_dedup"),
        ("Dates invalides", sales.invalid_dates, "nb_dates_invalides_attendu"),
        ("Montants invalides", sales.invalid_amounts, "nb_montants_invalides_attendu"),
        ("Catégories canoniques", catalogue.n_categories_canonical,
         "nb_categories_apres_normalisation"),
    ]
    for label, obtained, key in oracle_int:
        want = expected[key]
        checks.append(ReconciliationCheck(label, obtained, want, ok=(obtained == want)))

    # CA recalculé en SQL sur les faits (tolérance relative pour les arrondis).
    ca = compute_ca_livree(star["fait_ligne_commande"], star["fait_commande"])
    want_ca = float(expected["ca_reconcilie_attendu_eur"])
    gap = abs(ca - want_ca)
    tol = SETTINGS.reconciliation_tolerance_pct
    gap_pct = gap / want_ca if want_ca else (0.0 if ca == 0 else 1.0)
    checks.append(
        ReconciliationCheck(
            "CA réconcilié (€)", round(ca, 2), want_ca, ok=(gap_pct <= tol),
            detail=f"écart {gap:.4f} € = {gap_pct * 100:.5f} % ≤ tol {tol * 100:.3f} %",
        )
    )

    # --- Contrôles structurels complémentaires ---
    checks.append(
        ReconciliationCheck(
            "Segments CRM canoniques", crm.invalid_segments, 0,
            ok=(crm.invalid_segments == 0),
            detail=f"{crm.distinct_segments} segments distincts",
        )
    )
    expected_budget_rows = budget.n_regions * budget.n_categories * budget.n_periods
    checks.append(
        ReconciliationCheck(
            "Grille budget complète", budget.rows, expected_budget_rows,
            ok=(budget.rows == expected_budget_rows
                and budget.leaked_subtotals == 0
                and budget.duplicate_keys == 0),
            detail=f"{budget.missing_amounts} montant(s) manquant(s)",
        )
    )

    # --- Intégrité référentielle (faits -> dimensions) en SQL ---
    con = duckdb.connect()
    try:
        for name, df in star.items():
            con.register(name, df)
        fk_specs = [
            ("Lignes -> Produits", "fait_ligne_commande", "code_produit",
             "dim_produit", "code_produit"),
            ("Lignes -> Commandes", "fait_ligne_commande", "commande_id",
             "fait_commande", "commande_id"),
            ("Commandes -> Clients", "fait_commande", "code_client",
             "dim_client", "code_client"),
            ("Commandes -> Commerciaux", "fait_commande", "code_commercial",
             "dim_commercial", "code_commercial"),
            ("Budget -> Régions", "fait_budget", "code_region",
             "dim_region", "code_region"),
        ]
        for label, child, ckey, parent, pkey in fk_specs:
            orphans = _count_orphans(con, child, ckey, parent, pkey)
            checks.append(
                ReconciliationCheck(
                    f"FK {label}", orphans, 0, ok=(orphans == 0),
                    detail="clés étrangères orphelines",
                )
            )
    finally:
        con.close()

    recon = ReconciliationReport(checks=checks)
    logging.info("[Réconciliation unifiée]\n%s", recon.render())
    return recon


def run_ingestion(
    raw_dir: Path,
) -> tuple[dict[str, pd.DataFrame], dict[str, object]]:
    """Orchestre l'ingestion complète : charge les 5 sources et construit le star schema.

    C'est le point d'entrée logique du pipeline (appelé par le mode Batch et l'UI) :
    il enchaîne les loaders, regroupe les DataFrames propres et leurs rapports, puis
    consolide en modèle en étoile. La réconciliation se fait ensuite via
    ``reconcile_against_manifest``.

    Args:
        raw_dir: Répertoire des fichiers sources.

    Returns:
        Un couple ``(star, reports)`` : les tables du star schema et les rapports de
        nettoyage par source.
    """
    ventes, r_ventes = load_ventes(raw_dir)
    catalogue, r_catalogue = load_catalogue(raw_dir)
    crm, r_crm = load_crm(raw_dir)
    commerciaux, r_commerciaux = load_commerciaux(raw_dir)
    budget, r_budget = reshape_budget(raw_dir)

    cleaned = {
        "ventes": ventes,
        "catalogue": catalogue,
        "crm": crm,
        "commerciaux": commerciaux,
        "budget": budget,
    }
    reports = {
        "ventes": r_ventes,
        "catalogue": r_catalogue,
        "crm": r_crm,
        "commerciaux": r_commerciaux,
        "budget": r_budget,
    }
    star = build_star_schema(cleaned)
    return star, reports


# ---------------------------------------------------------------------------
# PONT STAR SCHEMA -> VUES DE REPORTING (BLOC 6)
#
# Le dashboard Excel/Streamlit (workbook.py) consomme trois structures héritées
# du modèle de présentation : un DataFrame transactionnel et deux agrégats.
# Ce bloc les *dérive* du star schema en SQL — il ne stocke rien. Le CA reste
# recalculé à la volée (composantes de la ligne), jamais matérialisé : c'est la
# frontière « pandas nettoie, DuckDB consolide, le CA se recalcule ».
# ---------------------------------------------------------------------------


def load_manifest(raw_dir: Path) -> dict:
    """Charge le manifeste-oracle (``_manifest_anomalies.json``) du répertoire sources.

    Le manifeste est la source de vérité de la réconciliation : il porte les
    attendus chiffrés (``attendus_apres_nettoyage``) auxquels le pipeline se compare.
    On échoue explicitement s'il est absent — un run d'audit sans oracle n'a pas de
    sens (on ne pourrait rien prouver).

    Args:
        raw_dir: Répertoire des fichiers sources (où réside le manifeste).

    Returns:
        Le manifeste désérialisé en dictionnaire.

    Raises:
        FileNotFoundError: Si le manifeste est introuvable dans ``raw_dir``.
    """
    manifest_path = Path(raw_dir) / SETTINGS.manifest_name
    if not manifest_path.exists():
        raise FileNotFoundError(
            f"Manifeste-oracle introuvable : {manifest_path}. "
            "Génère les données démo (generate_demo_data.py) ou place le manifeste "
            "dans le répertoire des sources."
        )
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def build_reporting_views(
    star: dict[str, pd.DataFrame],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Dérive les 3 vues de reporting du star schema, en SQL (CA jamais stocké).

    Produit le triplet attendu par ``workbook.py`` :

    1. ``df`` — transactions au **grain ligne**, commandes **livrées** uniquement.
       Colonnes : ``date`` (date de commande), ``mois`` (``YYYY-MM``),
       ``commercial`` (nom), ``montant`` (CA de la ligne, recalculé).
    2. ``report_months`` — agrégat ``[mois, montant]``.
    3. ``report_salespeople`` — agrégat ``[commercial, montant]``.

    **Choix de modélisation (justifiés)** :

    - **Filtre ``statut`` livré** (``config.STATUS_INCLUDED_IN_CA``) : le total du
      dashboard devient *strictement égal* au CA réconcilié (cohérence d'un chiffre
      unique entre la feuille Dashboard et la feuille Réconciliation).
    - **``LEFT JOIN`` sur ``dim_commercial``** : l'intégrité référentielle garantit
      déjà 0 orphelin (donc ``LEFT`` ≡ ``INNER`` ici), mais on choisit ``LEFT`` par
      principe défensif : une dégradation future du référentiel rendrait la perte
      *visible* (``commercial = None``) au lieu de faire disparaître silencieusement
      du CA.
    - **``montant`` laissé à ``NaN`` quand le prix est invalide** : ``groupby.sum()``
      ignore les ``NaN`` (la somme recolle au CA livré), tout en conservant la ligne
      dans ``df``. Le compteur de transactions valides (KPI) reste ainsi cohérent
      avec la métrique « montants invalides » de l'oracle — on ne maquille pas une
      ligne KO en 0 €.
    - **``mois`` issu de ``date_commande``** : une date invalide (``NaT``) donne un
      ``mois`` nul, exclu des agrégats mensuels mais conservé dans ``df``.

    Args:
        star: Tables du star schema (sortie de ``build_star_schema`` /
            ``run_ingestion``).

    Returns:
        Le triplet ``(df, report_months, report_salespeople)``.
    """
    placeholders = ", ".join("?" for _ in STATUS_INCLUDED_IN_CA)
    con = duckdb.connect()
    try:
        for name, table in star.items():
            con.register(name, table)
        # CA de la ligne = quantité × prix × (1 - remise). NULL si prix manquant
        # (NULL se propage dans l'expression -> NaN côté pandas), jamais stocké.
        df = con.execute(
            f"""
            SELECT
                c.date_commande                          AS date,
                strftime(c.date_commande, '%Y-%m')       AS mois,
                s.nom                                     AS commercial,
                l.quantite * l.prix_unitaire_vente
                    * (1 - l.remise_pct)                 AS montant
            FROM fait_ligne_commande AS l
            JOIN fait_commande       AS c ON l.commande_id     = c.commande_id
            LEFT JOIN dim_commercial AS s ON c.code_commercial = s.code_commercial
            WHERE c.statut IN ({placeholders})
            """,
            list(STATUS_INCLUDED_IN_CA),
        ).df()
    finally:
        con.close()

    # Agrégats en pandas : groupby.sum() ignore les NaN (montants invalides) et
    # dropna sur la clé écarte les lignes sans mois/commercial exploitable.
    report_months = (
        df.dropna(subset=["mois"])
        .groupby("mois", as_index=False)["montant"]
        .sum()
        .sort_values("mois")
        .reset_index(drop=True)
    )
    report_salespeople = (
        df.dropna(subset=["commercial"])
        .groupby("commercial", as_index=False)["montant"]
        .sum()
        .sort_values("montant", ascending=False)
        .reset_index(drop=True)
    )

    logging.info(
        "[vues reporting] %d lignes livrées | %d mois | %d commerciaux | CA %.2f €",
        len(df),
        len(report_months),
        len(report_salespeople),
        float(df["montant"].sum()),
    )
    return df, report_months, report_salespeople


# ---------------------------------------------------------------------------
# COUCHE ANALYTIQUE DAF (BLOC 7)
#
# Distincte du « reporting opérationnel » (build_reporting_views, centré CA) :
# cette couche produit les agrégats d'aide à la décision pour une direction
# financière — marge, écart au budget, mix catégoriel, comparaison annuelle.
# Comme partout, le CA ET la marge sont recalculés en SQL, jamais stockés : la
# table de faits ne porte que les composantes (quantité, prix, remise) et la
# dimension produit porte le coût unitaire. Le budget (ca_budgete, marge_budgetee)
# est lui un *objectif* fourni en entrée — il est légitimement stocké.
# ---------------------------------------------------------------------------


def build_analytics_views(
    star: dict[str, pd.DataFrame],
    *,
    months: Optional[Sequence[str]] = None,
    regions: Optional[Sequence[str]] = None,
    segments: Optional[Sequence[str]] = None,
    salespeople: Optional[Sequence[str]] = None,
) -> dict[str, pd.DataFrame]:
    """Dérive les vues analytiques DAF du star schema, en SQL (marge recalculée).

    Toutes les vues sont restreintes aux commandes **livrées** (cohérence avec le CA
    réconcilié). La marge brute d'une ligne vaut ``CA - coût`` avec
    ``CA = quantité × prix × (1 - remise)`` et ``coût = quantité × coût_unitaire``.
    L'écart au budget suit la convention financière ``réel - budget`` : un écart
    négatif est défavorable (sous le plan) sur le CA comme sur la marge.

    Vues retournées (clés du dictionnaire) :

    - ``summary`` (1 ligne) : CA, marge, taux de marge, budget CA/marge, écarts
      (€ et %), croissance YoY du CA, variation YoY du taux de marge (en points),
      nombre de clients actifs. Alimente la feuille de synthèse.
    - ``by_year`` : CA, marge, taux de marge et clients actifs par année (base du YoY).
    - ``monthly`` : CA, marge, taux de marge par mois (courbe d'érosion de la marge).
    - ``by_category`` : CA, marge, taux de marge et part de CA par catégorie produit.
    - ``budget_monthly`` : réel vs budget (CA et marge) par mois, avec écarts.
    - ``budget_by_region`` : réel vs budget (CA et marge) par région, avec écarts.
    - ``category_mix`` : part de CA par catégorie pour chaque année (effet mix).
    - ``by_salesperson`` : CA, marge, taux et part par commercial.
    - ``by_segment`` : CA, marge, taux, nb de clients et part par segment client.
    - ``client_pareto`` : CA cumulé par client (rang trié), pour la courbe ABC.
    - ``client_top`` : les 15 premiers clients (CA, part, taux de marge).
    - ``by_discount_band`` : nb de lignes, CA, taux et part par tranche de remise.

    Filtres de présentation (optionnels, ``None`` = aucun filtre, comportement
    rétrocompatible avec le comportement antérieur). Ils sont injectés en SQL dans la vue ``_base``,
    si bien que les 12 vues restent cohérentes entre elles. Le budget n'étant
    dimensionné que par mois et région, seuls ``months`` et ``regions`` filtrent les
    CTE budgétaires ; un filtre ``segments`` ou ``salespeople`` ne réduit que le réel,
    à charge pour la couche de présentation de ne pas afficher d'écart au budget dans
    ce cas (le budget resterait sur 100 % du périmètre, l'écart serait trompeur).

    Args:
        star: Tables du star schema (sortie de ``run_ingestion``).
        months: Mois (``'AAAA-MM'``) à conserver, ou ``None`` pour tous.
        regions: Noms de régions à conserver, ou ``None`` pour toutes.
        segments: Segments clients à conserver, ou ``None`` pour tous.
        salespeople: Noms de commerciaux à conserver, ou ``None`` pour tous.

    Returns:
        Un dictionnaire ``{nom_vue: DataFrame}``.
    """
    con = duckdb.connect()
    try:
        for name, table in star.items():
            con.register(name, table)

        # Statuts inclus dans le CA : constantes internes, injectées en littéraux SQL
        # échappés (un paramètre lié n'est pas autorisé dans un CREATE VIEW DuckDB).
        status_literals = ", ".join(
            "'" + s.replace("'", "''") + "'" for s in STATUS_INCLUDED_IN_CA
        )

        def _in_clause(column: str, values: Optional[Sequence[str]]) -> str:
            """Construit ``AND column IN (...)`` échappé, ou '' si aucun filtre.

            Même contrainte que ``status_literals`` : DuckDB n'autorise pas de
            paramètre lié dans un ``CREATE VIEW``, d'où l'injection par littéraux
            échappés. Les valeurs proviennent du référentiel (mois, régions,
            segments, commerciaux), jamais d'une saisie libre.
            """
            if not values:
                return ""
            literals = ", ".join(
                "'" + str(v).replace("'", "''") + "'" for v in values
            )
            return f" AND {column} IN ({literals})"

        # Filtres de présentation injectés dans _base (le réel). Les axes sont les
        # expressions sous-jacentes : DuckDB ne référence pas les alias du SELECT
        # dans le WHERE.
        base_filters = (
            _in_clause("strftime(c.date_commande, '%Y-%m')", months)
            + _in_clause("r.nom_region", regions)
            + _in_clause("cl.segment", segments)
            + _in_clause("s.nom", salespeople)
        )
        # Le budget n'a que deux axes exploitables : mois (periode) et région.
        budget_period_filter = _in_clause("periode", months)
        budget_region_filter = _in_clause("nom_region", regions)

        # Base ligne livrée enrichie : CA et coût par ligne (jamais stockés), plus
        # les axes d'analyse (mois, année, catégorie, région, client). LEFT JOIN
        # défensif sur les dimensions (l'intégrité FK garantit 0 orphelin, mais une
        # dégradation future resterait ainsi visible plutôt que muette).
        con.execute(
            f"""
            CREATE TEMP VIEW _base AS
            SELECT
                strftime(c.date_commande, '%Y-%m')      AS mois,
                strftime(c.date_commande, '%Y')         AS annee,
                c.code_client                           AS code_client,
                cl.raison_sociale                       AS raison_sociale,
                cl.segment                              AS segment,
                p.categorie                             AS categorie,
                r.nom_region                            AS region,
                s.nom                                   AS commercial,
                l.remise_pct                            AS remise,
                l.quantite * l.prix_unitaire_vente
                    * (1 - l.remise_pct)                AS ca,
                l.quantite * p.cout_unitaire            AS cout
            FROM fait_ligne_commande AS l
            JOIN fait_commande       AS c  ON l.commande_id     = c.commande_id
            LEFT JOIN dim_produit    AS p  ON l.code_produit    = p.code_produit
            LEFT JOIN dim_client     AS cl ON c.code_client     = cl.code_client
            LEFT JOIN dim_region     AS r  ON cl.code_region    = r.code_region
            LEFT JOIN dim_commercial AS s  ON c.code_commercial = s.code_commercial
            WHERE c.statut IN ({status_literals}){base_filters}
            """
        )

        # --- monthly : courbe d'érosion (taux de marge mensuel) ---
        monthly = con.execute(
            """
            SELECT mois,
                   sum(ca)                                   AS ca,
                   sum(ca - cout)                            AS marge,
                   100.0 * sum(ca - cout) / NULLIF(sum(ca), 0) AS taux_marge
            FROM _base
            GROUP BY mois ORDER BY mois
            """
        ).df()

        # --- by_year : base du YoY ---
        by_year = con.execute(
            """
            SELECT annee,
                   sum(ca)                                   AS ca,
                   sum(ca - cout)                            AS marge,
                   100.0 * sum(ca - cout) / NULLIF(sum(ca), 0) AS taux_marge,
                   count(DISTINCT code_client)               AS clients_actifs
            FROM _base
            GROUP BY annee ORDER BY annee
            """
        ).df()

        # --- by_category : CA, marge et part de CA par catégorie ---
        by_category = con.execute(
            """
            SELECT categorie,
                   sum(ca)                                    AS ca,
                   sum(ca - cout)                             AS marge,
                   100.0 * sum(ca - cout) / NULLIF(sum(ca), 0) AS taux_marge,
                   100.0 * sum(ca) / NULLIF(sum(sum(ca)) OVER (), 0) AS part_ca_pct
            FROM _base
            GROUP BY categorie ORDER BY ca DESC
            """
        ).df()

        # --- budget_monthly : réel vs budget (CA + marge), spine = budget ---
        budget_monthly = con.execute(
            f"""
            WITH act AS (
                SELECT mois, sum(ca) ca_reel, sum(ca - cout) marge_reel
                FROM _base GROUP BY mois
            ),
            bud AS (
                SELECT b.periode AS mois,
                       sum(b.ca_budgete)    AS ca_budget,
                       sum(b.marge_budgetee) AS marge_budget
                FROM fait_budget b
                JOIN dim_region r ON b.code_region = r.code_region
                WHERE 1 = 1{budget_period_filter}{budget_region_filter}
                GROUP BY b.periode
            )
            SELECT
                COALESCE(b.mois, a.mois)                 AS mois,
                COALESCE(a.ca_reel, 0)                   AS ca_reel,
                b.ca_budget                              AS ca_budget,
                COALESCE(a.ca_reel, 0) - b.ca_budget     AS ecart_ca,
                100.0 * (COALESCE(a.ca_reel, 0) - b.ca_budget)
                    / NULLIF(b.ca_budget, 0)             AS ecart_ca_pct,
                COALESCE(a.marge_reel, 0)                AS marge_reel,
                b.marge_budget                           AS marge_budget,
                COALESCE(a.marge_reel, 0) - b.marge_budget AS ecart_marge,
                100.0 * (COALESCE(a.marge_reel, 0) - b.marge_budget)
                    / NULLIF(b.marge_budget, 0)          AS ecart_marge_pct
            FROM bud b FULL OUTER JOIN act a ON a.mois = b.mois
            ORDER BY mois
            """
        ).df()

        # --- budget_by_region : réel vs budget par région ---
        budget_by_region = con.execute(
            f"""
            WITH act AS (
                SELECT region, sum(ca) ca_reel, sum(ca - cout) marge_reel
                FROM _base GROUP BY region
            ),
            bud AS (
                SELECT r.nom_region AS region,
                       sum(b.ca_budgete)    AS ca_budget,
                       sum(b.marge_budgetee) AS marge_budget
                FROM fait_budget b JOIN dim_region r ON b.code_region = r.code_region
                WHERE 1 = 1{budget_period_filter}{budget_region_filter}
                GROUP BY r.nom_region
            )
            SELECT
                COALESCE(b.region, a.region)             AS region,
                COALESCE(a.ca_reel, 0)                   AS ca_reel,
                b.ca_budget                              AS ca_budget,
                COALESCE(a.ca_reel, 0) - b.ca_budget     AS ecart_ca,
                100.0 * (COALESCE(a.ca_reel, 0) - b.ca_budget)
                    / NULLIF(b.ca_budget, 0)             AS ecart_ca_pct,
                COALESCE(a.marge_reel, 0)                AS marge_reel,
                b.marge_budget                           AS marge_budget,
                COALESCE(a.marge_reel, 0) - b.marge_budget AS ecart_marge,
                100.0 * (COALESCE(a.marge_reel, 0) - b.marge_budget)
                    / NULLIF(b.marge_budget, 0)          AS ecart_marge_pct
            FROM bud b FULL OUTER JOIN act a ON a.region = b.region
            ORDER BY ca_reel DESC
            """
        ).df()

        # --- category_mix : part de CA par catégorie pour chaque année (effet mix) ---
        # Pivot année -> colonnes part_YYYY. Met en évidence le glissement du mix.
        category_mix = con.execute(
            """
            WITH parts AS (
                SELECT categorie, annee, sum(ca) AS ca,
                       100.0 * sum(ca) / sum(sum(ca)) OVER (PARTITION BY annee) AS part
                FROM _base GROUP BY categorie, annee
            )
            PIVOT parts ON 'part_' || annee USING first(part) GROUP BY categorie
            ORDER BY categorie
            """
        ).df()

        # --- by_salesperson : CA, marge et taux par commercial (créateur de valeur ?) ---
        by_salesperson = con.execute(
            """
            SELECT commercial,
                   sum(ca)                                    AS ca,
                   sum(ca - cout)                             AS marge,
                   100.0 * sum(ca - cout) / NULLIF(sum(ca), 0) AS taux_marge,
                   100.0 * sum(ca) / NULLIF(sum(sum(ca)) OVER (), 0) AS part_ca_pct
            FROM _base
            GROUP BY commercial ORDER BY ca DESC
            """
        ).df()

        # --- by_segment : poids et rentabilité par segment client (concentration) ---
        by_segment = con.execute(
            """
            SELECT segment,
                   sum(ca)                                    AS ca,
                   sum(ca - cout)                             AS marge,
                   100.0 * sum(ca - cout) / NULLIF(sum(ca), 0) AS taux_marge,
                   count(DISTINCT code_client)                AS nb_clients,
                   100.0 * sum(ca) / NULLIF(sum(sum(ca)) OVER (), 0) AS part_ca_pct
            FROM _base
            GROUP BY segment ORDER BY ca DESC
            """
        ).df()

        # --- client_pareto : CA cumulé par client (courbe de concentration ABC) ---
        client_pareto = con.execute(
            """
            WITH per_client AS (
                SELECT code_client, sum(ca) AS ca FROM _base GROUP BY code_client
            )
            SELECT
                row_number() OVER (ORDER BY ca DESC)              AS rang,
                ca,
                100.0 * sum(ca) OVER (ORDER BY ca DESC
                    ROWS UNBOUNDED PRECEDING)
                    / sum(ca) OVER ()                             AS ca_cumule_pct
            FROM per_client ORDER BY ca DESC
            """
        ).df()

        # --- client_top : top 15 clients (CA, part, marge) ---
        client_top = con.execute(
            """
            SELECT raison_sociale, segment,
                   sum(ca)                                    AS ca,
                   100.0 * sum(ca) / NULLIF(sum(sum(ca)) OVER (), 0) AS part_ca_pct,
                   100.0 * sum(ca - cout) / NULLIF(sum(ca), 0) AS taux_marge
            FROM _base
            GROUP BY raison_sociale, segment
            ORDER BY ca DESC LIMIT 15
            """
        ).df()

        # --- by_discount_band : impact des remises sur la marge (fuite de marge) ---
        # Tranches choisies sur la distribution réelle (cluster 0-8 % + cluster >20 %).
        by_discount_band = con.execute(
            """
            SELECT tranche, lignes, ca, taux_marge, part_ca_pct
            FROM (
                SELECT
                    CASE WHEN remise = 0      THEN '0 %'
                         WHEN remise <= 0.05  THEN '1-5 %'
                         WHEN remise <= 0.10  THEN '6-10 %'
                         ELSE '> 10 %' END                  AS tranche,
                    min(remise)                             AS ord,
                    count(*)                                AS lignes,
                    sum(ca)                                 AS ca,
                    100.0 * sum(ca - cout) / NULLIF(sum(ca), 0)        AS taux_marge,
                    100.0 * sum(ca) / NULLIF(sum(sum(ca)) OVER (), 0)  AS part_ca_pct
                FROM _base GROUP BY tranche
            ) ORDER BY ord
            """
        ).df()
    finally:
        con.close()

    # --- summary (1 ligne) : agrège les totaux et dérive le YoY depuis by_year ---
    ca_total = float(monthly["ca"].sum())
    marge_total = float(monthly["marge"].sum())
    ca_budget_total = float(budget_monthly["ca_budget"].sum())
    marge_budget_total = float(budget_monthly["marge_budget"].sum())

    # YoY : défini seulement si l'on dispose d'au moins deux années pleines.
    ca_yoy_pct: Optional[float] = None
    tm_delta_pts: Optional[float] = None
    if len(by_year) >= 2:
        first, last = by_year.iloc[0], by_year.iloc[-1]
        if first["ca"]:
            ca_yoy_pct = 100.0 * (last["ca"] - first["ca"]) / first["ca"]
        tm_delta_pts = float(last["taux_marge"] - first["taux_marge"])

    summary = pd.DataFrame(
        [
            {
                "ca": ca_total,
                "marge": marge_total,
                "taux_marge": 100.0 * marge_total / ca_total if ca_total else None,
                "ca_budget": ca_budget_total,
                "marge_budget": marge_budget_total,
                "ecart_ca": ca_total - ca_budget_total,
                "ecart_ca_pct": (
                    100.0 * (ca_total - ca_budget_total) / ca_budget_total
                    if ca_budget_total
                    else None
                ),
                "ecart_marge": marge_total - marge_budget_total,
                "ecart_marge_pct": (
                    100.0 * (marge_total - marge_budget_total) / marge_budget_total
                    if marge_budget_total
                    else None
                ),
                "ca_yoy_pct": ca_yoy_pct,
                "taux_marge_delta_pts": tm_delta_pts,
                "clients_actifs": int(by_year["clients_actifs"].max())
                if not by_year.empty
                else 0,
            }
        ]
    )

    logging.info(
        "[vues analytiques] CA %.0f € | marge %.0f € (%.1f%%) | écart budget CA %.1f%%",
        ca_total,
        marge_total,
        summary.loc[0, "taux_marge"] if ca_total else 0.0,
        summary.loc[0, "ecart_ca_pct"] if ca_budget_total else 0.0,
    )
    return {
        "summary": summary,
        "by_year": by_year,
        "monthly": monthly,
        "by_category": by_category,
        "budget_monthly": budget_monthly,
        "budget_by_region": budget_by_region,
        "category_mix": category_mix,
        "by_salesperson": by_salesperson,
        "by_segment": by_segment,
        "client_pareto": client_pareto,
        "client_top": client_top,
        "by_discount_band": by_discount_band,
    }