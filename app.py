"""
app.py - Interface utilisateur Streamlit du Projet 1 (ingestion & consolidation CHR).

Ce module gère UNIQUEMENT la couche de présentation. Il ne contient aucune logique
métier : l'ingestion des 5 sources, la consolidation en modèle en étoile, la
réconciliation contre l'oracle et la dérivation des vues analytiques sont déléguées
à ``utils.py`` ; l'assemblage du classeur Excel est délégué à ``workbook.py`` ;
toute la configuration vit dans ``config.py``.

Flux utilisateur :
    1. L'application lit les 5 sources « sales » depuis le répertoire ``raw_dir``
       (configurable via ``.env``). Si elles sont absentes, un bouton génère un jeu
       de données de démonstration réaliste.
    2. Un clic sur « Lancer le pipeline » exécute l'ingestion, la réconciliation et
       construit le classeur Excel (une seule fois, mémorisé pour la session) — le
       tout sur **100 % des sources**. Le journal d'exécution est capturé en mémoire.
    3. L'écran affiche le rapport de réconciliation, puis une vue analytique DAF
       organisée en onglets. Des **filtres de présentation** (période, région,
       segment, commercial) re-dérivent à la volée les vues affichées, sans jamais
       toucher à l'ingestion ni à la réconciliation (l'audit reste sur 100 %).

Le CA n'est jamais stocké : il est recalculé en SQL sur les faits du star schema.
"""

from __future__ import annotations

import io
import logging
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterable, Iterator, Optional, cast

import pandas as pd
import streamlit as st

import utils
from config import SETTINGS
from workbook import build_excel_bytes_with_dashboard

# La couche métier journalise abondamment ; en dehors de la capture de journal
# (voir ``_capture_logs``), on ne garde que les avertissements à l'écran.
logging.getLogger().setLevel(logging.WARNING)

st.set_page_config(page_title="Reporting CHR — Projet 1", layout="wide")

_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------------------
# CAPTURE DU JOURNAL D'EXÉCUTION (audit)
# ---------------------------------------------------------------------------


@contextmanager
def _capture_logs() -> Iterator[io.StringIO]:
    """Capture en mémoire toutes les traces ``INFO`` émises pendant le bloc.

    En mode Streamlit aucun fichier log n'est écrit sur disque (contrairement au
    mode Batch qui appelle ``utils.setup_logging``). On attache donc temporairement
    un handler vers un tampon mémoire : on récupère ainsi le journal complet du run
    — empreintes SHA-256 des sources comprises (émises par ``load_ventes``) — pour
    le proposer au téléchargement, sans effet de bord durable sur le logging.
    """
    buffer = io.StringIO()
    handler = logging.StreamHandler(buffer)
    handler.setLevel(logging.INFO)
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    root = logging.getLogger()
    previous_level = root.level
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    try:
        yield buffer
    finally:
        root.removeHandler(handler)
        root.setLevel(previous_level)


# ---------------------------------------------------------------------------
# HELPERS DE FORMATAGE (présentation uniquement)
# ---------------------------------------------------------------------------


def _format_currency_fr(x: float | None) -> str:
    """Formate un float en montant français lisible : ``1234567.89`` → ``1 234 567,89 €``."""
    if x is None or pd.isna(x):
        return "—"
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"{s} €"


def _format_int_fr(x: float | None) -> str:
    """Formate un entier avec séparateur de milliers français : ``57219`` → ``57 219``."""
    if x is None or pd.isna(x):
        return "—"
    return f"{int(x):,}".replace(",", " ")


def _format_pct(x: Optional[float], *, signed: bool = False) -> str:
    """Formate un pourcentage ; ``None``/``NaN`` → ``« — »`` (valeur non calculable).

    Args:
        x: Valeur en points de pourcentage (ex. ``16.9``), ou ``None``.
        signed: Si vrai, affiche le signe (``+16,9 %``) — utile pour les écarts/croissances.
    """
    if x is None or pd.isna(x):
        return "—"
    pattern = "{:+.1f} %" if signed else "{:.1f} %"
    return pattern.format(x).replace(".", ",")


def _scalar(value: object) -> Optional[float]:
    """Renvoie un float exploitable, ou ``None`` si la valeur est nulle/NaN."""
    if value is None or pd.isna(value):
        return None
    return float(cast(float, value))


# ---------------------------------------------------------------------------
# ÉTAT DES SOURCES & DÉMO
# ---------------------------------------------------------------------------


def _sources_status(raw_dir: Path) -> dict[str, object]:
    """Évalue la présence des sources de ventes et de l'oracle dans le répertoire."""
    ventes = sorted(raw_dir.glob("export_ventes_*.xlsx")) if raw_dir.exists() else []
    manifest_present = (raw_dir / SETTINGS.manifest_name).exists()
    return {
        "ventes_files": len(ventes),
        "manifest": manifest_present,
        "ready": len(ventes) > 0 and manifest_present,
    }


def _generate_demo(raw_dir: Path) -> None:
    """Génère un jeu de données de démonstration dans ``raw_dir`` (+ manifeste-oracle).

    L'import est local : le générateur est lourd (numpy, openpyxl) et n'a aucun
    intérêt à être chargé tant que l'utilisateur ne demande pas explicitement la démo.
    """
    from generate_demo_data import generate_dataset

    raw_dir.mkdir(parents=True, exist_ok=True)
    generate_dataset(raw_dir, raw_dir / SETTINGS.manifest_name)


def _hash_sources(raw_dir: Path) -> list[tuple[str, str]]:
    """Calcule l'empreinte SHA-256 de chaque fichier source (transparence d'audit).

    On hache l'intégralité des fichiers présents dans ``raw_dir`` (sources + oracle) :
    ces empreintes matérialisent à l'écran le standard de traçabilité du pipeline.
    """
    return [
        (p.name, utils.hash_file_on_disk(p))
        for p in sorted(raw_dir.iterdir())
        if p.is_file()
    ]


# ---------------------------------------------------------------------------
# RÉCONCILIATION (inchangé : audit sur 100 %)
# ---------------------------------------------------------------------------


def _checks_to_frame(recon: utils.ReconciliationReport) -> pd.DataFrame:
    """Transforme les points de contrôle en DataFrame lisible pour l'affichage."""

    def _fmt(label: str, value: float) -> str:
        if "€" in label:
            return _format_currency_fr(float(value))
        return _format_int_fr(int(round(float(value))))

    return pd.DataFrame(
        [
            {
                "Contrôle": c.label,
                "Obtenu": _fmt(c.label, c.obtained),
                "Attendu": _fmt(c.label, c.expected),
                "Statut": "✅ OK" if c.ok else "🚨 ALERTE",
                "Détail": c.detail or "—",
            }
            for c in recon.checks
        ]
    )


def _highlight_alert(row: pd.Series) -> list[str]:
    """Surligne en jaune pâle (charte ALERTE) toute ligne de contrôle en échec."""
    alert = "ALERTE" in str(row["Statut"])
    return ["background-color: #FFF2CC" if alert else "" for _ in row]


def _render_reconciliation(recon: utils.ReconciliationReport) -> None:
    """Affiche le rapport de réconciliation : bandeau global + tableau des contrôles."""
    n_ok = sum(1 for c in recon.checks if c.ok)
    n_total = len(recon.checks)
    st.caption(
        "La réconciliation porte sur **100 % des sources**, indépendamment des "
        "filtres d'analyse ci-contre."
    )
    if recon.integrity_ok:
        st.success(f"Intégrité validée — {n_ok}/{n_total} contrôles au vert.")
    else:
        st.error(
            f"Alerte intégrité — {n_ok}/{n_total} contrôles au vert. "
            "Le reporting ne doit pas être diffusé avant investigation."
        )
    styled = _checks_to_frame(recon).style.apply(_highlight_alert, axis=1)
    st.dataframe(styled, hide_index=True, width="stretch")


# ---------------------------------------------------------------------------
# FILTRES DE PRÉSENTATION
# ---------------------------------------------------------------------------


def _options_or_none(
    selection: Iterable[str], universe: list[str]
) -> Optional[list[str]]:
    """Convertit une sélection de widget en argument de filtre.

    Convention : une sélection vide (ou couvrant tout l'univers) signifie « aucun
    filtre » et renvoie ``None`` — ce qui déclenche le chemin rétrocompatible de
    ``build_analytics_views`` (vue sur l'intégralité du périmètre).
    """
    selected = list(selection)
    if not selected or len(selected) == len(universe):
        return None
    return selected


def _render_filters(options: dict[str, list[str]]) -> dict[str, Optional[list[str]]]:
    """Affiche les filtres dans la barre latérale et renvoie les kwargs d'analyse."""
    st.sidebar.header("🔎 Filtres d'analyse")
    st.sidebar.caption(
        "Ces filtres n'affectent que l'affichage. L'ingestion, la réconciliation et "
        "le classeur Excel portent toujours sur 100 % des sources."
    )

    months = options["months"]
    if len(months) >= 2:
        start, end = st.sidebar.select_slider(
            "Période (mois)", options=months, value=(months[0], months[-1])
        )
        months_sel: list[str] = [m for m in months if start <= m <= end]
    else:
        months_sel = list(months)

    regions_sel = st.sidebar.multiselect(
        "Régions", options["regions"], help="Vide = toutes les régions."
    )
    segments_sel = st.sidebar.multiselect(
        "Segments clients", options["segments"], help="Vide = tous les segments."
    )
    salespeople_sel = st.sidebar.multiselect(
        "Commerciaux", options["salespeople"], help="Vide = tous les commerciaux."
    )

    return {
        "months": months_sel if len(months_sel) != len(months) else None,
        "regions": _options_or_none(regions_sel, options["regions"]),
        "segments": _options_or_none(segments_sel, options["segments"]),
        "salespeople": _options_or_none(salespeople_sel, options["salespeople"]),
    }


# ---------------------------------------------------------------------------
# ONGLETS ANALYTIQUES
# ---------------------------------------------------------------------------


def _render_summary(summary: pd.DataFrame) -> None:
    """🧭 Synthèse DAF : KPIs filtrables issus de la vue ``summary`` (1 ligne)."""
    s = summary.iloc[0]
    c1, c2, c3 = st.columns(3)
    c1.metric("Chiffre d'affaires livré", _format_currency_fr(_scalar(s["ca"])))
    c2.metric("Marge brute", _format_currency_fr(_scalar(s["marge"])))
    c3.metric("Taux de marge", _format_pct(_scalar(s["taux_marge"])))

    ecart_pct = _scalar(s["ecart_ca_pct"])
    c4, c5, c6 = st.columns(3)
    c4.metric(
        "Écart au budget (CA)",
        _format_currency_fr(_scalar(s["ecart_ca"])),
        delta=_format_pct(ecart_pct, signed=True) if ecart_pct is not None else None,
    )
    c5.metric("Croissance CA (YoY)", _format_pct(_scalar(s["ca_yoy_pct"]), signed=True))
    c6.metric("Clients actifs", _format_int_fr(_scalar(s["clients_actifs"]) or 0))


def _render_margin(
    monthly: pd.DataFrame,
    by_category: pd.DataFrame,
    by_discount_band: pd.DataFrame,
) -> None:
    """💹 Marge & rentabilité : érosion mensuelle, mix catégorie, fuite par remise."""
    if monthly.empty:
        st.info("Aucune donnée sur ce périmètre.")
        return

    st.caption("Chiffre d'affaires et taux de marge par mois")
    m = monthly.sort_values("mois").set_index("mois")
    left, right = st.columns(2)
    left.caption("CA mensuel (€)")
    left.bar_chart(m["ca"])
    right.caption("Taux de marge mensuel (%)")
    right.line_chart(m["taux_marge"])

    st.caption("Rentabilité par catégorie produit")
    cat = by_category.assign(
        CA=lambda d: d["ca"].map(_format_currency_fr),
        Marge=lambda d: d["marge"].map(_format_currency_fr),
        **{
            "Taux de marge": lambda d: d["taux_marge"].map(_format_pct),
            "Part du CA": lambda d: d["part_ca_pct"].map(_format_pct),
        },
    ).rename(columns={"categorie": "Catégorie"})
    st.dataframe(
        cat[["Catégorie", "CA", "Marge", "Taux de marge", "Part du CA"]],
        hide_index=True,
        width="stretch",
    )

    st.caption("Impact des remises sur la marge (fuite de marge)")
    band = by_discount_band.set_index("tranche")
    st.bar_chart(band["taux_marge"])


def _render_budget(
    analytics: dict[str, pd.DataFrame], off_axis: bool
) -> None:
    """🎯 Réalisé vs Budget : comparaison par mois et région (axes du budget)."""
    if off_axis:
        st.info(
            "Le budget n'est dimensionné que par **mois** et **région**. Retire les "
            "filtres « Segments » et « Commerciaux » pour une comparaison au budget "
            "fiable (sinon l'écart opposerait un réel filtré à un budget complet)."
        )
        return

    bm = analytics["budget_monthly"]
    if bm.empty:
        st.info("Aucune donnée budgétaire sur ce périmètre.")
        return

    st.caption("Réalisé vs budget — chiffre d'affaires mensuel (€)")
    chart_df = bm.set_index("mois")[["ca_reel", "ca_budget"]].rename(
        columns={"ca_reel": "Réalisé", "ca_budget": "Budget"}
    )
    st.bar_chart(chart_df)

    st.caption("Écarts au budget par mois (écart = réalisé − budget)")
    table = bm.assign(
        Mois=lambda d: d["mois"],
        Réalisé=lambda d: d["ca_reel"].map(_format_currency_fr),
        Budget=lambda d: d["ca_budget"].map(_format_currency_fr),
        Écart=lambda d: d["ecart_ca"].map(_format_currency_fr),
        **{"Écart %": lambda d: d["ecart_ca_pct"].map(lambda v: _format_pct(v, signed=True))},
    )
    st.dataframe(
        table[["Mois", "Réalisé", "Budget", "Écart", "Écart %"]],
        hide_index=True,
        width="stretch",
    )

    st.caption("Réalisé vs budget par région")
    region = analytics["budget_by_region"].assign(
        Région=lambda d: d["region"],
        Réalisé=lambda d: d["ca_reel"].map(_format_currency_fr),
        Budget=lambda d: d["ca_budget"].map(_format_currency_fr),
        Écart=lambda d: d["ecart_ca"].map(_format_currency_fr),
        **{"Écart %": lambda d: d["ecart_ca_pct"].map(lambda v: _format_pct(v, signed=True))},
    )
    st.dataframe(
        region[["Région", "Réalisé", "Budget", "Écart", "Écart %"]],
        hide_index=True,
        width="stretch",
    )


def _render_clients(
    client_pareto: pd.DataFrame,
    client_top: pd.DataFrame,
    by_segment: pd.DataFrame,
) -> None:
    """👥 Clients & segments : concentration (Pareto), top clients, poids des segments."""
    if client_pareto.empty:
        st.info("Aucun client sur ce périmètre.")
        return

    st.caption("Concentration du CA clients (courbe ABC : % de CA cumulé par rang)")
    st.line_chart(client_pareto.set_index("rang")["ca_cumule_pct"])

    st.caption("Top 15 clients")
    top = client_top.assign(
        Client=lambda d: d["raison_sociale"],
        Segment=lambda d: d["segment"],
        CA=lambda d: d["ca"].map(_format_currency_fr),
        **{
            "Part du CA": lambda d: d["part_ca_pct"].map(_format_pct),
            "Taux de marge": lambda d: d["taux_marge"].map(_format_pct),
        },
    )
    st.dataframe(
        top[["Client", "Segment", "CA", "Part du CA", "Taux de marge"]],
        hide_index=True,
        width="stretch",
    )

    st.caption("Poids et rentabilité par segment")
    seg = by_segment.assign(
        Segment=lambda d: d["segment"],
        CA=lambda d: d["ca"].map(_format_currency_fr),
        Clients=lambda d: d["nb_clients"].map(_format_int_fr),
        **{
            "Taux de marge": lambda d: d["taux_marge"].map(_format_pct),
            "Part du CA": lambda d: d["part_ca_pct"].map(_format_pct),
        },
    )
    st.dataframe(
        seg[["Segment", "CA", "Clients", "Taux de marge", "Part du CA"]],
        hide_index=True,
        width="stretch",
    )


def _render_salespeople(by_salesperson: pd.DataFrame) -> None:
    """👤 Par commercial : contribution au CA et rentabilité."""
    if by_salesperson.empty:
        st.info("Aucun commercial sur ce périmètre.")
        return

    st.caption("Chiffre d'affaires par commercial (€)")
    st.bar_chart(by_salesperson.set_index("commercial")["ca"])

    sp = by_salesperson.assign(
        Commercial=lambda d: d["commercial"],
        CA=lambda d: d["ca"].map(_format_currency_fr),
        Marge=lambda d: d["marge"].map(_format_currency_fr),
        **{
            "Taux de marge": lambda d: d["taux_marge"].map(_format_pct),
            "Part du CA": lambda d: d["part_ca_pct"].map(_format_pct),
        },
    )
    st.dataframe(
        sp[["Commercial", "CA", "Marge", "Taux de marge", "Part du CA"]],
        hide_index=True,
        width="stretch",
    )


# ---------------------------------------------------------------------------
# EXPORT & AUDIT
# ---------------------------------------------------------------------------


def _render_export(
    excel_bytes: bytes,
    run_log: str,
    source_hashes: list[tuple[str, str]],
    integrity_ok: bool,
) -> None:
    """Propose les téléchargements (Excel 100 % + journal) et les empreintes sources."""
    if not integrity_ok:
        st.caption(
            "⚠️ Le classeur contient une alerte d'intégrité (feuille Réconciliation). "
            "À ne pas diffuser tel quel."
        )
    stamp = f"{datetime.now():%Y-%m-%d_%Hh%M}"
    col1, col2 = st.columns(2)
    col1.download_button(
        "📊 Classeur Excel enrichi (sur 100 % des sources)",
        data=excel_bytes,
        file_name=f"reporting_chr_{stamp}.xlsx",
        mime=_XLSX_MIME,
        type="primary",
    )
    col2.download_button(
        "📝 Journal d'exécution (.txt)",
        data=run_log.encode("utf-8"),
        file_name=f"journal_execution_{stamp}.txt",
        mime="text/plain",
    )
    with st.expander("🔐 Empreintes SHA-256 des fichiers sources"):
        st.caption(
            "Preuve d'intégrité des fichiers ingérés : toute modification d'un octet "
            "change l'empreinte. Ces hachages figurent aussi dans le journal."
        )
        st.dataframe(
            pd.DataFrame(source_hashes, columns=["Fichier", "SHA-256"]),
            hide_index=True,
            width="stretch",
        )


# ---------------------------------------------------------------------------
# PAGE
# ---------------------------------------------------------------------------

st.title("📊 Reporting CHR — Projet 1 (ingestion & consolidation)")
st.write(
    "Distributeur B2B pour cafés, hôtels et restaurants. L'application ingère "
    "5 sources hétérogènes (ventes, CRM, catalogue, commerciaux, budget), les "
    "consolide en modèle en étoile, **prouve l'intégrité** contre un oracle, puis "
    "produit une vue d'aide à la décision (DAF) et un classeur Excel enrichi."
)

raw_dir = SETTINGS.raw_dir
status = _sources_status(raw_dir)

st.subheader("📁 Données sources")
st.caption(f"Répertoire lu : `{raw_dir}`  (surchargeable via la variable `.env` RAW_DIR)")

if not status["ready"]:
    st.warning(
        "Sources incomplètes : "
        f"{status['ventes_files']} fichier(s) de ventes détecté(s), "
        f"manifeste {'présent' if status['manifest'] else 'absent'}. "
        "Génère un jeu de démonstration pour tester le pipeline."
    )
    if st.button("🧪 Générer les données de démonstration"):
        with st.spinner("Génération des 5 sources + manifeste-oracle…"):
            _generate_demo(raw_dir)
        st.success("Données de démonstration générées.")
        st.rerun()
else:
    st.success(f"{status['ventes_files']} fichiers de ventes + manifeste-oracle détectés.")
    if st.button("▶️ Lancer le pipeline", type="primary"):
        with st.spinner(
            "Ingestion, consolidation, réconciliation et construction du classeur…"
        ):
            with _capture_logs() as logbuf:
                star, reports = utils.run_ingestion(raw_dir)
                manifest = utils.load_manifest(raw_dir)
                recon = utils.reconcile_against_manifest(star, reports, manifest)
                report_df, report_months, report_salespeople = utils.build_reporting_views(
                    star
                )
                # Vues analytiques + classeur : sur 100 % des données. Le classeur
                # (graphiques matplotlib) n'est construit qu'ICI, une seule fois ;
                # les filtres d'affichage re-dérivent ensuite des vues à coût quasi nul.
                analytics_full = utils.build_analytics_views(star)
                excel_bytes = build_excel_bytes_with_dashboard(
                    report_df,
                    report_months,
                    report_salespeople,
                    recon,
                    analytics=analytics_full,
                )
            source_hashes = _hash_sources(raw_dir)

        st.session_state["results"] = {
            "star": star,
            "recon": recon,
            "excel_bytes": excel_bytes,
            "run_log": logbuf.getvalue(),
            "source_hashes": source_hashes,
            "options": {
                "months": analytics_full["monthly"]["mois"].tolist(),
                "regions": sorted(star["dim_region"]["nom_region"].dropna().tolist()),
                "segments": sorted(
                    star["dim_client"]["segment"].dropna().unique().tolist()
                ),
                "salespeople": sorted(star["dim_commercial"]["nom"].dropna().tolist()),
            },
        }

# --- Affichage des résultats (persistés en session) ---
if "results" in st.session_state:
    res = st.session_state["results"]
    star = res["star"]

    selections = _render_filters(res["options"])
    # Le budget n'a pas d'axe segment/commercial : on le signale à l'onglet dédié.
    off_axis = selections["segments"] is not None or selections["salespeople"] is not None
    # Re-dérivation analytique filtrée (DuckDB, quelques ms) — aucune logique ici.
    analytics = utils.build_analytics_views(star, **selections)

    active = [
        label
        for label, value in [
            ("période", selections["months"]),
            ("région", selections["regions"]),
            ("segment", selections["segments"]),
            ("commercial", selections["salespeople"]),
        ]
        if value is not None
    ]
    st.divider()
    if active:
        st.info(f"Vue filtrée par : {', '.join(active)}. (L'audit reste sur 100 %.)")

    tab_synth, tab_marge, tab_budget, tab_clients, tab_com, tab_recon = st.tabs(
        [
            "🧭 Synthèse",
            "💹 Marge & rentabilité",
            "🎯 Réalisé vs Budget",
            "👥 Clients & segments",
            "👤 Par commercial",
            "🔐 Réconciliation",
        ]
    )
    with tab_synth:
        _render_summary(analytics["summary"])
    with tab_marge:
        _render_margin(
            analytics["monthly"], analytics["by_category"], analytics["by_discount_band"]
        )
    with tab_budget:
        _render_budget(analytics, off_axis)
    with tab_clients:
        _render_clients(
            analytics["client_pareto"], analytics["client_top"], analytics["by_segment"]
        )
    with tab_com:
        _render_salespeople(analytics["by_salesperson"])
    with tab_recon:
        _render_reconciliation(res["recon"])

    st.divider()
    st.subheader("⬇️ Export & audit")
    _render_export(
        res["excel_bytes"],
        res["run_log"],
        res["source_hashes"],
        res["recon"].integrity_ok,
    )