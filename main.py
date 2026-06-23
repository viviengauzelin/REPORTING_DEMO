"""
main.py - Point d'entrée pour l'exécution en mode Batch (automatisé).

Usage direct :
    python main.py

Usage planifié (Windows - Planificateur de tâches) :
    Lancer RUN_BATCH.bat selon la fréquence souhaitée.

Ce script est intentionnellement minimal : il orchestre les appels à utils.py et
workbook.py sans contenir de logique métier. Toute la logique d'ingestion, de
consolidation et de réconciliation est dans utils.py ; l'assemblage du classeur
Excel est dans workbook.py. Les chemins et constantes proviennent de
config.SETTINGS (surchargeable via .env).

Pipeline (modèle en étoile) :
    sources (5 fichiers)  ->  run_ingestion  ->  star schema
    star schema           ->  reconcile_against_manifest  (checkpoint d'audit)
    star schema           ->  build_reporting_views  ->  classeur Excel enrichi

Le CA n'est jamais stocké : il est recalculé en SQL sur les faits, à chaque étape
(réconciliation et vues de reporting). Le livrable est un classeur Excel unique
(le PDF a été retiré : il faisait doublon avec le dashboard Excel, plus riche).

Stratégie anti-coupure :
    Le fichier de sortie est d'abord écrit avec le préfixe ``CORROMPU_`` dans son
    nom. À la toute fin du run (après succès complet), il est renommé vers son nom
    définitif. Si le process est interrompu (coupure secteur, plantage, Ctrl+C), le
    fichier incomplet reste préfixé ``CORROMPU_`` sur le disque et ne peut pas être
    diffusé par erreur comme s'il était valide.
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime

import utils
from config import SETTINGS
from workbook import export_excel_with_dashboard


def main() -> None:
    """Exécute le pipeline complet de reporting en mode Batch.

    Pipeline :
        1. Horodatage du run (``run_ts = HHhMM``), calculé une seule fois.
        2. Initialisation du logging (fichier horodaté ``log_AAAA-MM-JJ_HHhMM.txt``).
        3. Ingestion des 5 sources et consolidation en modèle en étoile
           (``run_ingestion``).
        4. Chargement du manifeste-oracle (``load_manifest``).
        5. Checkpoint de réconciliation contre l'oracle
           (``reconcile_against_manifest``) : 5 métriques + intégrité référentielle.
        6. Dérivation des vues de reporting depuis le star schema
           (``build_reporting_views``) — CA recalculé en SQL, jamais stocké.
        7. Construction du chemin temporaire (``CORROMPU_*.xlsx``) et définitif.
        8. Export du classeur Excel enrichi -> chemin temporaire.
        9. Commit atomique : renommage du fichier temporaire vers son nom définitif.

    Stratégie anti-coupure : le fichier est écrit sous son nom temporaire
    (préfixe ``CORROMPU_``). Si le process est interrompu avant l'étape 9, il reste
    sur le disque préfixé et signale visuellement son incomplétude. Le renommage
    (étape 9) est atomique : soit il réussit, soit le fichier reste préfixé.

    Returns:
        None. Effet de bord : écriture du classeur dans ``output/<annee>/``.

    Raises:
        SystemExit(0): Si aucune source n'est trouvée (pas une erreur fatale).
        SystemExit(2): Si le manifeste-oracle est absent (audit impossible) ou si le
            checkpoint de réconciliation échoue (intégrité non garantie).
    """
    # -----------------------------------------------------------------------
    # 1. Horodatage du run — calculé UNE SEULE FOIS et partagé avec tous les
    #    modules (log, Excel). Garantit la cohérence des noms de fichiers.
    # -----------------------------------------------------------------------
    run_ts = datetime.now().strftime("%Hh%M")

    # 2. Logging — le fichier log inclut l'heure pour distinguer plusieurs runs
    #    déclenchés le même jour (planificateur, relances manuelles).
    log_path = utils.setup_logging(base_dir=str(SETTINGS.output_dir), run_ts=run_ts)
    logging.info("=== DÉMARRAGE - %s v%s ===", SETTINGS.app_name, SETTINGS.app_version)
    logging.info("Run : %s | Log : %s", run_ts, log_path)
    logging.info("Répertoire source : %s", SETTINGS.raw_dir)
    logging.info("Répertoire sortie : %s", SETTINGS.output_dir)
    logging.info(
        "Stratégie anti-coupure activée : fichier écrit sous préfixe "
        "'%s' jusqu'au commit final.",
        utils._CORRUPT_PREFIX,
    )

    # -----------------------------------------------------------------------
    # 3. Ingestion + consolidation -> star schema.
    #    FileNotFoundError si aucune source : sortie propre (code 0, pas fatal).
    # -----------------------------------------------------------------------
    try:
        star, reports = utils.run_ingestion(SETTINGS.raw_dir)
    except FileNotFoundError as e:
        logging.warning("%s → Fin sans traitement.", e)
        logging.info("=== FIN (AUCUNE SOURCE) ===")
        sys.exit(0)

    # -----------------------------------------------------------------------
    # 4. Manifeste-oracle. Son absence est bloquante : sans oracle, on ne peut
    #    rien prouver -> on refuse de produire un livrable non audité (code 2).
    # -----------------------------------------------------------------------
    try:
        manifest = utils.load_manifest(SETTINGS.raw_dir)
    except FileNotFoundError as e:
        logging.critical("Manifeste-oracle absent : audit impossible. %s", e)
        logging.info("=== FIN (MANIFESTE ABSENT) ===")
        sys.exit(2)

    # -----------------------------------------------------------------------
    # 5. Checkpoint de réconciliation contre l'oracle.
    #    reconcile_against_manifest journalise déjà le rapport détaillé ;
    #    en cas d'échec, on bloque la diffusion (code 2).
    # -----------------------------------------------------------------------
    recon = utils.reconcile_against_manifest(star, reports, manifest)
    if not recon.integrity_ok:
        n_ko = sum(1 for c in recon.checks if not c.ok)
        failed = ", ".join(c.label for c in recon.checks if not c.ok)
        logging.critical(
            f"ALERTE INTÉGRITÉ - {n_ko} contrôle(s) en échec ({failed}). "
            "Le reporting NE DOIT PAS être diffusé avant investigation."
        )
        logging.info("=== FIN (ÉCHEC INTÉGRITÉ) ===")
        sys.exit(2)

    # -----------------------------------------------------------------------
    # 6. Vues de reporting dérivées du star schema (CA recalculé en SQL).
    # -----------------------------------------------------------------------
    df, report_months, report_salespeople = utils.build_reporting_views(star)
    analytics = utils.build_analytics_views(star)
    total_ca = float(df["montant"].sum())
    logging.info("Chiffre d'affaires livré : %s €", f"{total_ca:,.2f}")
    if not report_months.empty:
        logging.info(
            "Période : %s → %s",
            report_months["mois"].min(),
            report_months["mois"].max(),
        )

    # -----------------------------------------------------------------------
    # 7. Construction des chemins — temporaire (CORROMPU_*) et définitif.
    #    La période vient de l'agrégat mensuel (mois au format AAAA-MM, donc le
    #    tri lexicographique = tri chronologique). Repli sur un nommage générique
    #    si aucun mois n'est exploitable (toutes dates invalides).
    # -----------------------------------------------------------------------
    if not report_months.empty:
        min_month = str(report_months["mois"].min())
        max_month = str(report_months["mois"].max())
        year_min = min_month.split("-")[0]
        year_max = max_month.split("-")[0]
        year = f"{year_min}_{year_max}"
        base_stem = (
            f"reporting_{min_month}"
            if min_month == max_month
            else f"reporting_{min_month}_to_{max_month}"
        )
    else:
        year = datetime.now().strftime("%Y")
        base_stem = "reporting"
        logging.warning(
            "Aucun mois exploitable (dates toutes invalides ?) → nommage générique."
        )

    out_dir = SETTINGS.output_dir / year
    out_dir.mkdir(parents=True, exist_ok=True)

    excel_temp, excel_final = utils.build_output_paths(
        base_stem, out_dir, run_ts, ".xlsx"
    )
    logging.info("Excel → écriture dans : %s", excel_temp.name)

    # 8. Export du classeur Excel enrichi -> chemin temporaire.
    export_excel_with_dashboard(
        df, report_months, report_salespeople, recon, excel_temp, analytics=analytics
    )

    # -----------------------------------------------------------------------
    # 9. Commit atomique — renommage vers le nom définitif. Ne survient que si
    #    l'écriture a réussi. En cas d'erreur ici (disque plein, droits), le
    #    fichier CORROMPU_* reste sur le disque comme signal d'alerte.
    # -----------------------------------------------------------------------
    logging.info("--- Finalisation du fichier (commit atomique) ---")
    utils.commit_output_file(excel_temp, excel_final)

    logging.info("Fichier disponible dans : %s", out_dir)
    logging.info("=== FIN OK ===")


if __name__ == "__main__":
    main()