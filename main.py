"""
main.py - Point d'entrée pour l'exécution en mode Batch (automatisé).

Usage direct :
    python main.py

Usage planifié (Windows - Planificateur de tâches) :
    Lancer RUN_BATCH.bat selon la fréquence souhaitée.

Ce script est intentionnellement minimal : il orchestre les appels à utils.py
sans contenir de logique métier. Toute la logique est dans utils.py.
Les chemins et constantes proviennent de config.SETTINGS (surchargeable via .env).

Formats sources supportés : .xlsx et .csv (détectés automatiquement).

Stratégie anti-coupure :
    Tous les fichiers de sortie sont d'abord écrits avec le préfixe ``CORROMPU_``
    dans leur nom. À la toute fin du run (après succès complet), ils sont renommés
    vers leur nom définitif. Si le process est interrompu (coupure secteur, plantage,
    Ctrl+C), les fichiers incomplets restent préfixés ``CORROMPU_`` sur le disque et
    ne peuvent pas être diffusés par erreur comme s'ils étaient valides.
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

import utils
from config import SETTINGS
from workbook import export_excel_with_dashboard


def main() -> None:
    """Exécute le pipeline complet de reporting en mode Batch.

    Pipeline :
        1. Calcul de l'horodatage du run (``run_ts = HHhMM``).
        2. Initialisation du logging (fichier horodaté ``log_YYYY-MM-DD_HHhMM.txt``).
        3. Chargement robuste des fichiers sources (.xlsx et .csv depuis data/).
        4. Nettoyage et enrichissement (colonne mois).
        5. Checkpoint de réconciliation des données (audit).
        6. Calcul des reportings (par mois, par commercial).
        7. Construction des chemins temporaires (``CORROMPU_*.xlsx``, ``CORROMPU_*.pdf``).
        8. Export Excel multi-feuilles → chemin temporaire.
        9. Export PDF → chemin temporaire.
       10. Commit atomique : renommage des fichiers temporaires vers leur nom définitif.

    **Stratégie anti-coupure :**
    Les fichiers sont écrits sous leur nom temporaire (préfixe ``CORROMPU_``).
    Si le process est interrompu à n'importe quel moment avant l'étape 10,
    les fichiers restent sur le disque avec le préfixe ``CORROMPU_`` et signalent
    visuellement leur incomplétude. Le renommage (étape 10) est atomique :
    soit il réussit entièrement, soit le fichier reste préfixé.

    En cas d'alerte d'intégrité (écart des données anormal), le script
    loggue une erreur critique et quitte avec le code 2 pour permettre
    une détection par un orchestrateur ou le Planificateur de tâches.

    Returns:
        None. Effets de bord : écriture de fichiers dans output/<annee>/.

    Raises:
        SystemExit(0): Si aucun fichier n'est trouvé (pas une erreur).
        SystemExit(2): Si le checkpoint de réconciliation échoue.
    """
    # -----------------------------------------------------------------------
    # 1. Horodatage du run — calculé UNE SEULE FOIS et partagé avec tous
    #    les modules (log, Excel, PDF). Garantit la cohérence des noms.
    # -----------------------------------------------------------------------
    run_ts = datetime.now().strftime("%Hh%M")

    # 2. Logging — le fichier log inclut l'heure pour distinguer plusieurs
    #    runs déclenchés le même jour (planificateur, relances manuelles).
    log_path = utils.setup_logging(base_dir=str(SETTINGS.output_dir), run_ts=run_ts)
    logging.info(f"=== DÉMARRAGE - {SETTINGS.app_name} v{SETTINGS.app_version} ===")
    logging.info(f"Run : {run_ts} | Log : {log_path}")
    logging.info(f"Répertoire source : {SETTINGS.data_dir}")
    logging.info(f"Répertoire sortie : {SETTINGS.output_dir}")
    logging.info(
        f"Formats sources acceptés : {sorted(SETTINGS.csv_source_extensions)}"
    )
    logging.info(
        "Stratégie anti-coupure activée : fichiers écrits sous préfixe "
        f"'{utils._CORRUPT_PREFIX}' jusqu'au commit final."
    )

    # 3. Chargement robuste (.xlsx + .csv)
    # ValueError si aucun fichier lisible → sortie propre (code 0, pas une erreur fatale).
    try:
        df_source = utils.load_source_files(str(SETTINGS.data_dir))
    except ValueError as e:
        logging.warning(f"{e} → Fin sans traitement.")
        logging.info("=== FIN (AUCUN FICHIER) ===")
        sys.exit(0)

    source_row_count = len(df_source)
    logging.info(f"Lignes chargées (avant nettoyage) : {source_row_count}")

    # 4. Nettoyage + enrichissement
    # clean_data() est format-agnostique : fonctionne identiquement qu'il
    # provienne de fichiers .xlsx, .csv, ou d'un mix des deux.
    df = utils.clean_data(df_source)
    df = utils.add_month_column(df)

    # 5. Checkpoint de réconciliation des données
    recon_report = utils.reconcile_data(df_source, df)

    if not recon_report.integrity_ok:
        logging.critical(
            f"ALERTE INTÉGRITÉ - Écart de {recon_report.absolute_gap:.4f} € "
            f"({recon_report.gap_pct * 100:.2f} %) détecté. "
            "Le reporting NE DOIT PAS être diffusé avant investigation. "
            f"Somme attendue : {recon_report.expected_source_sum:.2f} € | "
            f"Somme traitée : {recon_report.processed_sum:.2f} €"
        )
        logging.info("=== FIN (ÉCHEC INTÉGRITÉ) ===")
        sys.exit(2)

    # 6. Reportings
    report_months = utils.aggregate_by_month(df)
    report_salespeople = utils.aggregate_by_salesperson(df)

    total_amount = float(df["montant"].sum())
    logging.info(f"Chiffre d'affaires total : {total_amount:,.2f} €")
    logging.info(f"Période : {df['mois'].min()} → {df['mois'].max()}")

    # -----------------------------------------------------------------------
    # 7. Construction des chemins — temporaires (CORROMPU_*) et définitifs.
    #    La racine commune ``base_stem`` est calculée une seule fois pour
    #    garantir la cohérence entre Excel et PDF (même période, même heure).
    # -----------------------------------------------------------------------
    min_month = df["mois"].min()
    max_month = df["mois"].max()
    year = str(min_month).split("-")[0]

    base_stem = (
        f"reporting_{min_month}"
        if min_month == max_month
        else f"reporting_{min_month}_to_{max_month}"
    )

    out_dir = SETTINGS.output_dir / year
    out_dir.mkdir(parents=True, exist_ok=True)

    excel_temp, excel_final = utils.build_output_paths(base_stem, out_dir, run_ts, ".xlsx")
    pdf_temp,   pdf_final   = utils.build_output_paths(base_stem, out_dir, run_ts, ".pdf")

    logging.info(f"Excel → écriture dans : {excel_temp.name}")
    logging.info(f"PDF   → écriture dans : {pdf_temp.name}")

    # 8. Export Excel → chemin temporaire
    export_excel_with_dashboard(
        df, report_months, report_salespeople, recon_report, excel_temp
    )

    # 9. Export PDF → chemin temporaire
    utils.export_pdf(report_months, report_salespeople, pdf_temp)

    # -----------------------------------------------------------------------
    # 10. Commit atomique — renommage vers les noms définitifs.
    #     Cette étape ne survient que si TOUTES les écritures ont réussi.
    #     Chaque appel à commit_output_file() est atomique (un seul syscall rename).
    #     En cas d'erreur ici (disque plein, droits insuffisants), les fichiers
    #     CORROMPU_* restent sur le disque comme signal d'alerte.
    # -----------------------------------------------------------------------
    logging.info("--- Finalisation des fichiers (commit atomique) ---")
    utils.commit_output_file(excel_temp, excel_final)
    utils.commit_output_file(pdf_temp,   pdf_final)

    logging.info(f"Fichiers disponibles dans : {out_dir}")
    logging.info("=== FIN OK ===")


if __name__ == "__main__":
    main()