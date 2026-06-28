# Spécification — Générateur de données (dataset B2B CHR) — v1.1

> Document de référence du générateur, aligné sur le code livré.
> Convention de nommage : identifiants Python (fonctions, variables) en **anglais** ;
> en-têtes des fichiers sources et schéma métier en **français** (réalisme ERP/CRM
> français + cohérence avec `config.py`).

## Journal des modifications (v1.0 figée → v1.1 implémentée)

Quatre écarts assumés par rapport à la v1, tous dans l'esprit de la spec :

1. **Colonne `N° Ligne` ajoutée à l'export ventes.** Réaliste (un export ERP porte un
   n° de ligne) et surtout : donne la **clé de déduplication naturelle**
   `(N° Commande, N° Ligne)`. Sans elle, dédupliquer sur la ligne entière risquait
   de supprimer des lignes identiques par coïncidence.
2. **`ca_perdue_livree_eur` ajouté au manifeste** (bloc `ventes`). Rend l'oracle
   *exact* : le CA réconcilié attendu = CA vrai − CA des lignes manuelles livrées
   (montant détruit, donc irrécupérable).
3. **Précision de `remise_pct` = 3 décimales** (au lieu de 4), alignée sur ce que la
   colonne `Remise %` du fichier peut porter (pas de 0,1 %). Évite toute perte
   d'information sur la remise et garantit un bouclage au centime.
4. **Facteur budget recentré sur `U(0,93 ; 1,07)`** : les régions hors Sud oscillent
   autour de 0 % d'écart, le **Sud ressort comme seul point noir** (~−16 %).

Volumétrie réalisée : **~62 500 lignes** pour ~17 100 commandes (au-dessus du repère
indicatif de ~56 000, dans la marge du « ~ »). **29 fichiers sources + 1 manifeste.**

---

## 1. Objectif & philosophie

Générer un jeu de données **réaliste, reproductible et auto-vérifiable** simulant le système
d'information d'un **distributeur B2B pour la restauration (CHR)**. Le générateur produit :

- **5 types de sources « sales »** (29 fichiers : 24 ventes mensuelles + CRM + catalogue
  + 2 budgets + commerciaux), entrée du Projet 1,
- **1 manifeste de vérité** (`_manifest_anomalies.json`), oracle indépendant pour la
  réconciliation et les tests pytest.

Trois principes non négociables :

1. **Signaux ≠ anomalies.** Les *signaux métier* (glissement de mix, churn, sous-budget,
   sur-remise…) sont de vraies réalités économiques à **découvrir** ; ils ne sont jamais
   « nettoyés ». Les *anomalies qualité* (formats FR, doublons, dates KO…) sont des défauts
   à corriger et tracer.
2. **La saleté est majoritairement systématique**, pas aléatoire. Un export ERP est FR à 100 %
   parce que c'est la locale du poste ; seule une minorité de défauts est sporadique (champs
   saisis à la main, accidents d'export).
3. **L'oracle est calculé indépendamment du code de nettoyage.** La vérité financière du
   manifeste est calculée sur les valeurs **propres en mémoire, AVANT salissage** — jamais en
   nettoyant les données sales (sinon l'oracle est circulaire).

---

## 2. Paramètres globaux

| Paramètre | Valeur |
|---|---|
| Graine (`SEED`) | `42` (reproductibilité totale) |
| Période | `2024-01` → `2025-12` (24 mois, pour le YoY) |
| Clients | ~500 |
| Produits | ~120 |
| Commerciaux | 10 (2 par région) |
| Régions | 5 |
| Commandes | ~17 000 |
| Lignes de commande | ~62 500 réalisées (repère initial ~56 000 ; ≈ 3,7 lignes/commande) |
| Croissance YoY de base | 2025 ≈ 2024 × **1,08** (CA en hausse → YoY positif) |
| Répertoire de sortie | `data_raw/` (sources sales) + `_manifest_anomalies.json` à la racine |

Tout est paramétrable en tête de module (constantes). Le `SEED` garantit qu'un même paramétrage
produit exactement les mêmes fichiers et le même manifeste.

---

## 3. Étape 0 — Construire la vérité propre (en mémoire)

Le générateur construit d'abord **toutes les entités parfaitement propres** sous forme de
DataFrames pandas, dans l'ordre des dépendances : régions → commerciaux → produits → clients →
commandes → lignes. Aucune anomalie à ce stade.

### 3.1 Régions (5)

`code_region` (R01…R05), `nom_region` ∈ {Île-de-France, Nord, Est, **Sud**, Ouest}.
→ **Sud (R04)** porte le signal *sous-budget persistant* (cf. §5.3).

### 3.2 Commerciaux (10)

`code_commercial` (C01…C10), `nom`, `code_region` (2 par région), `date_embauche`.
- Dates d'embauche réparties entre 2019 et 2024.
- **1 commercial « en montée »** : embauché `2024-09`, sa part de commandes monte de ~0 à
  niveau normal sur ~6 mois (signal *ramp-up*).
- **1 commercial « en déclin »** : sa part de commandes décroît d'environ −30 % sur les 24 mois.

### 3.3 Produits (120) — profils de marge contrastés

`code_produit` (P0001…), `libelle`, `categorie`, `sous_categorie`, `cout_unitaire`,
`prix_catalogue`, `fournisseur`.

`prix_catalogue = round(cout_unitaire / (1 - taux_marge_cible), 2)`. **La marge n'est jamais
stockée** : elle se recalcule en SQL (vue analytique).

| Catégorie | Profil | `cout_unitaire` (€) | Taux de marge cible | Ticket |
|---|---|---|---|---|
| Consommables | marge ↑↑ | 1 – 12 | 45 – 60 % | petit |
| Ingrédients secs | marge ↑ | 3 – 25 | 30 – 45 % | petit |
| Hygiène & Entretien | marge ↑ | 4 – 40 | 35 – 50 % | petit/moyen |
| Petit équipement | marge ~ | 30 – 300 | 20 – 30 % | moyen |
| Gros équipement | marge ↓↓ | 400 – 3 500 | 8 – 16 % | gros |
| Mobilier | marge ↓ | 50 – 800 | 12 – 22 % | gros |

Famille **« équipement »** = {Petit équipement, Gros équipement, Mobilier} (sert au signal de mix).

### 3.4 Clients (500)

`code_client` (CL0001…), `raison_sociale`, `segment`, `type_etablissement`, `code_region`,
`ville`, `code_commercial`, `date_premiere_commande`.

- `segment` : Grands Comptes **10 %** (~50), PME **50 %** (~250), Indépendants **40 %** (~200).
  Les Grands Comptes commandent plus souvent et plus gros.
- `type_etablissement` ∈ {Restaurant, Hôtel, Bar, Collectivité}.
- `code_commercial` tiré parmi les commerciaux de la région du client.
- **Cohorte de churn** : ~35 clients **PME** à `date_premiere_commande` ancienne (2024 H1) qui
  **cessent de commander après `2025-06`** (signal *rétention*).

---

## 4. Étape 1 — Générer les faits propres (signaux métier inclus)

### 4.1 Commandes (`fait_commande`, en-tête)

`commande_id` (CMDyyyymmNNNNN), `date_commande`, `code_client`, `code_commercial` (hérité du
client), `statut`.

- **Volume mensuel** = base × multiplicateur de saisonnalité × croissance YoY :

  | Mois | J | F | M | A | M | J | J | A | S | O | N | D |
  |---|---|---|---|---|---|---|---|---|---|---|---|---|
  | Mult. | 0,80 | 0,85 | 0,95 | 1,00 | 1,05 | 1,20 | 1,30 | 1,25 | 1,05 | 1,00 | 0,95 | 1,25 |

  (pics été + décembre, logique CHR ; même forme les deux années, base 2025 × 1,08).

- **`statut`** : Livrée **92 %**, Retour **5 %**, Annulée **3 %**.

### 4.2 Lignes de commande (`fait_ligne_commande`, grain analytique)

`ligne_id`, `commande_id`, `code_produit`, `quantite` (1–8 selon catégorie), `prix_unitaire_vente`
(= `prix_catalogue` ± petit bruit), `remise_pct`.

- **`remise_pct` normale** : U(0 %, 8 %), **arrondie à 3 décimales** (pas de 0,1 %) pour
  s'aligner sur la précision de la colonne `Remise %` du fichier et garantir un bouclage exact.
- **Signal « glissement de mix »** : la part de lignes « équipement » monte de **~22 %**
  (début 2024) à **~35 %** (fin 2025), rampe concentrée sur **H2 2025**. Conséquence mécanique :
  **CA en hausse, taux de marge global en érosion** → l'insight phare.
- **Signal « fuite de marge »** : **3 produits** de catégories à marge élevée reçoivent une
  `remise_pct` anormale tirée dans U(20 %, 40 %) à chaque apparition → marge réalisée effondrée
  sur ces références.

### 4.3 Définition de la vérité financière (CA & marge)

- **CA vrai** = Σ sur lignes de commandes **Livrées** de `quantite × prix_unitaire_vente × (1 − remise_pct)`.
- **Marge vraie** = Σ sur les mêmes lignes de `quantite × (prix_unitaire_vente × (1 − remise_pct) − cout_unitaire)`.
- Les commandes **Annulées** sont **exclues** du CA ; les **Retours** sont conservés comme drapeau
  (l'analyste décide du traitement en Projet 2 — bon point d'analyse « brut vs net »).

Ces sommes sont calculées **ici, sur les valeurs propres**, et figées dans le manifeste (§7).

---

## 5. Étape 2 — Salissage par source (injection des anomalies)

Ordre d'injection **fixe** (reproductibilité) : Ventes → CRM → Catalogue → Budget. Chaque
injecteur **incrémente ses compteurs** dans le manifeste au moment où il agit.

### 5.1 `export_ventes_AAAA-MM.xlsx` (×24) — export ERP dénormalisé

Fichier **plat au grain ligne** (les colonnes d'en-tête se répètent sur les lignes d'une même
commande — réaliste). En-têtes « sales », **codes et non noms** (force les JOINs) :

`Date Commande` · `N° Commande` · `N° Ligne` · `Code Client` · `Code Produit` · `Quantité `
· `PU Vente HT ` · `Remise %` · `Code Commercial` · `Statut`

> `N° Ligne` = identifiant de ligne (`ligne_id`). Clé de déduplication naturelle
> `(N° Commande, N° Ligne)` : un doublon d'export recopie la ligne entière, n° compris.

| Anomalie | Nature | Volume | Détail |
|---|---|---|---|
| Montants format FR | **systématique** | **100 %** | `PU Vente HT` écrit « 1 234,56 » (virgule décimale, espace insécable milliers) |
| En-têtes à espaces | **systématique** | 2 colonnes | `Quantité ` et `PU Vente HT ` (espace final) |
| Lignes d'ajustement manuelles | sporadique | **~0,4 %** des lignes | `PU Vente HT` = « N/A » ou vide → montant non numérique |
| … dont date impossible | sporadique | ~moitié des lignes manuelles (~0,2 %) | `Date Commande` = « 32/13/AAAA » (seul un humain produit ça) |
| Doublons d'export | sporadique | **~1 %** des lignes | lignes dupliquées, concentrées sur commandes des 3 premiers/derniers jours du mois |

### 5.2 `export_clients_crm.csv` — référentiel client (saisie commerciale)

Colonnes : `code_client` · `raison_sociale` · `segment` · `type_etablissement` · `code_region`
· `ville` · `code_commercial` · `date_premiere_commande`.

| Anomalie | Nature | Volume | Détail |
|---|---|---|---|
| Encodage **cp1252** | **systématique** | tout le fichier | les autres fichiers sont en utf-8 → la diversité est *entre* fichiers, chacun étant cohérent |
| `date_premiere_commande` KO | sporadique | **~1,5 %** des clients | faute de frappe / collision JJ-MM vs MM-JJ |
| Casse & espaces `segment` | modéré | **~12 %** des lignes | « Grands Comptes » vs « grands comptes » vs « GRANDS COMPTES » |

### 5.3 `catalogue_produits.xlsx` — référentiel maintenu à la main

Colonnes : `code_produit` · `libelle` · `categorie` · `sous_categorie` · `cout_unitaire`
· `prix_catalogue` · `fournisseur`.

| Anomalie | Nature | Volume | Détail |
|---|---|---|---|
| Doublons orthographiques de catégorie | **structurel** | 2 catégories sur 6 | variantes sur une fraction des produits concernés |

Variantes exactes injectées (à figer) :
- `Consommables` → certains en `consommables`, `Conso.`
- `Hygiène & Entretien` → certains en `hygiene et entretien`

⚠️ Anomalie **la plus importante à raconter** : non corrigée, elle éclate une catégorie en
plusieurs → **fausse silencieusement l'effet mix**, cœur de l'analyse.

### 5.4 `budget_AAAA.xlsx` (×2, un par année) — fichier du contrôle de gestion

Format **large**, jamais un export système. Mise en page **structurelle (100 %)** :
- 1 ligne de titre (« Budget commercial AAAA »), cellules fusionnées.
- 1 feuille `CA` et 1 feuille `Marge`.
- Lignes = `région × catégorie` ; **colonnes = 12 mois** (`Janv`…`Déc`).
- **Lignes de sous-total par région** intercalées (à supprimer au nettoyage).
- Le signal **Sud sous budget** est porté ici : facteur budget général `U(0,93 ; 1,07)`
  (écart ~0 % hors Sud), et **Sud sur-budgété** (`U(1,15 ; 1,25)`) → le réel y reste ~16 %
  sous l'objectif, seul point noir.

Nettoyage attendu = **reshape `melt`** (large → long) + suppression des sous-totaux (type de
transformation différent du simple `strip`).

### 5.5 `commerciaux.csv` — export sales-ops / RH

Colonnes : `code_commercial` · `nom` · `code_region` · `date_embauche`.

**Propre, volontairement.** Sert de contraste (« toutes les sources ne sont pas sales ») et
fournit le JOIN qui apporte `nom` et `date_embauche` (nécessaires aux signaux commerciaux et au
Projet 2). Encodage utf-8.

---

## 6. Écriture des fichiers — récapitulatif

| Fichier | Format | Encodage | Grain |
|---|---|---|---|
| `export_ventes_AAAA-MM.xlsx` (×24) | xlsx | — | ligne de commande |
| `export_clients_crm.csv` | csv `;` | **cp1252** | client |
| `catalogue_produits.xlsx` | xlsx | — | produit |
| `budget_2024.xlsx`, `budget_2025.xlsx` | xlsx (large) | — | région×catégorie |
| `commerciaux.csv` | csv `;` | utf-8 | commercial |
| `_manifest_anomalies.json` | json | utf-8 | vérité terrain |

> La table calendrier `dim_date` n'est **pas** un fichier source : elle est dérivée de la
> plage de dates dans le Projet 2 (modélisation Postgres / Power BI).

---

## 7. Format du manifeste (`_manifest_anomalies.json`)

```json
{
  "generation": {
    "seed": 42,
    "generated_at": "2025-…T…Z",
    "periode": "2024-01..2025-12",
    "version_generateur": "1.1.0"
  },
  "verite_financiere": {
    "ca_total_livre_eur": 0.00,
    "marge_total_livree_eur": 0.00,
    "taux_marge_global_pct": 0.00,
    "nb_commandes": 0,
    "nb_lignes_total": 0,
    "nb_lignes_livrees": 0,
    "nb_categories_canoniques": 6
  },
  "anomalies_injectees": {
    "ventes": {
      "montants_format_fr": { "nature": "systematique", "nb_lignes": 0 },
      "entetes_espaces": { "colonnes": ["Quantité ", "PU Vente HT "] },
      "lignes_manuelles_montant_invalide": { "nb": 0, "ca_perdue_livree_eur": 0.00 },
      "lignes_manuelles_date_invalide": { "nb": 0 },
      "doublons_export": { "nb_lignes": 0, "commandes_concernees": [] }
    },
    "crm": {
      "encodage": "cp1252",
      "date_premiere_commande_invalide": { "nb": 0 },
      "segment_casse_variante": { "nb_lignes": 0 }
    },
    "catalogue": {
      "categories_variantes": {
        "Consommables": ["consommables", "Conso."],
        "Hygiène & Entretien": ["hygiene et entretien"],
        "nb_produits_impactes": 0
      }
    },
    "budget": { "format": "large", "sous_totaux_intercales": true, "signal_sud_sous_budget": true }
  },
  "attendus_apres_nettoyage": {
    "ca_reconcilie_attendu_eur": 0.00,
    "nb_lignes_apres_dedup": 0,
    "nb_dates_invalides_attendu": 0,
    "nb_montants_invalides_attendu": 0,
    "nb_categories_apres_normalisation": 6
  }
}
```

Le bloc **`attendus_apres_nettoyage`** est ce que les tests pytest de bout en bout vérifient :
le pipeline du Projet 1, lancé sur `data_raw/`, doit retrouver `ca_reconcilie_attendu_eur`
(à la tolérance flottante près, < 0,01 €), exactement `nb_dates_invalides_attendu` dates KO,
et ramener à `nb_categories_apres_normalisation` catégories.

> **`ca_reconcilie_attendu_eur` = `ca_total_livre_eur` − `ca_perdue_livree_eur`.** Les lignes
> manuelles à montant invalide sont irrécupérables (montant détruit) : le CA réconciliable est
> donc le CA vrai diminué de leur contribution livrée. C'est ce qui rend l'oracle exact plutôt
> qu'approché.

---

## 8. Signatures de fonctions (pseudo-code, à figer avant implémentation)

```python
# --- dimensions propres (anglais pour le code, schéma métier en français) ---
def build_regions() -> pd.DataFrame: ...
def build_salespeople(rng) -> pd.DataFrame: ...                 # date d'embauche recrue figée
def build_products(rng) -> pd.DataFrame: ...                    # profils de marge par catégorie
def build_customers(rng, salespeople) -> pd.DataFrame: ...      # salespeople injecté (déterminisme)

# --- signaux explicites (rendus testables hors des faits) ---
def select_margin_leak_products(products, rng) -> list[str]: ...
def select_churn_cohort(customers) -> list[str]: ...

# --- faits (porte tous les signaux métier) ---
def build_orders_and_lines(rng, customers, products, leak_products, churn_customers
                           ) -> tuple[pd.DataFrame, pd.DataFrame]: ...

# --- oracle : calculé sur les données PROPRES ---
def compute_truth(orders, order_lines, products) -> dict: ...

# --- salissage : chaque injecteur retourne (df_sale, compteurs) ---
def inject_sales_anomalies(rng, orders, order_lines) -> tuple[pd.DataFrame, dict]:
    """df_sale porte une colonne _month pour le découpage ; calcule ca_perdue_livree_eur."""
def inject_crm_anomalies(rng, customers) -> tuple[pd.DataFrame, dict]: ...
def inject_catalog_anomalies(rng, products) -> tuple[pd.DataFrame, dict]: ...
def build_budget_long(rng, orders, order_lines, products, customers, regions
                      ) -> tuple[pd.DataFrame, dict]:
    """Budget en format long (réel + budgété) ; la forme large est appliquée à l'écriture."""

# --- écriture (un écrivain par source) ---
def write_sales(sales_dirty, raw_dir) -> int: ...      # 24 fichiers mensuels
def write_crm(crm_dirty, raw_dir) -> None: ...         # csv cp1252
def write_catalog(catalog_dirty, raw_dir) -> None: ...
def write_salespeople(salespeople, raw_dir) -> None: ...   # csv utf-8 propre
def write_budget_wide(budget_long, raw_dir) -> None:   # mise en page large via openpyxl
def write_manifest(truth, c_sales, c_crm, c_catalog, c_budget, manifest_path) -> dict: ...

# --- orchestration + preuve ---
def generate_dataset(raw_dir, manifest_path) -> dict:  # déterministe, logging
def verify_roundtrip(raw_dir, manifest) -> None:       # nettoyage minimal vs oracle
def main() -> None: ...
```

---

## 9. Invariants vérifiables (ce que le manifeste permet d'affirmer en test)

1. `ca_reconcilie_attendu_eur` == CA recalculé par le pipeline (tolérance = `RECON_TOLERANCE_PCT`).
2. Nombre de dates invalides détectées == `nb_dates_invalides_attendu`.
3. Nombre de montants non convertibles == `nb_montants_invalides_attendu`.
4. Après dédup : `nb_lignes_apres_dedup` lignes restantes.
5. Après normalisation : exactement **6** catégories canoniques (les variantes ont fusionné).

> Dose de tests : **quelques tests qui prouvent quelque chose** (les 5 invariants + les cas
> limites unitaires existants), pas une cathédrale. Le testing est un signal de sérieux, pas
> le cœur de la vitrine analyste.