# Inactivity Model - Data Governance Naming Standards

## Document Overview

This document outlines the required changes to bring the inactivity model pipeline into compliance with PLDT Data Mesh Standards 2026. The changes focus on table naming conventions, catalog usage, data classification tags, and documentation standards.

## Catalog Usage by Environment

### Local Development (local target)
- **Catalog**: `ds_dev_e` (Exception catalog)
- **Purpose**: Exploratory analysis and local testing only
- **Restrictions**: NOT for production pipelines or scheduled jobs

### Shared Development (dev target)
- **Catalog**: `ds_dev` (Standard catalog)
- **Purpose**: Integration testing and development validation
- **Compliance**: Must follow all naming and tagging standards

### Pre-Production (pet target)
- **Catalog**: `ds_pet` (Standard catalog)
- **Purpose**: Pre-production testing and UAT
- **Compliance**: Must follow all naming and tagging standards

### Production (prod target)
- **Catalog**: `ds` (Standard catalog)
- **Purpose**: Production deployment
- **Compliance**: Must follow all naming and tagging standards

## Schema Assignment

Based on PLDT Data Mesh Standards 2026:

| Current Schema | Required Schema | Layer | Reason |
|---------------|-----------------|-------|--------|
| `tmp_kacd` | `CU_S` | Silver | Customer domain, curated features |
| N/A (future) | `CU_G` | Gold | Customer domain, ML predictions |

- **CU_S** (Customer Silver): For feature engineering tables (base_subs, dos_features, topup_features, training_set, oot_set)
- **CU_G** (Customer Gold): For inference prediction outputs (fact tables)

## Table Naming Conventions

### Naming Pattern Requirements

PLDT standards require:
1. **Wireless Indicator**: `wi_` prefix for all wireless/telecom tables
2. **Frequency Indicator**: Position varies by layer
   - Silver: `<business_group>_<freq>_<dataset_name>`
   - Gold: `<table_type>_<business_group>_<dataset_name>_<freq>`
3. **Lowercase**: All table names must be lowercase (following existing PLDT convention)
4. **Underscores**: Use underscores as separators (no hyphens)

### Frequency Codes

| Code | Description | Use Case |
|------|-------------|----------|
| `DLY` | Daily | Daily feature snapshots |
| `WKY` | Weekly | Weekly aggregations |
| `MLY` | Monthly | Monthly aggregations |
| `HRY` | Hourly | Hourly updates |
| `RTM` | Realtime | Streaming/realtime predictions |

## Required Table Name Changes

### Training Feature Tables (Silver Layer - CU_S schema)

| Current Name | Compliant Name | Description |
|--------------|----------------|-------------|
| `inactivity_base_subs` | `wi_ml_inactivity_base_subs_dly` | Base subscriber dataset with DoS labels and splits |
| `inactivity_dos_features` | `wi_ml_inactivity_dos_features_dly` | Days of Silence features (50+ features) |
| `inactivity_topup_features` | `wi_ml_inactivity_topup_features_dly` | Topup/recharge behavior features (40+ features) |
| `inactivity_training_set` | `wi_ml_inactivity_training_set_dly` | Final ML-ready training dataset (train/val/test) |
| `inactivity_oot_set` | `wi_ml_inactivity_oot_set_dly` | Out-of-time evaluation dataset |

### Inference Prediction Tables (Gold Layer - CU_G schema)

| Current Name | Compliant Name | Description |
|--------------|----------------|-------------|
| TBD | `f_wi_ml_inactivity_pred_dly` | Daily inactivity predictions (fact table) |
| TBD | `f_wi_ml_inactivity_pred_rtm` | Realtime inactivity predictions (if needed) |

**Naming breakdown for prediction tables:**
- `f_` = Fact table prefix (Gold layer convention)
- `wi_` = Wireless indicator
- `ml_` = Machine Learning business group
- `inactivity_` = Dataset name
- `pred_` = Prediction output
- `dly/rtm` = Frequency indicator

## Data Classification Tags

### PII Tagging Requirements

All tables containing `msisdn` must be tagged with:
```python
table_properties={
    "dataprivacy": "personal",  # PII classification
    "quality": "silver",        # or "gold" for inference
    "delta.autoOptimize.optimizeWrite": "true",
}
```

### PII Column Tagging

The `msisdn` column must include comment with PII indicator:
```python
# In @dp.table decorator or Delta table properties
# Column comments should indicate: "Subscriber mobile number (PII)"
```

### Tables Requiring PII Tags

All 5 training tables + inference output tables:
1. wi_ml_inactivity_base_subs_dly
2. wi_ml_inactivity_dos_features_dly
3. wi_ml_inactivity_topup_features_dly
4. wi_ml_inactivity_training_set_dly
5. wi_ml_inactivity_oot_set_dly
6. f_wi_ml_inactivity_pred_dly (inference)

## Column Comments Standards

### Required Comments

Every column must have a descriptive comment following these guidelines:

1. **Standard Columns** (consistent across all tables):
   - `msisdn`: "Subscriber mobile number (PII)"
   - `feature_dt`: "Feature observation date (YYYY-MM-DD)"
   - `split_set`: "Dataset split indicator: train/val/test/oot"
   - `promo_topup_cluster`: "Subscriber segmentation cluster based on promo and topup behavior"
   - `brand`: "Subscriber brand: SMT_CEL_PRE or TNT_CEL_PRE"
   - `dbx_process_dttm`: "Databricks processing timestamp (UTC)"

2. **Feature Columns**: Brief description of calculation logic
   - Example: `dos_lag1`: "Days of Silence from previous day (0-7 range)"
   - Example: `topup_amt_30d`: "Total topup amount in last 30 days (PHP)"

3. **Target/Label Columns**: Definition of positive class
   - Example: `dos7_inactive_label`: "Binary label: 1=subscriber became inactive (DoS increased), 0=remained active"

### Comment Consistency

- Same column name across tables = same comment text
- Use standard units (PHP for currency, days for time)
- Include value ranges where applicable

## Table Comments Standards

Each table must have a comprehensive comment in the `@dp.table` decorator:

```python
@dp.table(
    name="WI_ML_INACTIVITY_BASE_SUBS_DLY",
    comment="Daily base subscriber dataset for inactivity prediction model training. " \
            "Contains eligible prepaid subscribers (TZWB<90, target brands) with DoS-based " \
            "labels, train/val/test/oot splits, and promo_topup_cluster segmentation. " \
            "Target label identifies subscribers at risk of becoming inactive (DoS 0-7 who experience DoS increase).",
    partition_cols=["feature_dt", "promo_topup_cluster", "split_set"],
    table_properties={
        "dataprivacy": "personal",
        "quality": "silver",
        "delta.autoOptimize.optimizeWrite": "true",
    }
)
```

**Comment requirements:**
- Purpose of the table
- Key filtering logic
- Business definition of target variable
- Data lineage (source tables)
- Refresh frequency

## Implementation Plan

### Phase 1: Local Development Environment (Immediate)

**Target**: local (ds_dev_e catalog)

1. Update pipeline configuration:
   - `models/src/inactivity_model/resources/inactivity_feature_prep_training.pipeline.yml`
   - Change catalog references if needed for local testing
   
2. Rename tables in transformation scripts:
   - `01_inactivity_base_subs.ipynb` → update `@dp.table(name="wi_ml_inactivity_base_subs_dly")`
   - `02_inactivity_dos_features.ipynb` → update `@dp.table(name="wi_ml_inactivity_dos_features_dly")`
   - `03_inactivity_topup_features.ipynb` → update `@dp.table(name="wi_ml_inactivity_topup_features_dly")`
   - `05_inactivity_training_set.ipynb` → update `@dp.table(name="wi_ml_inactivity_training_set_dly")`
   - `06_inactivity_oot_set.ipynb` → update `@dp.table(name="wi_ml_inactivity_oot_set_dly")`

3. Update JOIN references in downstream tables:
   - Training set: Update FROM/JOIN to use new table names
   - OOT set: Update FROM/JOIN to use new table names

4. Add PII tags to all table properties

5. Add comprehensive table comments

6. Add column comments (prioritize PII columns first)

**Status**: Use exception catalog (ds_dev_e) - acceptable for local testing

### Phase 2: Shared Development Environment (Before merge to non-prod)

**Target**: dev (ds_dev catalog)

1. Update `databricks.yml` variables for dev target:
   ```yaml
   targets:
     dev:
       variables:
         catalog_postfix: "_dev"  # Results in ds_dev
         output_schema: "cu_s"    # Changed from tmp_kacd
   ```

2. Update external table references:
   - `vab_promo_topup_cluster` table reference (currently in ds_dev_e.tmp_kacd)
   - Either: Move to ds_dev.cu_s or keep in exception catalog with proper documentation

3. Test full pipeline deployment:
   ```bash
   databricks bundle deploy --target dev
   databricks bundle run inactivity_feature_prep_training --target dev
   ```

4. Validate data quality expectations still pass

**Status**: Must use standard catalog (ds_dev)

### Phase 3: Pre-Production Environment (Before merge to pet)

**Target**: pet (ds_pet catalog)

1. Update `databricks.yml` variables for pet target:
   ```yaml
   targets:
     pet:
       variables:
         catalog_postfix: "_pet"  # Results in ds_pet
         output_schema: "cu_s"
   ```

2. Create inference prediction job (if not exists):
   - Use table name: `f_wi_ml_inactivity_pred_dly`
   - Schema: `cu_g` (Gold layer)
   - Add all required tags and comments

3. Test full training + inference workflow

4. Validate with Data Governance team

**Status**: Must use standard catalog (ds_pet)

### Phase 4: Production Environment (Before merge to prod)

**Target**: prod (ds catalog)

1. Update `databricks.yml` variables for prod target:
   ```yaml
   targets:
     prod:
       variables:
         catalog_postfix: ""  # Results in ds
         output_schema: "cu_s"
   ```

2. Verify all compliance requirements:
   - [ ] All tables use lowercase names with wi_ prefix
   - [ ] All tables include frequency indicator (_dly)
   - [ ] All tables with msisdn have dataprivacy=personal tag
   - [ ] All tables have comprehensive comments
   - [ ] All columns have descriptive comments
   - [ ] Standard catalog (ds) is used
   - [ ] Proper schema assignment (cu_s/cu_g)

3. Final approval from Data Governance

**Status**: Must use standard catalog (ds)

## Code Changes Required

### Example: Update @dp.table decorator

**Before:**
```python
@dp.table(
    name="inactivity_base_subs",
    comment="Base subscriber dataset for inactivity model training with DoS features and labels",
    partition_cols=["feature_dt", "promo_topup_cluster", "split_set"]
)
```

**After:**
```python
@dp.table(
    name="wi_ml_inactivity_base_subs_dly",
    comment="Daily base subscriber dataset for inactivity prediction model training. " \
            "Contains eligible prepaid subscribers (TZWB<90 days, SMT_CEL_PRE and TNT_CEL_PRE brands) " \
            "with DoS-based inactivity labels, train/val/test/oot splits, and promo_topup_cluster segmentation. " \
            "Target label (dos7_inactive_label) identifies subscribers with DoS 0-7 who become inactive (DoS increase). " \
            "Source: wde.cu_s.wi_prepaid_account_profile_dly, wa.cu_s.wi_dly_pre_dwedos_pd, vab_promo_topup_cluster.",
    partition_cols=["feature_dt", "promo_topup_cluster", "split_set"],
    table_properties={
        "dataprivacy": "personal",  # Contains msisdn (PII)
        "quality": "silver",
        "delta.autoOptimize.optimizeWrite": "true",
    }
)
```

### Example: Update JOIN references

**Before:**
```python
FROM inactivity_base_subs subs
LEFT JOIN inactivity_dos_features dos
  ON subs.msisdn = dos.msisdn AND subs.feature_dt = dos.feature_dt
LEFT JOIN inactivity_topup_features topup
  ON subs.msisdn = topup.msisdn AND subs.feature_dt = topup.feature_dt
```

**After:**
```python
FROM wi_ml_inactivity_base_subs_dly subs
LEFT JOIN wi_ml_inactivity_dos_features_dly dos
  ON subs.msisdn = dos.msisdn AND subs.feature_dt = dos.feature_dt
LEFT JOIN wi_ml_inactivity_topup_features_dly topup
  ON subs.msisdn = topup.msisdn AND subs.feature_dt = topup.feature_dt
```

## Training Notebook Changes

Update model training notebooks to reference new table names:

**Before:**
```python
training_data = spark.read.table(f"{catalog}.{schema}.inactivity_training_set")
oot_data = spark.read.table(f"{catalog}.{schema}.inactivity_oot_set")
```

**After:**
```python
training_data = spark.read.table(f"{catalog}.{schema}.wi_ml_inactivity_training_set_dly")
oot_data = spark.read.table(f"{catalog}.{schema}.wi_ml_inactivity_oot_set_dly")
```

## Validation Checklist

Use this checklist before deploying to each environment:

### Table Naming
- [ ] All table names are lowercase (following existing PLDT convention)
- [ ] All table names start with `wi_` (wireless indicator)
- [ ] All table names include frequency indicator (`_dly`)
- [ ] Silver tables follow pattern: `wi_ml_inactivity_<dataset>_dly`
- [ ] Gold tables follow pattern: `f_wi_ml_inactivity_<dataset>_dly`

### Catalog & Schema
- [ ] Local target uses `ds_dev_e` (exception catalog)
- [ ] Dev/PET/PROD targets use standard catalogs (`ds_dev`, `ds_pet`, `ds`)
- [ ] Silver tables use schema `cu_s`
- [ ] Gold tables (inference) use schema `cu_g`

### Data Privacy
- [ ] All tables with `msisdn` have `dataprivacy=personal` tag
- [ ] `msisdn` column comment includes "(PII)" indicator
- [ ] No other PII columns exist (or properly tagged if they do)

### Documentation
- [ ] All tables have comprehensive comment (purpose, logic, sources)
- [ ] All columns have descriptive comments
- [ ] Standard columns use consistent comment text across tables
- [ ] Comments include units (PHP, days) and ranges where applicable

### Code References
- [ ] All JOIN references updated to new table names
- [ ] All spark.read.table() calls updated to new table names
- [ ] Training notebooks updated to use new table names
- [ ] Pipeline configuration reviewed for any hardcoded references

### Testing
- [ ] Pipeline deploys successfully: `databricks bundle deploy --target <env>`
- [ ] Pipeline runs successfully: `databricks bundle run inactivity_feature_prep_training --target <env>`
- [ ] All data quality expectations pass
- [ ] Training notebooks can read from new table names
- [ ] No broken references or missing tables

## References

- **PLDT Data Mesh Standards 2026**: `C:\Users\kcdelasalas\OneDrive - PLDT\ai-data-assist\Data Mesh Standards 2026 Cascade v2.pdf`
- **Genie Space Instructions**: `C:\Users\kcdelasalas\OneDrive - PLDT\ai-data-assist\genie_space_instructions.md`
- **Inactivity Model Pipeline**: `models/src/inactivity_model/`
- **Feature Preparation Transformations**: `models/src/inactivity_model/feature_preparation/transformations/training/`

## Summary of Changes

| Category | Change Count | Priority |
|----------|--------------|----------|
| Table name changes | 5 tables (+ inference TBD) | HIGH |
| Schema changes | 1 (tmp_kacd → cu_s) | HIGH |
| PII tags to add | 5-6 tables | HIGH |
| Table comments to enhance | 5-6 tables | MEDIUM |
| Column comments to add | ~100+ columns | MEDIUM |
| JOIN reference updates | 2 files | HIGH |
| Training notebook updates | 3 files | HIGH |
| Catalog configuration | 3 targets (dev/pet/prod) | HIGH |

## Contact

For questions about these standards or implementation guidance:
- **Data Governance Team**: [Contact info]
- **MLOps Team Lead**: 
- **Databricks Support**: [Contact info]

---

**Document Version**: 1.0  
**Last Updated**: 2026-05-13  
**Author**: KC Dela Salas  
**Status**: Draft - Pending Review
