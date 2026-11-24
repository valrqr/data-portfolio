-- 01_staging.sql
-- Objectif : une table propre, typée et renommée, SANS supprimer de lignes.
-- Les valeurs absentes de la source n°2 restent NULL (absence d'information ≠ zéro vol).

CREATE OR REPLACE VIEW `eco-flying-nice-vr.airports_traffic.stg_airport_traffic` AS
SELECT
  DATE(FLT_DATE)                    AS flight_date,
  APT_ICAO                          AS airport_icao,
  APT_NAME                          AS airport_name,
  STATE_NAME                        AS country,

  -- Source n°1 (toujours renseignée)
  FLT_DEP_1                         AS src1_dep,
  FLT_ARR_1                         AS src1_arr,
  FLT_TOT_1                         AS src1_total,

  -- Source n°2 (renseignée pour une partie des aéroports seulement)
  CAST(FLT_DEP_IFR_2 AS INT64)      AS src2_dep,
  CAST(FLT_ARR_IFR_2 AS INT64)      AS src2_arr,
  CAST(FLT_TOT_IFR_2 AS INT64)      AS src2_total
FROM `eco-flying-nice-vr.airports_traffic.raw_airport_traffic_2025`;
