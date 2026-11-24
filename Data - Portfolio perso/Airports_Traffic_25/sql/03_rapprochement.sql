-- 03_rapprochement.sql
-- Question : quand les deux sources existent, sont-elles d'accord ?
-- SAFE_DIVIDE renvoie NULL au lieu de planter si on divise par zéro.

CREATE OR REPLACE VIEW `eco-flying-nice-vr.airports_traffic.rapprochement` AS
SELECT
  flight_date,
  airport_icao,
  airport_name,
  country,
  src1_total,
  src2_total,
  src2_total - src1_total                                  AS ecart,
  SAFE_DIVIDE(src2_total - src1_total, src1_total)         AS ecart_pct,
  CASE
    WHEN src2_total IS NULL                                         THEN 'source 2 absente'
    WHEN ABS(SAFE_DIVIDE(src2_total - src1_total, src1_total)) > 0.05 THEN 'écart > 5 %'
    WHEN src2_total <> src1_total                                    THEN 'écart ≤ 5 %'
    ELSE 'identique'
  END                                                      AS statut
FROM `eco-flying-nice-vr.airports_traffic.stg_airport_traffic`;

-- Synthèse rapide à lancer ensuite :
-- SELECT statut, COUNT(*) AS nb_lignes
-- FROM `eco-flying-nice-vr.airports_traffic.rapprochement`
-- GROUP BY statut ORDER BY nb_lignes DESC;
