-- 02_couverture_sources.sql
-- Question : pour chaque aéroport, sur combien de jours la source n°2 est-elle disponible ?
-- On s'attend à trois profils : toujours (≈365 j), jamais (0 j), ou partiel (à expliquer).

SELECT
  airport_icao,
  airport_name,
  country,
  COUNT(*)                                       AS nb_jours,
  COUNTIF(src2_total IS NOT NULL)                AS nb_jours_src2,
  ROUND(COUNTIF(src2_total IS NOT NULL) / COUNT(*) * 100, 1) AS pct_couverture_src2,
  CASE
    WHEN COUNTIF(src2_total IS NOT NULL) = 0        THEN 'jamais'
    WHEN COUNTIF(src2_total IS NOT NULL) = COUNT(*) THEN 'toujours'
    ELSE 'partiel'
  END                                            AS profil_couverture
FROM `eco-flying-nice-vr.airports_traffic.stg_airport_traffic`
GROUP BY airport_icao, airport_name, country
ORDER BY pct_couverture_src2 DESC, nb_jours DESC;
