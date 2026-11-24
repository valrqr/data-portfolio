-- 04_jours_perturbes.sql  (version 2 : référence centrée)
-- Question : quels jours s'écartent anormalement du trafic habituel ?
--
-- Principe : on compare chaque jour aux MÊMES jours de la semaine
-- situés 1 à 4 semaines AVANT et APRÈS (8 « voisins » au maximum).
-- Pourquoi avant ET après ? La v1 ne regardait que les 8 semaines précédentes :
-- fin octobre, elle comparait l'hiver IATA à l'été et voyait des « anomalies »
-- partout (simple changement de programme). Une référence centrée corrige ça.
--
-- Le JOIN de la table sur elle-même (« self-join ») sert à aller chercher ces voisins.
-- z_score = (trafic du jour - moyenne des voisins) / écart-type des voisins.
-- En dessous de -3 : journée très inhabituelle → probablement un événement.

CREATE OR REPLACE VIEW `eco-flying-nice-vr.airports_traffic.jours_perturbes` AS
WITH voisins AS (
  SELECT
    jour.flight_date,
    jour.airport_icao,
    jour.airport_name,
    jour.country,
    jour.src1_total,
    AVG(voisin.src1_total)    AS moy_voisins,
    STDDEV(voisin.src1_total) AS ecart_type_voisins,
    COUNT(*)                  AS nb_voisins
  FROM `eco-flying-nice-vr.airports_traffic.stg_airport_traffic` AS jour
  JOIN `eco-flying-nice-vr.airports_traffic.stg_airport_traffic` AS voisin
    ON  voisin.airport_icao = jour.airport_icao
    AND ABS(DATE_DIFF(voisin.flight_date, jour.flight_date, DAY)) BETWEEN 7 AND 28
    AND MOD(DATE_DIFF(voisin.flight_date, jour.flight_date, DAY), 7) = 0   -- même jour de semaine
  GROUP BY jour.flight_date, jour.airport_icao, jour.airport_name, jour.country, jour.src1_total
)
SELECT
  flight_date,
  airport_icao,
  airport_name,
  country,
  src1_total,
  ROUND(moy_voisins, 1)                                             AS moy_voisins,
  ROUND(SAFE_DIVIDE(src1_total - moy_voisins, ecart_type_voisins), 2) AS z_score
FROM voisins
WHERE nb_voisins >= 4          -- pas de verdict sans un minimum de comparaisons
  AND moy_voisins >= 20        -- on ignore les très petits terrains, trop instables
  AND SAFE_DIVIDE(src1_total - moy_voisins, ecart_type_voisins) < -3;
