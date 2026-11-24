-- 05_jours_perturbes_annotes.sql
-- Objectif : rattacher chaque jour perturbé (vue 04) à un événement connu
-- (table evenements_2025, remplie à la main à partir de recherches sourcées).
--
-- Deux jointures, car un événement peut être :
--   - propre à un pays (ex. grève en France)       → jointure sur date + pays
--   - commun à plusieurs pays (ex. Noël)           → jointure sur date + pays = 'Plusieurs'
-- COALESCE garde l'explication la plus précise (le pays) en priorité.

CREATE OR REPLACE VIEW `eco-flying-nice-vr.airports_traffic.jours_perturbes_annotes` AS
SELECT
  jp.*,
  COALESCE(ev_pays.categorie, ev_multi.categorie, 'Non expliqué') AS categorie,
  COALESCE(ev_pays.evenement, ev_multi.evenement)                 AS evenement,
  COALESCE(ev_pays.source_url, ev_multi.source_url)               AS source_url
FROM `eco-flying-nice-vr.airports_traffic.jours_perturbes` AS jp
LEFT JOIN `eco-flying-nice-vr.airports_traffic.evenements_2025` AS ev_pays
  ON  ev_pays.date = jp.flight_date
  AND ev_pays.pays = jp.country
LEFT JOIN `eco-flying-nice-vr.airports_traffic.evenements_2025` AS ev_multi
  ON  ev_multi.date = jp.flight_date
  AND ev_multi.pays = 'Plusieurs';

-- Synthèse : quelle part des anomalies est expliquée ?
-- SELECT categorie, COUNT(*) AS nb_jours_aeroport,
--        ROUND(COUNT(*) / SUM(COUNT(*)) OVER () * 100, 1) AS pct
-- FROM `eco-flying-nice-vr.airports_traffic.jours_perturbes_annotes`
-- GROUP BY categorie
-- ORDER BY nb_jours_aeroport DESC;
