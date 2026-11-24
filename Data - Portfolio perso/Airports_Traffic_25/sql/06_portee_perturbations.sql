-- 06_portee_perturbations.sql
-- Question : une anomalie touche-t-elle tout un pays (perturbation « réseau »)
-- ou un seul aéroport (incident « local ») ?
--
-- COUNT(*) OVER (PARTITION BY ...) compte, SANS regrouper les lignes,
-- combien d'aéroports du même pays sont perturbés le même jour.
-- Seuil retenu : 3 aéroports ou plus = réseau (choix simple, à justifier dans le README).

CREATE OR REPLACE VIEW `eco-flying-nice-vr.airports_traffic.jours_perturbes_portee` AS
SELECT
  *,
  COUNT(*) OVER (PARTITION BY flight_date, country)          AS nb_aeroports_pays_jour,
  CASE
    WHEN COUNT(*) OVER (PARTITION BY flight_date, country) >= 3 THEN 'Réseau'
    ELSE 'Local'
  END                                                        AS portee,
  ROUND(SAFE_DIVIDE(src1_total, moy_voisins) * 100 - 100, 1) AS variation_pct
FROM `eco-flying-nice-vr.airports_traffic.jours_perturbes_annotes`;

-- Synthèse portée × catégorie :
-- SELECT portee, categorie, COUNT(*) AS nb
-- FROM `eco-flying-nice-vr.airports_traffic.jours_perturbes_portee`
-- GROUP BY portee, categorie
-- ORDER BY portee, nb DESC;

-- Pistes d'enquête : perturbations réseau encore non expliquées
-- SELECT flight_date, country, ANY_VALUE(nb_aeroports_pays_jour) AS nb_aeroports,
--        STRING_AGG(airport_name, ', ') AS aeroports
-- FROM `eco-flying-nice-vr.airports_traffic.jours_perturbes_portee`
-- WHERE portee = 'Réseau' AND categorie = 'Non expliqué'
-- GROUP BY flight_date, country
-- ORDER BY flight_date;
