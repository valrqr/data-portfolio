 -- models/intermediate/int_diabetic_data.sql

with staging as (

    select * from {{ ref('stg_diabetic_data') }}

),

specialites as (

    select * from {{ ref('stg_specialites_medicales') }}

),

enriched as (

    select
        staging.* except (insulin, specialite_medecin, race, gender),

        specialites.specialite_id,

    --transformations/traitement des valeurs nulles, ?, etc...
        --- cas ethnie : si ? ou plusieurs ethnies différentes enregistrées pour un même patient
         case
            when count(distinct race) over (partition by patient_nbr) > 1 then '?'
            else race
        end as ethnie,


        --- cas identique pour le genre, si deux genres connus pour un même patient 
        case
            when count(distinct gender) over (partition by patient_nbr) > 1 then '?'
            else gender
        end as genre

        from source


        --traduction des valeurs de la colonne insulin
        case
            when insulin = 'Up' then 'augmentation'
            when insulin = 'Down' then 'reduction'
            when insulin = 'Steady' then 'stable'
            when insulin = 'No' then 'Non'
        end as insulin,

        -- création de la colonne n_changement_posologie qui va compter les changements de traitements indiqués par Up et Down dans les colonnes médicaments
        (
            (case when metformin in ('Up','Down') then 1 else 0 end)
            + (case when repaglinide in ('Up', 'Down') then 1 else 0 end)
            + (case when nateglinide in ('Up', 'Down') then 1 else 0 end)
            + (case when chlorpropamide in ('Up', 'Down') then 1 else 0 end)
            + (case when glimepiride in ('Up', 'Down') then 1 else 0 end)
            + (case when acetohexamide in ('Up', 'Down') then 1 else 0 end)
            + (case when glipizide in ('Up', 'Down') then 1 else 0 end)
            + (case when glyburide in ('Up', 'Down') then 1 else 0 end)
            + (case when tolbutamide in ('Up', 'Down') then 1 else 0 end)
            + (case when pioglitazone in ('Up', 'Down') then 1 else 0 end)
            + (case when rosiglitazone in ('Up', 'Down') then 1 else 0 end)
            + (case when acarbose in ('Up', 'Down') then 1 else 0 end)
            + (case when miglitol in ('Up', 'Down') then 1 else 0 end)
            + (case when troglitazone in ('Up', 'Down') then 1 else 0 end)
            + (case when tolazamide in ('Up', 'Down') then 1 else 0 end)
            + (case when examide in ('Up', 'Down') then 1 else 0 end)
            + (case when insulin in ('Up', 'Down') then 1 else 0 end)
            + (case when glyburide_metformin in ('Up', 'Down') then 1 else 0 end)
            + (case when glipizide_metformin in ('Up', 'Down') then 1 else 0 end)
            + (case when glimepiride_pioglitazone in ('Up', 'Down') then 1 else 0 end)
            + (case when metformin_rosiglitazone in ('Up', 'Down') then 1 else 0 end)
            + (case when metformin_pioglitazone in ('Up', 'Down') then 1 else 0 end)
        ) as n_changement_posologie,

        -- création de la colonne readmis à partir du contenu de la colonne readmitted
        --- si on a la valeur <30 dans readmitted, readmis sera True, sinon False
        case
            when readmitted = '<30' then true
            when readmitted = 'NO' or readmitted = '>30' then false
        end as readmis

    from staging
    left join specialites
        on staging.specialite_medecin = specialites.specialite

)

select * from enriched