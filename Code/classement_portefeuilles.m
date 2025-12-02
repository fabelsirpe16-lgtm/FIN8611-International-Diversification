%2e code à excuté
% Lire les fichiers contenant les classements
class98 = readtable('Classements_Pays_1998_2007.xlsx');
class08 = readtable('Classements_Pays_2008_2020.xlsx');

% Trier les pays selon le classement du risque systématique ajusté (Rank_Risk)
[~, classement_risque_98] = sort(class98.Rank_Risk, 'ascend'); % Tri ascendant pour les pays les plus segmentés
[~, classement_risque_08] = sort(class08.Rank_Risk, 'ascend'); % Tri ascendant pour les pays les plus segmentés

% Extraire les pays les plus et les moins segmentés pour 1998-2007
plus_segmentes_98 = class98(classement_risque_98(1:6), :); % 6 premiers (les plus segmentés)
moins_segmentes_98 = class98(classement_risque_98(end-5:end), :); % 6 derniers (les moins segmentés)

% Extraire les pays les plus et les moins segmentés pour 2008-2020
plus_segmentes_08 = class08(classement_risque_08(1:6), :); % 6 premiers (les plus segmentés)
moins_segmentes_08 = class08(classement_risque_08(end-5:end), :); % 6 derniers (les moins segmentés)

% Ajouter une colonne pour indiquer le type
plus_segmentes_98.Type = repmat("Plus segmenté", height(plus_segmentes_98), 1);
moins_segmentes_98.Type = repmat("Moins segmenté", height(moins_segmentes_98), 1);

plus_segmentes_08.Type = repmat("Plus segmenté", height(plus_segmentes_08), 1);
moins_segmentes_08.Type = repmat("Moins segmenté", height(moins_segmentes_08), 1);

% Combiner les pays segmentés pour chaque période
tableau_segmentes_98 = [plus_segmentes_98; moins_segmentes_98];
tableau_segmentes_08 = [plus_segmentes_08; moins_segmentes_08];

% Ajouter une colonne pour indiquer la période
tableau_segmentes_98.Periode = repmat("1998-2007", height(tableau_segmentes_98), 1);
tableau_segmentes_08.Periode = repmat("2008-2020", height(tableau_segmentes_08), 1);

% Conserver uniquement les colonnes pertinentes
colonnes_a_conserver = {'Country', 'Periode', 'Type', 'rendement_moyen', 'Alpha', 'Ris_Sys_Ajustee', 'Sharp_Ratio'};
tableau_segmentes_98 = tableau_segmentes_98(:, colonnes_a_conserver);
tableau_segmentes_08 = tableau_segmentes_08(:, colonnes_a_conserver);

% Exporter les tableaux distincts pour chaque période
writetable(tableau_segmentes_98, 'Tableau_Segmentes_1998_2007.xlsx');
writetable(tableau_segmentes_08, 'Tableau_Segmentes_2008_2020.xlsx');

% Lire les fichiers contenant les classements
class98 = readtable('Classements_Pays_1998_2007.xlsx');
class08 = readtable('Classements_Pays_2008_2020.xlsx');

% Trier les pays selon le classement du risque systématique ajusté (Rank_Risk)
[~, classement_risque_98] = sort(class98.Rank_Risk, 'ascend'); % Tri ascendant pour les pays les plus segmentés
[~, classement_risque_08] = sort(class08.Rank_Risk, 'ascend'); % Tri ascendant pour les pays les plus segmentés

% Extraire les pays les plus et les moins segmentés pour 1998-2007
plus_segmentes_98 = class98(classement_risque_98(1:6), :); % 6 premiers (les plus segmentés)
moins_segmentes_98 = class98(classement_risque_98(end-5:end), :); % 6 derniers (les moins segmentés)

% Extraire les pays les plus et les moins segmentés pour 2008-2020
plus_segmentes_08 = class08(classement_risque_08(1:6), :); % 6 premiers (les plus segmentés)
moins_segmentes_08 = class08(classement_risque_08(end-5:end), :); % 6 derniers (les moins segmentés)

% Ajouter une colonne pour indiquer le type
plus_segmentes_98.Type = repmat("Plus segmenté", height(plus_segmentes_98), 1);
moins_segmentes_98.Type = repmat("Moins segmenté", height(moins_segmentes_98), 1);

plus_segmentes_08.Type = repmat("Plus segmenté", height(plus_segmentes_08), 1);
moins_segmentes_08.Type = repmat("Moins segmenté", height(moins_segmentes_08), 1);

% Combiner les pays segmentés pour chaque période
tableau_segmentes_98 = [plus_segmentes_98; moins_segmentes_98];
tableau_segmentes_08 = [plus_segmentes_08; moins_segmentes_08];

% Ajouter une colonne pour indiquer la période
tableau_segmentes_98.Periode = repmat("1998-2007", height(tableau_segmentes_98), 1);
tableau_segmentes_08.Periode = repmat("2008-2020", height(tableau_segmentes_08), 1);

% Conserver uniquement les colonnes pertinentes
colonnes_a_conserver = {'Country', 'Periode', 'Type', 'rendement_moyen', 'Alpha', 'Ris_Sys_Ajustee', 'Sharp_Ratio'};
tableau_segmentes_98 = tableau_segmentes_98(:, colonnes_a_conserver);
tableau_segmentes_08 = tableau_segmentes_08(:, colonnes_a_conserver);

% Exporter les tableaux distincts pour chaque période
writetable(tableau_segmentes_98, 'Tableau_Segmentes_1998_2007.xlsx');
writetable(tableau_segmentes_08, 'Tableau_Segmentes_2008_2020.xlsx');

%fin