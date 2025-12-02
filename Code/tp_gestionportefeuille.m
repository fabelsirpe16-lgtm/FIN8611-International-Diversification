%1er code à executer 
clear all;
rf = 0.035; % Taux sans risque de 3,5%

% Charger les données
data08 = readtable('gport2.xlsx', 'Sheet', 'r 2008-2020');
data98 = readtable('gport2.xlsx', 'Sheet', 'r 1998-2007');

% Calculer les variances de RendementsMSCI
var_rw_08 = var(data08.RendementsMSCI); % Variance RW pour 2008-2020
var_rw_98 = var(data98.RendementsMSCI); % Variance RW pour 1998-2007

% Charger et traiter les feuilles indépendamment
results08 = process_sheet('gport2.xlsx', 'r 2008-2020');
results98 = process_sheet('gport2.xlsx', 'r 1998-2007');

% Charger les données pour la feuille de capitalisation boursière
capital_data = readtable('gport2.xlsx', 'Sheet', 'cap');
capital_data.Properties.VariableNames = {'Country', 'Capitalisation', 'Poids'};
capital_data.Country = regexprep(capital_data.Country, '^\s+|\s+$', ''); % Nettoyage des noms

% Correspondance des pays pour 1998-2007
cap_data_98 = capital_data(1:30, :);
[isMatch98, idx98] = ismember(results98.Country, cap_data_98.Country);
results98.Capitalisation = nan(height(results98), 1);
results98.Poids = nan(height(results98), 1);
results98.Capitalisation(isMatch98) = cap_data_98.Capitalisation(idx98(isMatch98));
results98.Poids(isMatch98) = cap_data_98.Poids(idx98(isMatch98));

% Calculs supplémentaires pour 1998-2007
results98.Fraction_RS_Pi = (results98.Beta .^ 2 .* var_rw_98) ./ results98.Variance_Ri;
results98.Ris_Sys_Ajustee = results98.Fraction_RS_Pi ./ results98.Poids;
results98.Risk_Premium = results98.rendement_moyen - rf;
results98.Sharp_Ratio = results98.Risk_Premium ./ sqrt(results98.Variance_Ri);
results98.Treynor = results98.Risk_Premium ./ results98.Beta;

% Correspondance des pays pour 2008-2020
cap_data_08 = capital_data(35:65, :);
[isMatch08, idx08] = ismember(results08.Country, cap_data_08.Country);
results08.Capitalisation = nan(height(results08), 1);
results08.Poids = nan(height(results08), 1);
results08.Capitalisation(isMatch08) = cap_data_08.Capitalisation(idx08(isMatch08));
results08.Poids(isMatch08) = cap_data_08.Poids(idx08(isMatch08));

% Calculs supplémentaires pour 2008-2020
results08.Fraction_RS_Pi = (results08.Beta .^ 2 .* var_rw_08) ./ results08.Variance_Ri;
results08.Ris_Sys_Ajustee = results08.Fraction_RS_Pi ./ results08.Poids;
results08.Risk_Premium = results08.rendement_moyen - rf;
results08.Sharp_Ratio = results08.Risk_Premium ./ sqrt(results08.Variance_Ri);
results08.Treynor = results08.Risk_Premium ./ results08.Beta;

% Création des portefeuilles pour 1998-2007
[~, sorted_Sharp_Ratio_98] = sort(results98.Sharp_Ratio, 'descend'); % Tri décroissant pour Sharp_Ratio
[~, sorted_risk_98] = sort(results98.Ris_Sys_Ajustee, 'descend'); % Tri décroissant par risque systématique
portefeuille_alpha_98 = results98(sorted_Sharp_Ratio_98(1:6), :); % Meilleurs Sharp_Ratio
portefeuille_beta_98 = results98(sorted_Sharp_Ratio_98(end-5:end), :); % Plus bas Sharp_Ratio
portefeuille_segmented_98 = results98(sorted_risk_98(1:6), :); % Segmentés
portefeuille_integrated_98 = results98(sorted_risk_98(end-5:end), :); % Intégrés

% Création des portefeuilles pour 2008-2020
[~, sorted_Sharp_Ratio_08] = sort(results08.Sharp_Ratio, 'descend'); % Tri décroissant pour Sharp_Ratio
[~, sorted_risk_08] = sort(results08.Ris_Sys_Ajustee, 'descend'); % Tri décroissant par risque systématique
portefeuille_delta_08 = results08(sorted_Sharp_Ratio_08(1:6), :); % Meilleurs Sharp_Ratio
portefeuille_gamma_08 = results08(sorted_Sharp_Ratio_08(end-5:end), :); % Plus bas Sharp_Ratio
portefeuille_segmented_08 = results08(sorted_risk_08(1:6), :); % Segmentés
portefeuille_integrated_08 = results08(sorted_risk_08(end-5:end), :); % Intégrés

% Exportation des portefeuilles pour 1998-2007
writetable(portefeuille_alpha_98, 'Portefeuille_Alpha_1998_2007.xlsx');
writetable(portefeuille_beta_98, 'Portefeuille_Beta_1998_2007.xlsx');
writetable(portefeuille_segmented_98, 'Portefeuille_Segmentes_1998_2007.xlsx');
writetable(portefeuille_integrated_98, 'Portefeuille_Integres_1998_2007.xlsx');

% Exportation des portefeuilles pour 2008-2020
writetable(portefeuille_delta_08, 'Portefeuille_Delta_2008_2020.xlsx');
writetable(portefeuille_gamma_08, 'Portefeuille_Gamma_2008_2020.xlsx');
writetable(portefeuille_segmented_08, 'Portefeuille_Segmentes_2008_2020.xlsx');
writetable(portefeuille_integrated_08, 'Portefeuille_Integres_2008_2020.xlsx');

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Garder uniquement les colonnes demandées : pays, période, type, rendement moyen, alpha, risque systématique, sharp ratio
columns_to_keep = {'Country', 'Periode', 'Type', 'rendement_moyen', 'Alpha', 'Ris_Sys_Ajustee', 'Sharp_Ratio'};

% Ajouter une colonne "Type" et "Période" pour identifier chaque portefeuille
portefeuille_alpha_98.Type = repmat("Alpha", height(portefeuille_alpha_98), 1);
portefeuille_alpha_98.Periode = repmat("1998-2007", height(portefeuille_alpha_98), 1);

portefeuille_beta_98.Type = repmat("Beta", height(portefeuille_beta_98), 1);
portefeuille_beta_98.Periode = repmat("1998-2007", height(portefeuille_beta_98), 1);

portefeuille_delta_08.Type = repmat("Delta", height(portefeuille_delta_08), 1);
portefeuille_delta_08.Periode = repmat("2008-2020", height(portefeuille_delta_08), 1);

portefeuille_gamma_08.Type = repmat("Gamma", height(portefeuille_gamma_08), 1);
portefeuille_gamma_08.Periode = repmat("2008-2020", height(portefeuille_gamma_08), 1);

portefeuille_segmented_98.Type = repmat("Segmenté", height(portefeuille_segmented_98), 1);
portefeuille_segmented_98.Periode = repmat("1998-2007", height(portefeuille_segmented_98), 1);

portefeuille_integrated_98.Type = repmat("Intégré", height(portefeuille_integrated_98), 1);
portefeuille_integrated_98.Periode = repmat("1998-2007", height(portefeuille_integrated_98), 1);

portefeuille_segmented_08.Type = repmat("Segmenté", height(portefeuille_segmented_08), 1);
portefeuille_segmented_08.Periode = repmat("2008-2020", height(portefeuille_segmented_08), 1);

portefeuille_integrated_08.Type = repmat("Intégré", height(portefeuille_integrated_08), 1);
portefeuille_integrated_08.Periode = repmat("2008-2020", height(portefeuille_integrated_08), 1);

% Conserver uniquement les colonnes demandées
portefeuille_alpha_98 = portefeuille_alpha_98(:, columns_to_keep);
portefeuille_beta_98 = portefeuille_beta_98(:, columns_to_keep);
portefeuille_delta_08 = portefeuille_delta_08(:, columns_to_keep);
portefeuille_gamma_08 = portefeuille_gamma_08(:, columns_to_keep);

portefeuille_segmented_98 = portefeuille_segmented_98(:, columns_to_keep);
portefeuille_integrated_98 = portefeuille_integrated_98(:, columns_to_keep);
portefeuille_segmented_08 = portefeuille_segmented_08(:, columns_to_keep);
portefeuille_integrated_08 = portefeuille_integrated_08(:, columns_to_keep);

% Combiner les portefeuilles Alpha et Beta pour 1998, Gamma et Delta pour 2008
portefeuilles_sharpe_98 = [
    portefeuille_alpha_98;
    portefeuille_beta_98;    
];

portefeuilles_sharpe_08 = [
    portefeuille_delta_08;
    portefeuille_gamma_08   
];

% Combiner les portefeuilles Segmentés et Intégrés
portefeuilles_segmented_integrated_98 = [
    portefeuille_segmented_98;
    portefeuille_integrated_98;
];
portefeuilles_segmented_integrated_08 = [
    portefeuille_segmented_08;
    portefeuille_integrated_08
];

% Exporter les portefeuilles combinés avec uniquement les colonnes demandées
writetable(portefeuilles_sharpe_98, 'Portefeuilles_Sharpe_Combinés_98.xlsx');
writetable(portefeuilles_segmented_integrated_98, 'Portefeuilles_Segmentés_Intégrés_Combinés_98.xlsx');

writetable(portefeuilles_sharpe_08, 'Portefeuilles_Sharpe_Combinés_08.xlsx');
writetable(portefeuilles_segmented_integrated_08, 'Portefeuilles_Segmentés_Intégrés_Combinés_08.xlsx');

disp('Portefeuilles créés et exportés avec uniquement les colonnes demandées.');
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Classement des pays selon le risque systématique ajusté
[~, risk_rank_98] = sort(results98.Ris_Sys_Ajustee, 'ascend'); % Ascendant pour faible risque systématique
[~, risk_rank_08] = sort(results08.Ris_Sys_Ajustee, 'ascend');

% Classement des pays selon le ratio de Sharpe
[~, sharpe_rank_98] = sort(results98.Sharp_Ratio, 'descend'); % Descendant pour meilleur ratio de Sharpe
[~, sharpe_rank_08] = sort(results08.Sharp_Ratio, 'descend');

% Ajouter les colonnes de classement
results98.Rank_Risk = nan(height(results98), 1);
results98.Rank_Sharpe = nan(height(results98), 1);
results98.Rank_Risk(risk_rank_98) = 1:height(results98); % Classement par risque systématique ajusté
results98.Rank_Sharpe(sharpe_rank_98) = 1:height(results98); % Classement par ratio de Sharpe

results08.Rank_Risk = nan(height(results08), 1);
results08.Rank_Sharpe = nan(height(results08), 1);
results08.Rank_Risk(risk_rank_08) = 1:height(results08);
results08.Rank_Sharpe(sharpe_rank_08) = 1:height(results08);

% Exporter les résultats avec les classements
writetable(results98, 'Classements_Pays_1998_2007.xlsx');
writetable(results08, 'Classements_Pays_2008_2020.xlsx');



disp('Portefeuilles créés et exportés avec succès.');

% --- Définir la fonction de traitement ---
function results = process_sheet(file, sheet)
    % Charger les données depuis le fichier et la feuille spécifiée
    data = readtable(file, 'Sheet', sheet);
    data = rmmissing(data); % Supprimer les lignes avec des valeurs manquantes
    
    % Nom de la variable cible
    Y = data.RendementsMSCI;
    country_columns = data.Properties.VariableNames(3:end); % Colonnes des pays
    
    % Initialisation des résultats
    results = table('Size', [length(country_columns), 10], ...
                    'VariableTypes', {'string', 'double', 'double', 'double', 'double', 'double', 'double', 'double', 'double', 'double'}, ...
                    'VariableNames', {'Country', 'R_Square', 'Adj_R_Square', 'Std_Error', 'F_Stat', 'P_Value', 'Beta', 'Alpha', 'Variance_Ri', 'rendement_moyen'});
    
    for i = 1:length(country_columns)
        X = data{:, country_columns{i}}; % Données pour un pays
        rendement_moyen = mean(X);
        beta = cov(X, Y) / var(X); beta = beta(2);
        alpha = mean(Y) - beta * rendement_moyen;
        Variance_Ri = var(X);
        X_with_intercept = [ones(size(X, 1), 1), X];
        [~, ~, ~, ~, stats] = regress(Y, X_with_intercept);
        R_Square = stats(1);
        F_Stat = stats(2);
        P_Value = stats(3);
        Std_Error = sqrt(stats(4));
        n = length(Y); k = 1;
        Adj_R_Square = 1 - ((1 - R_Square) * (n - 1) / (n - k - 1));
        
        results.Country(i) = string(country_columns{i});
        results.R_Square(i) = R_Square;
        results.Adj_R_Square(i) = Adj_R_Square;
        results.Std_Error(i) = Std_Error;
        results.F_Stat(i) = F_Stat;
        results.P_Value(i) = P_Value;
        results.Beta(i) = beta;
        results.Alpha(i) = alpha;
        results.Variance_Ri(i) = Variance_Ri;
        results.rendement_moyen(i) = rendement_moyen;
        
    end

end

