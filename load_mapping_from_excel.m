%% load_mapping_from_excel.m
% Legge un file Excel con colonne "KPI_DB_Name" e "KPI_LoLa_Name",
% costruisce il dizionario di mappatura saltando righe con valori "NA" o vuoti,
% e salva il risultato in un file .mat nella stessa cartella del file Excel.
%
% USO:
%   mapping = load_mapping_from_excel()
%
% OUTPUT:
%   mapping.LoLa_Name  : cell array di stringhe (nomi sorgente)
%   mapping.DB_Name    : cell array di stringhe (nomi destinazione)
%   Salva anche <nome_file>_mapping.mat nella cartella del file selezionato.

function mapping = load_mapping_from_excel()

    mapping = [];

    % ------------------------------------------------------------------ %
    %  1. SELEZIONE FILE TRAMITE UI                                       %
    % ------------------------------------------------------------------ %
    [fname, fpath] = uigetfile( ...
        {'*.xlsx;*.xlsm;*.xls', 'File Excel (*.xlsx, *.xlsm, *.xls)'; ...
         '*.*', 'Tutti i file (*.*)'}, ...
        'Seleziona il file Excel con il mapping KPI');

    if isequal(fname, 0)
        fprintf('Operazione annullata dall''utente.\n');
        return;
    end

    input_file = fullfile(fpath, fname);
    fprintf('File selezionato: %s\n', input_file);

    % ------------------------------------------------------------------ %
    %  2. LETTURA FILE EXCEL                                              %
    % ------------------------------------------------------------------ %
    opts = detectImportOptions(input_file, 'NumHeaderLines', 0);
    opts.VariableNamesRange = 'A1';
    opts.DataRange          = 'A2';
    opts = setvaropts(opts, opts.VariableNames, 'Type', 'char');
    T = readtable(input_file, opts, 'ReadVariableNames', true);

    fprintf('Righe lette (esclusa intestazione): %d\n', height(T));

    % ------------------------------------------------------------------ %
    %  3. VERIFICA COLONNE RICHIESTE                                      %
    % ------------------------------------------------------------------ %
    col_names = T.Properties.VariableNames;

    idx_db   = find(strcmpi(strtrim(col_names), 'KPI_DB_Name'),   1);
    idx_lola = find(strcmpi(strtrim(col_names), 'KPI_LoLa_Name'), 1);

    if isempty(idx_db)
        error('Colonna "KPI_DB_Name" non trovata nel file. Verificare le intestazioni.');
    end
    if isempty(idx_lola)
        error('Colonna "KPI_LoLa_Name" non trovata nel file. Verificare le intestazioni.');
    end

    db_col   = T.(col_names{idx_db});
    lola_col = T.(col_names{idx_lola});

    % Normalizza in cell array di char
    if ~iscell(db_col),   db_col   = cellstr(string(db_col));   end
    if ~iscell(lola_col), lola_col = cellstr(string(lola_col)); end

    % ------------------------------------------------------------------ %
    %  4. COSTRUZIONE MAPPING (skip righe NA o vuote)                     %
    % ------------------------------------------------------------------ %
    db_out   = {};
    lola_out = {};
    skipped  = 0;

    for i = 1:numel(db_col)
        db_val   = strtrim(db_col{i});
        lola_val = strtrim(lola_col{i});

        % Salta se uno dei due è vuoto, NaN (da cella numerica vuota) o "NA"
        if is_invalid(db_val) || is_invalid(lola_val)
            skipped = skipped + 1;
            continue;
        end

        db_out{end+1}   = db_val;   %#ok<AGROW>
        lola_out{end+1} = lola_val; %#ok<AGROW>
    end

    fprintf('Coppie valide:  %d\n', numel(db_out));
    fprintf('Righe skippate: %d (vuote o NA)\n', skipped);

    if isempty(db_out)
        error('Nessuna coppia valida trovata nel file. Verificare i dati.');
    end

    % Costruisci struct mapping
    mapping.LoLa_Name = lola_out;
    mapping.DB_Name   = db_out;

    % ------------------------------------------------------------------ %
    %  5. SALVATAGGIO FILE .MAT                                           %
    % ------------------------------------------------------------------ %
    [~, name_only, ~] = fileparts(fname);
    mat_file = fullfile(fpath, [name_only '_mapping.mat']);

    save(mat_file, 'mapping');
    fprintf('Mapping salvato in: %s\n', mat_file);

    % Riepilogo a schermo
    fprintf('\n--- RIEPILOGO MAPPING ---\n');
    fprintf('%-40s  -->  %s\n', 'LoLa_Name', 'DB_Name');
    fprintf('%s\n', repmat('-', 1, 70));
    for i = 1:numel(mapping.LoLa_Name)
        fprintf('%-40s  -->  %s\n', mapping.LoLa_Name{i}, mapping.DB_Name{i});
    end
    fprintf('%s\n', repmat('-', 1, 70));
end

% ------------------------------------------------------------------ %
%  FUNZIONE HELPER: controlla se un valore è da skippare             %
% ------------------------------------------------------------------ %
function result = is_invalid(val)
    % Considera invalido: stringa vuota, solo spazi, "NA" (case-insensitive),
    % oppure "NaN" (risultato di conversione di celle numeriche vuote)
    result = isempty(val) || ...
             strcmpi(val, 'NA')  || ...
             strcmpi(val, 'NaN') || ...
             strcmpi(val, '<undefined>');
end