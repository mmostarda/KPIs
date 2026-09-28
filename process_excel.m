%% process_excel.m
% Legge uno o piu' file Excel selezionati tramite UI.
%
% Struttura attesa di ogni file:
%   - Riga 1: intestazioni (la colonna A ha header vuoto)
%   - Colonna A: numero progressivo = Repetition_ID (nessun header)
%   - Righe completamente vuote: separatori di gruppo (reset Repetition_ID)
%   - Righe con colonna A vuota/non-numerica: ignorate
%
% Il mapping LoLa_Name->DB_Name decide quali colonne tenere e come rinominarle.
% La colonna A viene rinominata "Repetition_ID" nel file di output.
% Le righe con FileName vuoto vengono eliminate.
%
% Output per ogni file: stesso path del file di input, con suffisso "_DB".
% Output unificato:     stessa cartella del primo file, nome "ALL_DB.xlsx".
%
% USO:
%   process_excel()           % dizionario interno di default
%   process_excel(mapping)    % struct mapping.LoLa_Name / mapping.DB_Name

function all_output = process_excel(mapping)

    % ------------------------------------------------------------------ %
    %  1. SELEZIONE FILE DI INPUT TRAMITE UI (selezione multipla)        %
    % ------------------------------------------------------------------ %
    [fnames, fpath] = uigetfile( ...
        {'*.xlsx;*.xlsm;*.xls', 'File Excel (*.xlsx, *.xlsm, *.xls)'; ...
         '*.*', 'Tutti i file (*.*)'}, ...
        'Seleziona uno o piu'' file Excel di input', ...
        'MultiSelect', 'on');

    if isequal(fnames, 0)
        fprintf('Operazione annullata dall''utente.\n');
        return;
    end

    % Normalizza sempre a cell array anche se selezionato un solo file
    if ischar(fnames)
        fnames = {fnames};
    end

    n_files = numel(fnames);
    fprintf('File selezionati: %d\n', n_files);

    % ------------------------------------------------------------------ %
    %  2. DIZIONARIO DI MAPPATURA                                        %
    % ------------------------------------------------------------------ %
    if nargin < 1 || isempty(mapping)
        mapping.LoLa_Name = {'File Name', 'Column_B', 'Column_C'};
        mapping.DB_Name   = {'FileName',  'DB_B',     'DB_C'};
    end
    assert(numel(mapping.LoLa_Name) == numel(mapping.DB_Name), ...
        'mapping.LoLa_Name e mapping.DB_Name devono avere la stessa lunghezza.');

    % ------------------------------------------------------------------ %
    %  LOOP SUI FILE SELEZIONATI                                         %
    % ------------------------------------------------------------------ %
    T_all = [];   % tabella cumulativa per il file unificato

    for file_i = 1:n_files
        fname       = fnames{file_i};
        input_file  = fullfile(fpath, fname);
        [~, nonly, ext] = fileparts(fname);
        output_file = fullfile(fpath, [nonly '_DB' ext]);

        fprintf('\n========================================\n');
        fprintf('File %d/%d: %s\n', file_i, n_files, fname);
        fprintf('Output:  %s\n', output_file);
        fprintf('========================================\n');

        % Processa il singolo file e ottieni la tabella risultato
        T = process_single_file(input_file, mapping);

        if isempty(T)
            fprintf('[ATTENZIONE] File %s: nessuna riga valida, saltato.\n', fname);
            continue;
        end

        % Salva file individuale _DB
        writetable(T, output_file, 'WriteVariableNames', true);
        fprintf('File individuale salvato: %s\n', output_file);
        fprintf('Righe: %d  |  Colonne: %d\n', height(T), width(T));

        % Accumula per il file unificato
        if isempty(T_all)
            T_all = T;
        else
            % Allinea colonne: aggiungi colonne mancanti come vuote
            T_all = merge_tables(T_all, T);
        end
    end

    % ------------------------------------------------------------------ %
    %  SALVATAGGIO FILE UNIFICATO ALL_DB                                 %
    % ------------------------------------------------------------------ %
    if ~isempty(T_all)
        all_output = fullfile(fpath, 'ALL_DB.xlsx');
        writetable(T_all, all_output, 'WriteVariableNames', true);
        fprintf('\n========================================\n');
        fprintf('File unificato salvato: %s\n', all_output);
        fprintf('Righe totali: %d  |  Colonne: %d\n', height(T_all), width(T_all));
        fprintf('========================================\n');
    else
        fprintf('\n[ATTENZIONE] Nessun dato valido trovato in nessun file.\n');
    end
end

% ================================================================== %
%  FUNZIONE INTERNA: processa un singolo file, restituisce tabella   %
% ================================================================== %
function T = process_single_file(input_file, mapping)

    T = [];

    % ------------------------------------------------------------------ %
    %  3. LETTURA COMPLETA DEL FOGLIO CON readcell                       %
    %     readcell preserva i tipi originali senza rinominare le colonne %
    % ------------------------------------------------------------------ %
    fprintf('Lettura file...\n');
    try
        all_raw = readcell(input_file, 'TextType', 'char', 'UseExcel', true);
    catch
        all_raw = readcell(input_file, 'TextType', 'char');
    end

    n_cols     = size(all_raw, 2);
    n_rows_tot = size(all_raw, 1);   % include riga header

    % ------------------------------------------------------------------ %
    %  4. ESTRAZIONE HEADER RAW (riga 1)                                 %
    % ------------------------------------------------------------------ %
    header_raw = cell(1, n_cols);
    for c = 1:n_cols
        header_raw{c} = cell2str_safe(all_raw{1, c});
    end
    fprintf('Colonne trovate nel file: %d\n', n_cols);

    % ------------------------------------------------------------------ %
    %  5. ANALISI RIGHE DATI (righe 2..end)                              %
    %     - identifica righe vuote (separatori di gruppo)                %
    %     - identifica righe con numero in colonna A                     %
    % ------------------------------------------------------------------ %
    n_data = n_rows_tot - 1;

    is_empty_row  = false(n_data, 1);
    is_numeric_A  = false(n_data, 1);
    rep_id_raw    = zeros(n_data, 1);

    for r = 1:n_data
        row = all_raw(r+1, :);

        % Riga vuota = tutte le celle sono vuote/missing/NaN
        is_empty_row(r) = all(cellfun(@(x) isempty(cell2str_safe(x)) || ...
                                           strcmpi(cell2str_safe(x),'NaN'), row));

        % Colonna A numerica
        val_A = all_raw{r+1, 1};
        if isnumeric(val_A) && ~isnan(val_A)
            is_numeric_A(r) = true;
            rep_id_raw(r)   = val_A;
        elseif ischar(val_A)
            num = str2double(strtrim(val_A));
            if ~isnan(num)
                is_numeric_A(r) = true;
                rep_id_raw(r)   = num;
            end
        end
    end

    fprintf('Righe totali dati: %d\n', n_data);
    fprintf('Righe con numero in colonna A: %d\n', sum(is_numeric_A));

    % ------------------------------------------------------------------ %
    %  6. TROVA INDICE FISICO DELLA COLONNA "File Name"                  %
    % ------------------------------------------------------------------ %
    fn_excel_col = [];
    for c = 1:n_cols
        if strcmpi(strtrim(header_raw{c}), 'File Name')
            fn_excel_col = c;
            break;
        end
    end
    if isempty(fn_excel_col)
        % Cerca anche tramite mapping (LoLa_Name corrispondente a DB FileName)
        for m = 1:numel(mapping.LoLa_Name)
            if strcmpi(mapping.DB_Name{m}, 'FileName')
                for c = 1:n_cols
                    if strcmpi(strtrim(header_raw{c}), strtrim(mapping.LoLa_Name{m}))
                        fn_excel_col = c;
                        break;
                    end
                end
            end
            if ~isempty(fn_excel_col), break; end
        end
    end

    if isempty(fn_excel_col)
        fprintf('[ATTENZIONE] Colonna "File Name" non trovata: FileName sara'' vuoto.\n');
    else
        fprintf('Colonna "File Name" trovata alla colonna fisica %d ("%s").\n', ...
                fn_excel_col, col_idx_to_letter(fn_excel_col));
    end

    % ------------------------------------------------------------------ %
    %  7. COSTRUZIONE TABELLA RISULTATO                                  %
    %     Solo colonne presenti nel mapping, lette dalla posizione fisica %
    % ------------------------------------------------------------------ %
    % Mappa header_raw -> indice fisico colonna
    col_map = containers.Map('KeyType','char','ValueType','double');
    for c = 1:n_cols
        key = lower(strtrim(header_raw{c}));
        if ~isempty(key) && ~col_map.isKey(key)
            col_map(key) = c;
        end
    end

    % Seleziona colonne da tenere (secondo mapping)
    keep_excel_idx = [];   % indici fisici colonne da tenere
    keep_db_names  = {};   % nomi DB finali

    for m = 1:numel(mapping.LoLa_Name)
        lola = lower(strtrim(mapping.LoLa_Name{m}));
        if col_map.isKey(lola)
            keep_excel_idx(end+1) = col_map(lola); %#ok<AGROW>
            keep_db_names{end+1}  = mapping.DB_Name{m}; %#ok<AGROW>
        else
            fprintf('  [ATTENZIONE] LoLa_Name "%s" non trovato nel file.\n', ...
                    mapping.LoLa_Name{m});
        end
    end

    fprintf('Colonne mantenute dal mapping: %d\n', numel(keep_excel_idx));

    % Righe da includere: colonna A numerica
    data_row_idx = find(is_numeric_A);   % indici in dati (1..n_data)

    % ------------------------------------------------------------------ %
    %  8. LEGGI FILE NAME RAW CON FILL-DOWN (gestisce merged cells)      %
    % ------------------------------------------------------------------ %
    fn_col_data = repmat({''}, n_data, 1);
    if ~isempty(fn_excel_col)
        last_valid_fn = '';
        for r = 1:n_data
            raw_val = cell2str_safe(all_raw{r+1, fn_excel_col});
            % Valore valido = stringa non vuota, non numerica pura, non 'NaN'
            if ~isempty(raw_val) && isnan(str2double(raw_val)) && ...
               ~strcmpi(raw_val, 'NaN')
                last_valid_fn  = raw_val;
                fn_col_data{r} = raw_val;
            elseif ~is_empty_row(r)
                % Riga non vuota ma senza valore -> fill-down
                fn_col_data{r} = last_valid_fn;
            else
                % Riga separatore vuota -> reset fill-down
                last_valid_fn  = '';
                fn_col_data{r} = '';
            end
        end
    end

    % ------------------------------------------------------------------ %
    %  9. ASSEMBLA TABELLA FINALE                                        %
    % ------------------------------------------------------------------ %
    n_out = numel(data_row_idx);
    out_data = cell(n_out, numel(keep_excel_idx));

    for col_i = 1:numel(keep_excel_idx)
        ec = keep_excel_idx(col_i);
        db = keep_db_names{col_i};

        for row_i = 1:n_out
            r = data_row_idx(row_i);

            if strcmpi(db, 'FileName')
                % Usa il valore con fill-down
                out_data{row_i, col_i} = fn_col_data{r};
            elseif is_date_column(db)
                % Colonna data: usa conversione sicura (mai data di oggi)
                out_data{row_i, col_i} = cell2date_safe(all_raw{r+1, ec});
            else
                out_data{row_i, col_i} = cell2str_safe(all_raw{r+1, ec});
            end
        end
    end

    T = cell2table(out_data, 'VariableNames', keep_db_names);

    % ------------------------------------------------------------------ %
    %  10. AGGIUNGI Repetition_ID dalla colonna A originale              %
    % ------------------------------------------------------------------ %
    T.Repetition_ID = rep_id_raw(data_row_idx);
    % Sposta Repetition_ID come prima colonna
    T = [T(:,'Repetition_ID'), T(:,setdiff(T.Properties.VariableNames, {'Repetition_ID'}, 'stable'))];

    % ------------------------------------------------------------------ %
    %  11. ELIMINA RIGHE CON FileName VUOTO                              %
    % ------------------------------------------------------------------ %
    fn_col_name = '';
    for cn = T.Properties.VariableNames
        if strcmpi(cn{1}, 'FileName')
            fn_col_name = cn{1}; break;
        end
    end

    if ~isempty(fn_col_name)
        fn_vals     = T.(fn_col_name);
        if ~iscell(fn_vals), fn_vals = cellstr(string(fn_vals)); end
        valid_rows  = ~cellfun(@isempty, fn_vals);
        T           = T(valid_rows, :);
        fprintf('Righe dopo eliminazione FileName vuoto: %d\n', height(T));
    end

    % Sezione 12: T viene restituita al chiamante (process_excel)
    % La scrittura avviene nel loop principale.
end

% ================================================================== %
%  HELPER: unisce due tabelle allineando le colonne                  %
%  Colonne mancanti in una delle due vengono aggiunte come ''        %
% ================================================================== %
function T_out = merge_tables(T1, T2)
    cols1 = T1.Properties.VariableNames;
    cols2 = T2.Properties.VariableNames;

    % Aggiungi a T1 le colonne presenti in T2 ma non in T1
    for c = cols2
        if ~any(strcmpi(cols1, c{1}))
            T1.(c{1}) = repmat({''}, height(T1), 1);
        end
    end
    % Aggiungi a T2 le colonne presenti in T1 ma non in T2
    for c = T1.Properties.VariableNames
        if ~any(strcmpi(cols2, c{1}))
            T2.(c{1}) = repmat({''}, height(T2), 1);
        end
    end
    % Riordina T2 con lo stesso ordine colonne di T1
    T2    = T2(:, T1.Properties.VariableNames);
    T_out = [T1; T2];
end

% ------------------------------------------------------------------ %
%  HELPER: indice colonna -> lettera Excel (1='A', 27='AA', ...)      %
% ------------------------------------------------------------------ %
function letter = col_idx_to_letter(idx)
    letter = '';
    while idx > 0
        r      = mod(idx-1, 26);
        letter = [char(r+65) letter]; %#ok<AGROW>
        idx    = floor((idx-1)/26);
    end
end

% ------------------------------------------------------------------ %
%  HELPER: determina se un nome colonna e' una colonna data           %
%  Criteri: contiene 'date' o 'data' (case-insensitive), oppure       %
%  e' in un elenco di nomi noti.                                      %
% ------------------------------------------------------------------ %
function result = is_date_column(col_name)
    name_lower = lower(strtrim(col_name));
    % Considera colonna data se il nome contiene 'date' o 'data'
    result = ~isempty(regexp(name_lower, 'date|_date|data|_data', 'once'));
end

% ------------------------------------------------------------------ %
%  HELPER: cella raw -> stringa sicura (gestisce missing/NaN/numeric) %
%  I valori datetime vengono convertiti in stringa 'yyyy-MM-dd'.      %
%  Valori non validi (NA, missing, NaN, datetime non valido) -> ''.   %
% ------------------------------------------------------------------ %
function s = cell2str_safe(x)
    if ischar(x)
        s = strtrim(x);
    elseif isstring(x) && isscalar(x) && ~ismissing(x)
        s = strtrim(char(x));
    elseif isdatetime(x) && isscalar(x)
        if isnat(x)
            s = '';   % NaT -> cella vuota
        else
            s = char(datetime(x, 'Format', 'yyyy-MM-dd'));
        end
    elseif isnumeric(x) && isscalar(x) && ~isnan(x)
        s = num2str(x);
    else
        % Copre: NaN, <missing>, empty, datetime vettoriale, tipi non noti
        s = '';
    end
end

% ------------------------------------------------------------------ %
%  HELPER: valore cella -> stringa data sicura                        %
%  Usato per colonne che nel mapping sono identificate come date.     %
%  Regola: se il valore non e' una data riconoscibile -> '' vuoto.   %
%  NON inserisce mai la data di oggi come fallback.                   %
% ------------------------------------------------------------------ %
function s = cell2date_safe(x)
    % Placeholder noti da trattare come vuoto
    invalid_strings = {'na', 'n/a', 'nan', 'nat', '-', '', ...
                       '<missing>', '<undefined>', 'none', 'null'};

    % Caso 1: datetime MATLAB
    if isdatetime(x) && isscalar(x)
        if isnat(x)
            s = ''; return;
        else
            s = char(datetime(x, 'Format', 'yyyy-MM-dd')); return;
        end
    end

    % Caso 2: numero seriale Excel (giorni dal 30-dic-1899)
    if isnumeric(x) && isscalar(x) && ~isnan(x)
        if x >= 1 && x <= 2958465   % range valido: 1900..9999
            try
                dt = datetime(1899,12,30) + days(x);
                s  = char(datetime(dt, 'Format', 'yyyy-MM-dd'));
            catch
                s = '';
            end
        else
            s = '';   % numero fuori range -> non e' una data
        end
        return;
    end

    % Caso 3: stringa
    if ischar(x) || (isstring(x) && isscalar(x) && ~ismissing(x))
        raw = strtrim(lower(char(x)));

        % Stringa vuota o placeholder noto -> vuoto
        if any(strcmpi(raw, invalid_strings))
            s = ''; return;
        end

        % Tenta il parsing nei formati piu' comuni
        formats = {'yyyy-MM-dd','dd/MM/yyyy','MM/dd/yyyy', ...
                   'dd-MM-yyyy','yyyy/MM/dd','dd.MM.yyyy', ...
                   'yyyyMMdd','dd MMM yyyy','MMM dd, yyyy'};
        for k = 1:numel(formats)
            try
                dt = datetime(strtrim(char(x)), 'InputFormat', formats{k});
                s  = char(datetime(dt, 'Format', 'yyyy-MM-dd'));
                return;
            catch
                % prova prossimo formato
            end
        end

        % Nessun formato riconosciuto -> vuoto (NON data di oggi)
        s = ''; return;
    end

    % Caso 4: qualsiasi altro tipo non gestito
    s = '';
end