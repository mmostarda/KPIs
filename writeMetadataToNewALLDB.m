function outputFilePath = writeMetadataToNewALLDB(LoLa_excel_path, metadataStruct)
% writeMetadataToNewALLDB
% Create a copy of a consolidated LoLa Excel file and write GUI metadata
% into every row, with strict non-overwrite rules.
%
% INPUTS
%   LoLa_excel_path : char/string
%       Full path of the consolidated Excel created previously (e.g. ALL_DB.xlsx)
%   metadataStruct  : struct
%       Fields = metadata column names to write (e.g. from GUI mappings).
%
% RULES
%   - DO NOT overwrite column 'Repetition_ID' (keep Excel values)
%   - DO NOT overwrite column 'FileName'      (keep Excel values)
%   - DO NOT overwrite column 'StartSpeed' if the existing cell is numeric
%     (numeric = number or text convertible to number). Only fill when NOT numeric.
%   - DO NOT write to DB: only modify a COPY of LoLa_excel_path
%
% OUTPUT
%   outputFilePath : char
%       Full path of the generated copy: <same folder>\Metadata_ALL_DB.<ext>
%
% NOTES
%   - Uses readtable/writetable. If your MATLAB is old and cannot overwrite sheet
%     without destroying other sheets, this function overwrites the first sheet only
%     when possible; otherwise it rewrites the file (may drop other sheets).

    % -------------------------
    % 0) Validate input
    % -------------------------
    if nargin < 2 || isempty(metadataStruct) || ~isstruct(metadataStruct)
        error('metadataStruct must be a non-empty struct.');
    end
    if nargin < 1 || isempty(LoLa_excel_path) || ~isfile(LoLa_excel_path)
        error('LoLa_excel_path is empty or does not exist: %s', string(LoLa_excel_path));
    end

    % -------------------------
    % 1) Prepare output path and copy file
    % -------------------------
    [inputFolder, ~, inputExt] = fileparts(LoLa_excel_path);
    if isempty(inputExt)
        inputExt = '.xlsx';
    end
    outputFilePath = fullfile(inputFolder, ['Metadata_ALL_DB' inputExt]);

    if isfile(outputFilePath)
        delete(outputFilePath);
    end
    copyfile(LoLa_excel_path, outputFilePath);

    % -------------------------
    % 2) Choose the sheet that contains the consolidated table
    %    (By default: first sheet)
    % -------------------------
    dataSheetName = 'Sheet1';
    try
        sheetList = sheetnames(outputFilePath);
        if ~isempty(sheetList)
            dataSheetName = sheetList(1);
        end
    catch
        % Older MATLAB: keep default 'Sheet1'
    end

    % -------------------------
    % 3) Read table from copied file
    % -------------------------
    T = readtable(outputFilePath, ...
        'Sheet', dataSheetName, ...
        'ReadVariableNames', true, ...
        'VariableNamingRule', 'modify');

    if isempty(T) || height(T) == 0
        error('The Excel file has no rows to update.');
    end

    % -------------------------
    % 4) Define protected columns and remove them from metadataStruct
    % -------------------------
    protectedColumns = {'Repetition_ID', 'FileName'};

    % Drop protected fields if user provided them (case-sensitive remove only if exact)
    for i = 1:numel(protectedColumns)
        if isfield(metadataStruct, protectedColumns{i})
            metadataStruct = rmfield(metadataStruct, protectedColumns{i});
        end
    end

    % -------------------------
    % 5) Apply metadata to all rows
    % -------------------------
    metaFields = fieldnames(metadataStruct);
    rowCount   = height(T);

    for i = 1:numel(metaFields)
        colName = metaFields{i};

        % Extra safety: never touch protected columns (case-insensitive)
        if any(strcmpi(colName, protectedColumns))
            continue;
        end

        valueToWrite = metadataStruct.(colName);

        % Ensure the column exists; create if missing
        if ~ismember(colName, T.Properties.VariableNames)
            T.(colName) = expandScalarToColumn(valueToWrite, rowCount);
            continue;
        end

        % Special rule for StartSpeed
        if strcmpi(colName, 'StartSpeed')
            existingCol = T.(colName);

            numericMask = isNumericValueColumn(existingCol); % true where numeric already
            fillCol     = expandScalarToColumn(valueToWrite, rowCount);

            % Write only where NOT numeric
            T.(colName) = assignByMask(existingCol, fillCol, ~numericMask);
        else
            % Normal behavior: overwrite entire column
            T.(colName) = expandScalarToColumn(valueToWrite, rowCount);
        end
    end

    % -------------------------
    % 6) Write updated table back to the COPY
    % -------------------------
    try
        % Newer MATLAB: overwrite only the target sheet
        writetable(T, outputFilePath, ...
            'Sheet', dataSheetName, ...
            'WriteVariableNames', true, ...
            'WriteMode', 'overwritesheet');
    catch
        % Fallback: rewrite entire file (may drop other sheets)
        delete(outputFilePath);
        writetable(T, outputFilePath, ...
            'WriteVariableNames', true);
    end

    % =========================
    % Local helper functions
    % =========================

    function col = expandScalarToColumn(scalarValue, nRows)
        % Convert a scalar (or compatible vector) into an nRows column.
        % Excel-friendly behavior: chars -> string; unknown -> string.

        if iscell(scalarValue)
            if isscalar(scalarValue)
                col = repmat(scalarValue, nRows, 1);
            else
                if numel(scalarValue) == nRows
                    col = scalarValue(:);
                else
                    error('Cell metadata value has incompatible size.');
                end
            end
            return;
        end

        if isstring(scalarValue)
            if isscalar(scalarValue)
                col = repmat(scalarValue, nRows, 1);
            else
                if numel(scalarValue) == nRows
                    col = scalarValue(:);
                else
                    error('String metadata value has incompatible size.');
                end
            end
            return;
        end

        if ischar(scalarValue)
            col = repmat(string(scalarValue), nRows, 1);
            return;
        end

        if isnumeric(scalarValue) || islogical(scalarValue)
            if isscalar(scalarValue)
                col = repmat(scalarValue, nRows, 1);
            else
                if numel(scalarValue) == nRows
                    col = scalarValue(:);
                else
                    error('Numeric metadata value has incompatible size.');
                end
            end
            return;
        end

        % Fallback
        col = repmat(string(scalarValue), nRows, 1);
    end

    function numericMask = isNumericValueColumn(col)
        % Return logical mask where col contains numeric values already.
        % Numeric if:
        % - numeric type and not NaN
        % - string/char/cell convertible to number (str2double not NaN)

        if isnumeric(col)
            numericMask = ~isnan(col);
            return;
        end

        if isstring(col)
            numericMask = ~isnan(str2double(col));
            return;
        end

        if iscell(col)
            numericMask = false(size(col, 1), 1);
            for r = 1:size(col, 1)
                v = col{r};
                if isnumeric(v) && isscalar(v) && ~isnan(v)
                    numericMask(r) = true;
                elseif ischar(v) || (isstring(v) && isscalar(v))
                    numericMask(r) = ~isnan(str2double(string(v)));
                else
                    numericMask(r) = false;
                end
            end
            return;
        end

        % Anything else -> treat as non-numeric
        numericMask = false(size(col, 1), 1);
    end

    function outCol = assignByMask(outCol, inCol, mask)
        % Assign inCol into outCol where mask is true.
        % Handles numeric/string/cell combinations.

        if iscell(outCol) || iscell(inCol)
            if ~iscell(outCol)
                if isstring(outCol)
                    outCol = cellstr(outCol);
                else
                    outCol = num2cell(outCol);
                end
            end
            if ~iscell(inCol)
                if isstring(inCol)
                    inCol = cellstr(inCol);
                else
                    inCol = num2cell(inCol);
                end
            end
            outCol(mask) = inCol(mask);
        else
            outCol(mask) = inCol(mask);
        end
    end

end