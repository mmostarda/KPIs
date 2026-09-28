% MetadataKPIApp.m
% Simple UI app to manage Metadata and KPI tables in an Excel file.
% It uses a programmatic UI (uifigure, uitabgroup, etc.).
% You can reuse the logic inside an App Designer (.mlapp) if you prefer.
%
% ASSUMPTIONS ABOUT EXCEL FILE:
% - Row 1: informational header or title (NOT part of the data table)
% - Row 2: actual column headers (ID, IDName, Date, ...)
% - Row 3+: data rows
% - Tables start at column A. If they start at another column, adjust 'Range' (e.g. 'B3').

classdef MetadataKPIApp < matlab.apps.AppBase

    % ---------------------------------------------------------------------
    % PUBLIC UI PROPERTIES
    % ---------------------------------------------------------------------
    properties (Access = public)
        UIFigure                matlab.ui.Figure
        TabGroup                matlab.ui.container.TabGroup
        MetadataTab             matlab.ui.container.Tab
        KPITab                  matlab.ui.container.Tab
        
        % --- Metadata tab components ---
        Metadata_IDLabel        matlab.ui.control.Label
        Metadata_IDField        matlab.ui.control.NumericEditField
        
        Metadata_IDNameLabel    matlab.ui.control.Label
        Metadata_IDNameField    matlab.ui.control.EditField
        
        Metadata_DateLabel      matlab.ui.control.Label
        Metadata_DatePicker     matlab.ui.control.DatePicker
        
        Metadata_ClusterLabel   matlab.ui.control.Label
        Metadata_ClusterField   matlab.ui.control.EditField
        
        Metadata_SubClusterLabel matlab.ui.control.Label
        Metadata_SubClusterField matlab.ui.control.EditField
        
        Metadata_RefFilesLabel  matlab.ui.control.Label
        Metadata_RefFilesField  matlab.ui.control.EditField
        
        Metadata_BrandLabel     matlab.ui.control.Label
        Metadata_BrandField     matlab.ui.control.EditField
        
        Metadata_ModelLabel     matlab.ui.control.Label
        Metadata_ModelField     matlab.ui.control.EditField
        
        Metadata_SaveButton     matlab.ui.control.Button
        Metadata_StatusLabel    matlab.ui.control.Label
        
        % --- KPI tab components ---
        KPI_IDLabel             matlab.ui.control.Label
        KPI_IDDropDown          matlab.ui.control.DropDown
        
        KPI_NameLabel           matlab.ui.control.Label
        KPI_NameDropDown        matlab.ui.control.DropDown
        
        KPI_ValueLabel          matlab.ui.control.Label
        KPI_ValueField          matlab.ui.control.NumericEditField
        
        KPI_SaveButton          matlab.ui.control.Button
        KPI_StatusLabel         matlab.ui.control.Label
    end

    % ---------------------------------------------------------------------
    % PRIVATE DATA PROPERTIES
    % ---------------------------------------------------------------------
    properties (Access = private)
        excelFilePath       string = ""          % Path to Excel file
        metadataSheetName   string = "Metadata"  % Name of Metadata sheet
        kpiSheetName        string = "KPIs"      % Name of KPIs sheet
        
        metadataTable       table                % In-memory metadata table
        kpiTable            table                % In-memory KPI table
    end

    % ---------------------------------------------------------------------
    % PRIVATE METHODS - CORE LOGIC
    % ---------------------------------------------------------------------
    methods (Access = private)
        
        function createComponents(app)
            % Create and configure UI components programmatically
            
            % --- Main figure ---
            app.UIFigure = uifigure(...
                'Name', 'Metadata & KPI Editor', ...
                'Position', [100 100 900 450]);
            
            % --- Tab group ---
            app.TabGroup = uitabgroup(app.UIFigure, ...
                'Position', [1 1 900 450]);
            
            % -----------------------------------------------------------------
            % METADATA TAB
            % -----------------------------------------------------------------
            app.MetadataTab = uitab(app.TabGroup, ...
                'Title', 'Metadata');
            
            % ID (numeric, read-only)
            app.Metadata_IDLabel = uilabel(app.MetadataTab, ...
                'Position', [30 380 60 22], ...
                'Text', 'ID');
            app.Metadata_IDField = uieditfield(app.MetadataTab, 'numeric', ...
                'Position', [120 380 100 22], ...
                'Editable', 'off');   % auto-managed
            
            % IDName
            app.Metadata_IDNameLabel = uilabel(app.MetadataTab, ...
                'Position', [30 340 60 22], ...
                'Text', 'IDName');
            app.Metadata_IDNameField = uieditfield(app.MetadataTab, 'text', ...
                'Position', [120 340 200 22]);
            
            % Date
            app.Metadata_DateLabel = uilabel(app.MetadataTab, ...
                'Position', [30 300 60 22], ...
                'Text', 'Date');
            app.Metadata_DatePicker = uidatepicker(app.MetadataTab, ...
                'Position', [120 300 150 22], ...
                'Value', datetime('today'));
            
            % Cluster
            app.Metadata_ClusterLabel = uilabel(app.MetadataTab, ...
                'Position', [30 260 60 22], ...
                'Text', 'Cluster');
            app.Metadata_ClusterField = uieditfield(app.MetadataTab, 'text', ...
                'Position', [120 260 200 22]);
            
            % SubCluster
            app.Metadata_SubClusterLabel = uilabel(app.MetadataTab, ...
                'Position', [30 220 80 22], ...
                'Text', 'SubCluster');
            app.Metadata_SubClusterField = uieditfield(app.MetadataTab, 'text', ...
                'Position', [120 220 200 22]);
            
            % RefFiles
            app.Metadata_RefFilesLabel = uilabel(app.MetadataTab, ...
                'Position', [30 180 60 22], ...
                'Text', 'RefFiles');
            app.Metadata_RefFilesField = uieditfield(app.MetadataTab, 'text', ...
                'Position', [120 180 300 22]);
            
            % Brand
            app.Metadata_BrandLabel = uilabel(app.MetadataTab, ...
                'Position', [30 140 60 22], ...
                'Text', 'Brand');
            app.Metadata_BrandField = uieditfield(app.MetadataTab, 'text', ...
                'Position', [120 140 200 22]);
            
            % Model
            app.Metadata_ModelLabel = uilabel(app.MetadataTab, ...
                'Position', [30 100 60 22], ...
                'Text', 'Model');
            app.Metadata_ModelField = uieditfield(app.MetadataTab, 'text', ...
                'Position', [120 100 200 22]);
            
            % Save button
            app.Metadata_SaveButton = uibutton(app.MetadataTab, ...
                'Text', 'Save metadata row', ...
                'Position', [30 50 200 30], ...
                'ButtonPushedFcn', @(btn,event)app.onSaveMetadata(event));
            
            % Status label
            app.Metadata_StatusLabel = uilabel(app.MetadataTab, ...
                'Position', [250 50 600 30], ...
                'Text', '', ...
                'FontColor', [0.15 0.15 0.15]);
            
            % -----------------------------------------------------------------
            % KPI TAB
            % -----------------------------------------------------------------
            app.KPITab = uitab(app.TabGroup, ...
                'Title', 'KPIs');
            
            % ID dropdown
            app.KPI_IDLabel = uilabel(app.KPITab, ...
                'Position', [30 350 60 22], ...
                'Text', 'ID');
            app.KPI_IDDropDown = uidropdown(app.KPITab, ...
                'Position', [120 350 150 22], ...
                'Items', {}, ...
                'Editable', 'off');
            
            % KPI name dropdown
            app.KPI_NameLabel = uilabel(app.KPITab, ...
                'Position', [30 310 80 22], ...
                'Text', 'KPI name');
            app.KPI_NameDropDown = uidropdown(app.KPITab, ...
                'Position', [120 310 250 22], ...
                'Items', {}, ...
                'Editable', 'off');
            
            % KPI value
            app.KPI_ValueLabel = uilabel(app.KPITab, ...
                'Position', [30 270 80 22], ...
                'Text', 'KPI value');
            app.KPI_ValueField = uieditfield(app.KPITab, 'numeric', ...
                'Position', [120 270 150 22], ...
                'Value', 0);
            
            % Save KPI button
            app.KPI_SaveButton = uibutton(app.KPITab, ...
                'Text', 'Save KPI value', ...
                'Position', [30 220 200 30], ...
                'ButtonPushedFcn', @(btn,event)app.onSaveKPI(event));
            
            % KPI status label
            app.KPI_StatusLabel = uilabel(app.KPITab, ...
                'Position', [250 220 600 30], ...
                'Text', '', ...
                'FontColor', [0.15 0.15 0.15]);
        end
        
        % -----------------------------------------------------------------
        % STARTUP LOGIC
        % -----------------------------------------------------------------
        function startup(app)
            % Choose Excel file if not provided
            if strlength(app.excelFilePath) == 0
                [fileName, filePath] = uigetfile('*.xlsx', ...
                    'Select Excel file with Metadata & KPIs');
                if isequal(fileName, 0)
                    uialert(app.UIFigure, ...
                        'No Excel file selected. App will be closed.', ...
                        'Error');
                    delete(app);
                    return;
                else
                    app.excelFilePath = string(fullfile(filePath, fileName));
                end
            end
            
            % Load tables from Excel
            app.loadTablesFromExcel();
            
            % Initialize UI fields from tables
            app.updateNextMetadataID();
            app.refreshIDDropDown();
            app.refreshKPINameDropDown();
        end
        
        % -----------------------------------------------------------------
        % HELPER: load Metadata & KPI tables
        % -----------------------------------------------------------------
        function loadTablesFromExcel(app)
            % Load Metadata and KPI tables from the Excel file,
            % assuming that:
            % - Row 1: some info/title (not part of the table)
            % - Row 2: actual column headers
            % - Row 3+: data rows
            
            if ~isfile(app.excelFilePath)
                error('Excel file not found: %s', app.excelFilePath);
            end
            
            % Read Metadata sheet
            try
                app.metadataTable = readtable( ...
                    app.excelFilePath, ...
                    'Sheet', app.metadataSheetName, ...
                    'ReadVariableNames', true, ... % use row after header lines as header
                    'NumHeaderLines', 1);          % skip the first row
            catch ME
                warning('Error reading Metadata sheet: %s', ME.message);
                app.metadataTable = table();
            end
            
            % Read KPIs sheet
            try
                app.kpiTable = readtable( ...
                    app.excelFilePath, ...
                    'Sheet', app.kpiSheetName, ...
                    'ReadVariableNames', true, ... % use row after header lines as header
                    'NumHeaderLines', 1);          % skip the first row
            catch ME
                warning('Error reading KPIs sheet: %s', ME.message);
                app.kpiTable = table();
            end
        end
        
        % -----------------------------------------------------------------
        % HELPER: compute and set next ID for new Metadata row
        % -----------------------------------------------------------------
        function updateNextMetadataID(app)
            % Compute next ID as max(existing IDs)+1 or 1 if table is empty
            
            if ~isempty(app.metadataTable) && any(strcmp('ID', app.metadataTable.Properties.VariableNames))
                % Make sure ID is numeric
                maxID = max(app.metadataTable.ID);
                if isempty(maxID) || isnan(maxID)
                    nextID = 1;
                else
                    nextID = maxID + 1;
                end
            else
                nextID = 1;
            end
            
            app.Metadata_IDField.Value = nextID;
        end
        
        % -----------------------------------------------------------------
        % HELPER: refresh ID dropdown in KPI tab
        % -----------------------------------------------------------------
        function refreshIDDropDown(app)
            % Update the list of IDs available in the KPI tab
            
            ids = [];
            
            if ~isempty(app.kpiTable) && any(strcmp('ID', app.kpiTable.Properties.VariableNames))
                ids = app.kpiTable.ID;
            elseif ~isempty(app.metadataTable) && any(strcmp('ID', app.metadataTable.Properties.VariableNames))
                ids = app.metadataTable.ID;
            end
            
            if isempty(ids)
                app.KPI_IDDropDown.Items = {};
            else
                app.KPI_IDDropDown.Items = string(ids(:)');
                app.KPI_IDDropDown.Value = app.KPI_IDDropDown.Items{1};
            end
        end
        
        % -----------------------------------------------------------------
        % HELPER: refresh KPI-name dropdown in KPI tab
        % -----------------------------------------------------------------
        function refreshKPINameDropDown(app)
            % Use KPI table variable names, excluding ID, IDName, Date
            
            if isempty(app.kpiTable)
                app.KPI_NameDropDown.Items = {};
                return;
            end
            
            varNames = app.kpiTable.Properties.VariableNames;
            mask = ~ismember(varNames, {'ID', 'IDName', 'Date'});
            kpiNames = varNames(mask);
            
            if isempty(kpiNames)
                app.KPI_NameDropDown.Items = {};
            else
                app.KPI_NameDropDown.Items = kpiNames;
                app.KPI_NameDropDown.Value = kpiNames{1};
            end
        end
        
        % -----------------------------------------------------------------
        % CALLBACK: Save metadata row
        % -----------------------------------------------------------------
        function onSaveMetadata(app, event) %#ok<INUSD>
            try
                % Collect values from UI
                newID       = app.Metadata_IDField.Value;
                newIDName   = string(app.Metadata_IDNameField.Value);
                newDate     = app.Metadata_DatePicker.Value;
                newCluster  = string(app.Metadata_ClusterField.Value);
                newSubCl    = string(app.Metadata_SubClusterField.Value);
                newRefFiles = string(app.Metadata_RefFilesField.Value);
                newBrand    = string(app.Metadata_BrandField.Value);
                newModel    = string(app.Metadata_ModelField.Value);
                
                % Build a new Metadata row as a table.
                % NOTE: variable names must match your Excel columns.
                newMetadataRow = table( ...
                    newID, ...
                    newIDName, ...
                    newDate, ...
                    newCluster, ...
                    newSubCl, ...
                    newRefFiles, ...
                    newBrand, ...
                    newModel, ...
                    'VariableNames', { ...
                        'ID', 'IDName', 'Date', ...
                        'Cluster', 'SubCluster', 'RefFiles', ...
                        'Brand', 'Model'});
                
                % Append to metadataTable
                if isempty(app.metadataTable)
                    app.metadataTable = newMetadataRow;
                else
                    app.metadataTable = [app.metadataTable; newMetadataRow];
                end
                
                % Ensure the corresponding row exists in KPI table
                app.ensureKpiRowForID(newID, newIDName, newDate);
                
                % ---- WRITE METADATA (headers + data) ----
                % Headers are in row 2, data from row 3.
                % We overwrite headers + data, but ONLY in the Metadata sheet.
                writetable(app.metadataTable, app.excelFilePath, ...
                    'Sheet', app.metadataSheetName, ...
                    'Range', 'A2', ...              % start writing at row 2
                    'WriteVariableNames', true);    % write headers in row 2
                
                % ---- WRITE KPI DATA ONLY (no headers) ----
                % We do NOT overwrite column names of KPI sheet.
                % Excel:
                %   Row 1: title (untouched)
                %   Row 2: headers (untouched)
                %   Row 3+: KPI data  -> we overwrite/extend only this region.
                writetable(app.kpiTable, app.excelFilePath, ...
                    'Sheet', app.kpiSheetName, ...
                    'Range', 'A3', ...              % first data row
                    'WriteVariableNames', false);   % do NOT write headers
                
                % Update UI
                app.Metadata_StatusLabel.Text = sprintf('Metadata row with ID %d saved.', newID);
                app.updateNextMetadataID();
                app.refreshIDDropDown();
                
            catch ME
                app.Metadata_StatusLabel.Text = ['Error saving metadata: ' ME.message];
            end
        end
        
        % -----------------------------------------------------------------
        % HELPER: ensure KPI row for ID exists (and init NaN for KPI columns)
        % -----------------------------------------------------------------
        function ensureKpiRowForID(app, idValue, idNameValue, dateValue)
            % Make sure kpiTable has a row for the given ID.
            % If not, create one with NaN for all KPI columns.
            
            % If KPI table is empty, create a minimal structure first.
            % We assume that KPI header row already exists in Excel.
            if isempty(app.kpiTable)
                % Minimal columns
                app.kpiTable = table( ...
                    idValue, ...
                    string(idNameValue), ...
                    dateValue, ...
                    'VariableNames', {'ID', 'IDName', 'Date'});
                return;
            end
            
            % Check if ID already exists in KPI table
            if any(app.kpiTable.ID == idValue)
                % Row already exists; nothing to do
                return;
            end
            
            % Create a new row with same variable names as kpiTable
            varNames = app.kpiTable.Properties.VariableNames;
            numVars = numel(varNames);
            newRowCell = cell(1, numVars);
            
            for i = 1:numVars
                varName = varNames{i};
                switch varName
                    case 'ID'
                        newRowCell{i} = idValue;
                    case 'IDName'
                        newRowCell{i} = string(idNameValue);
                    case 'Date'
                        newRowCell{i} = dateValue;
                    otherwise
                        % Initialize KPI columns with NaN (double) by default
                        newRowCell{i} = NaN;
                end
            end
            
            newKpiRow = cell2table(newRowCell, 'VariableNames', varNames);
            app.kpiTable = [app.kpiTable; newKpiRow];
        end
        
        % -----------------------------------------------------------------
        % CALLBACK: Save KPI value
        % -----------------------------------------------------------------
        function onSaveKPI(app, event) %#ok<INUSD>
            try
                % Check if KPI table is available
                if isempty(app.kpiTable)
                    app.KPI_StatusLabel.Text = 'KPI table is empty. Create metadata rows first.';
                    return;
                end
                
                % Read selection from UI
                selectedIDStr = app.KPI_IDDropDown.Value;
                if isempty(selectedIDStr)
                    app.KPI_StatusLabel.Text = 'No ID selected.';
                    return;
                end
                selectedID = str2double(selectedIDStr);
                
                selectedKpiName = app.KPI_NameDropDown.Value;
                if isempty(selectedKpiName)
                    app.KPI_StatusLabel.Text = 'No KPI name selected.';
                    return;
                end
                
                kpiValue = app.KPI_ValueField.Value;
                
                % Find row index in KPI table
                rowIndex = find(app.kpiTable.ID == selectedID, 1, 'first');
                if isempty(rowIndex)
                    app.KPI_StatusLabel.Text = sprintf('ID %d not found in KPI table.', selectedID);
                    return;
                end
                
                % Check column exists
                if ~any(strcmp(selectedKpiName, app.kpiTable.Properties.VariableNames))
                    app.KPI_StatusLabel.Text = sprintf('KPI column "%s" not found.', selectedKpiName);
                    return;
                end
                
                % Assign value in memory
                app.kpiTable{rowIndex, selectedKpiName} = kpiValue;
                
                % ---- WRITE KPI DATA ONLY (no headers) ----
                % Same logic as in onSaveMetadata: only values, in the right columns.
                writetable(app.kpiTable, app.excelFilePath, ...
                    'Sheet', app.kpiSheetName, ...
                    'Range', 'A3', ...              % first data row
                    'WriteVariableNames', false);   % do NOT write headers
                
                app.KPI_StatusLabel.Text = sprintf('KPI "%s" for ID %d updated to %g.', ...
                    selectedKpiName, selectedID, kpiValue);
                
            catch ME
                app.KPI_StatusLabel.Text = ['Error saving KPI: ' ME.message];
            end
        end
        
    end % methods (Access = private)

    % ---------------------------------------------------------------------
    % PUBLIC METHODS (constructor, destructor)
    % ---------------------------------------------------------------------
    methods (Access = public)
        
        function app = MetadataKPIApp(excelFilePath)
            % Constructor
            %
            % Usage:
            %   app = MetadataKPIApp();                 % asks for Excel file
            %   app = MetadataKPIApp('C:\file.xlsx');   % uses given file
            
            % Create UI components
            createComponents(app);
            
            % Store optional Excel path
            if nargin >= 1 && ~isempty(excelFilePath)
                app.excelFilePath = string(excelFilePath);
            end
            
            % Run startup logic
            startup(app);
        end
        
        function delete(app)
            % Destructor, close figure if it still exists
            if isvalid(app.UIFigure)
                delete(app.UIFigure);
            end
        end
        
    end % methods (Access = public)

end