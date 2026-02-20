%% Mass Import of RHO, T, and SALO18 cores with GPS + per-file CoreID + StationNumber/Visit (minimal changes)
clear; close all; clc
tStart = tic;

% rootFolder = "C:\Users\evsalg001\Documents\MATLAB\Contrasts coring\Data";
projectRoot = fileparts(which("import_all_cores.m"));
if isempty(projectRoot); projectRoot = pwd; end
rootFolder = fullfile(projectRoot, "Data");

filePatterns = ["*-RHO.xlsx","*-RHO-TXT.xlsx","*-SALO18.xlsx","*-T.xlsx"];
allFiles = [];

for fp = filePatterns
    files = dir(fullfile(rootFolder,"**",fp));
    allFiles = [allFiles; files];
end

% Initialize empty tables for each core type
T_all = struct('rho', table(), 'T', table(), 'SALO18', table());

% Per-file core counters (new file = new core)
rhoCoreCounter  = 0;
tCoreCounter    = 0;
saloCoreCounter = 0;

for k = 1:numel(allFiles)

    filePath = fullfile(allFiles(k).folder, allFiles(k).name);

    pathParts = split(allFiles(k).folder, filesep);
    station   = pathParts{end-1};
    coreName  = pathParts{end};
    extraMeta = struct("Station",station,"Core",coreName);

    % ---------------- RHO cores ----------------
    if contains(filePath,'-RHO.xlsx') || contains(filePath,'-RHO-TXT.xlsx')

        % Detect options (same as your original)
        try
            opts = detectImportOptions(filePath,'Sheet','Density-densimetry','VariableNamingRule','preserve');
        catch ME
            warning("Skipping RHO file (no Density-densimetry sheet): %s\n  %s", filePath, ME.message);
            continue
        end

        % Your original used columns [1 2 7 8]. Some files (e.g., lead) may have fewer.
        nCols = numel(opts.VariableNames);
        if nCols < 3
            warning("Skipping RHO file (too few columns): %s", filePath);
            continue
        end

        % Robust selection: always depth1+depth2 = [1 2]
        % density: prefer col 7 if exists, else col 3
        % comments: prefer col 8 if exists, else last column
        cDepth1 = 1;
        cDepth2 = 2;
        cDens   = min(7, nCols);  if cDens < 3, cDens = 3; end
        cCom    = min(8, nCols);  if cCom < 3, cCom = nCols; end

        opts.SelectedVariableNames = opts.VariableNames([cDepth1 cDepth2 cDens cCom]);

        try
            T = readtable(filePath, opts);
        catch ME
            warning("Skipping RHO file (readtable failed): %s\n  %s", filePath, ME.message);
            continue
        end

        % Convert numeric columns EXACTLY like your original: skip comment column (4)
        varNames = T.Properties.VariableNames;
        for c = 1:numel(varNames)
            if c==4, continue, end
            v = T.(varNames{c});
            if iscell(v) || isstring(v)
                T.(varNames{c}) = str2double(string(v));
            end
        end

        % Parse comments EXACTLY like your original
        comments = string(T{:,4});
        rho_all  = str2double(regexp(comments, '(?i)(?<=parafine[^=]*=\s*)[-+]?\d*\.?\d+', 'match','once'));
        Tlab_all = str2double(regexp(comments, '(?i)(?<=lab[^=]*=\s*)[-+]?\d*\.?\d+', 'match','once'));

        rho_all  = fillmissing(rho_all,'previous');
        Tlab_all = fillmissing(Tlab_all,'previous');

        f = find(~isnan(rho_all),1,'first');  if ~isempty(f), rho_all(1:f-1)=rho_all(f); end
        f = find(~isnan(Tlab_all),1,'first'); if ~isempty(f), Tlab_all(1:f-1)=Tlab_all(f); end

        T.rho_parafine = rho_all;
        T.Tlab         = Tlab_all;

        % Keep rows with valid depth
        T = T(~isnan(T{:,1}), :);

        % SALO18 safe import (same as your original)
        try
            T_salo = readtable(filePath,'Sheet','SALO18','VariableNamingRule','preserve');
            if size(T_salo,2)>=5
                col5 = T_salo{:,5};
                if iscell(col5) || isstring(col5)
                    col5 = str2double(string(col5));
                end
                col5 = col5(:);
                col5 = padOrTrim(col5, height(T));
            else
                col5 = NaN(height(T),1);
            end
        catch
            col5 = NaN(height(T),1);
        end
        T.salinity = col5;

        % Add metadata and new columns
        newTbl = add_metadata(T,filePath,{'A7','A8','A10','A2'},{'C7','C8','C10','C2'},extraMeta);

        % Add GPS (first waypoint)
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        % Per-file core ID (new)
        if ~isempty(newTbl)
            rhoCoreCounter = rhoCoreCounter + 1;
            newTbl.CoreID_RHO = repmat(rhoCoreCounter, height(newTbl), 1);
            T_all.rho = [T_all.rho; newTbl];
        end

    % ---------------- T cores ----------------
    elseif contains(filePath,'-T.xlsx')

        try
            T = readtable(filePath,'Sheet','TEMP','Range','A2:B1000', ...
                'ReadVariableNames',true,'VariableNamingRule','preserve');
        catch ME
            warning("Skipping T file (cannot read TEMP): %s\n  %s", filePath, ME.message);
            continue
        end

        T = forceNumericTable(T);
        T = T(~isnan(T{:,1}), :);

        newTbl = add_metadata(T,filePath,{'A7','A8','A10'},{'C7','C8','C10'},extraMeta);
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        if ~isempty(newTbl)
            tCoreCounter = tCoreCounter + 1;
            newTbl.CoreID_T = repmat(tCoreCounter, height(newTbl), 1);
            T_all.T = [T_all.T; newTbl];
        end

    % ---------------- SALO18 cores ----------------
    elseif contains(filePath,'-SALO18.xlsx')

        try
            opts_S = detectImportOptions(filePath,'Sheet','SALO18','VariableNamingRule','preserve');
            opts_S.SelectedVariableNames = opts_S.VariableNames([1 2 5]);
            T = readtable(filePath, opts_S);
        catch ME
            warning("Skipping SALO18 file (cannot read SALO18): %s\n  %s", filePath, ME.message);
            continue
        end

        T = forceNumericTable(T);
        T = T(~isnan(T{:,1}), :);
        if ~isempty(T)
            T.Salinity = T{:,3}; T(:,3) = [];
        end

        newTbl = add_metadata(T,filePath,{'A7','A8','A10'},{'C7','C8','C10'},extraMeta);
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        if ~isempty(newTbl)
            saloCoreCounter = saloCoreCounter + 1;
            newTbl.CoreID_SALO18 = repmat(saloCoreCounter, height(newTbl), 1);
            T_all.SALO18 = [T_all.SALO18; newTbl];
        end
    end
end

elapsed = toc(tStart);
fprintf('Mass import completed in %.0f seconds.\n', elapsed);
fprintf("Imported cores: RHO=%d, T=%d, SALO18=%d\n", rhoCoreCounter, tCoreCounter, saloCoreCounter);

save('Coring_data_imported.mat','T_all');
fprintf('Imported data saved to Coring_data_imported.mat\n');
clearvars -except T_all

% ---------------- helper functions ----------------
function T = forceNumericTable(T)
    vars = T.Properties.VariableNames;
    for i = 1:numel(vars)
        v = T.(vars{i});
        if iscell(v) || isstring(v)
            T.(vars{i}) = str2double(string(v));
        end
    end
end

function v = padOrTrim(v,n)
    if numel(v)<n
        v(end+1:n)=v(end);
    elseif numel(v)>n
        v=v(1:n);
    end
end

function tbl = add_metadata(tbl,file,name_cells,value_cells,extraMeta)
    if isempty(tbl), return; end

    % Metadata from sheet
    for j = 1:numel(name_cells)
        nm = readcell(file,'Sheet','metadata-core','Range',name_cells{j});
        varname = matlab.lang.makeValidName(string(nm{1}));

        val = readcell(file,'Sheet','metadata-core','Range',value_cells{j});
        val = val{1};

        if ischar(val) || isstring(val)
            tmp = str2double(string(val));
            if ~isnan(tmp), val = tmp; end
        end
        if isnumeric(val) && contains(lower(varname),'date')
            val = datetime(val,'ConvertFrom','excel');
        end
        if isempty(val) || (~isnumeric(val) && ~isdatetime(val))
            val = NaN;
        end

        tbl.(varname) = repmat(val,height(tbl),1);
    end

    % Extra metadata
    fn = fieldnames(extraMeta);
    for i = 1:numel(fn)
        tbl.(fn{i}) = repmat(string(extraMeta.(fn{i})),height(tbl),1);
    end

    % StationNumber + StationVisit from Station text like "Station1a"
    if ismember("Station", tbl.Properties.VariableNames)
        stationStr = string(tbl.Station);
        stationNum = str2double(regexp(stationStr, '\d+', 'match', 'once'));
        stationVisit = regexp(stationStr, '[a-zA-Z]$', 'match', 'once');
        stationVisit = lower(string(stationVisit));
        stationVisit(ismissing(stationVisit)) = "";
        tbl.StationNumber = stationNum;
        tbl.StationVisit  = stationVisit;
    else
        tbl.StationNumber = NaN(height(tbl),1);
        tbl.StationVisit  = repmat("", height(tbl), 1);
    end

    % SourceFile
    [~,fname,ext] = fileparts(file);
    tbl.SourceFile = repmat(string(strcat(fname, ext)), height(tbl),1);

    % IceAge
    iceAge = repmat("Unknown", height(tbl),1);
    if contains(fname,'FYI','IgnoreCase',true), iceAge(:)="FYI";
    elseif contains(fname,'SMYI','IgnoreCase',true), iceAge(:)="SMYI";
    elseif contains(fname,'SYI','IgnoreCase',true), iceAge(:)="SYI";
    end
    tbl.IceAge = iceAge;

    % Event label
    expr = "PS\d+.*?(?=-SI_corer)";
    tokens = regexp(fname, expr, 'match', 'once');
    if isempty(tokens)
        tbl.EventLabel = repmat("Unknown", height(tbl), 1);
    else
        tbl.EventLabel = repmat(string(tokens), height(tbl), 1);
    end

    % MeltPond
    tbl.MeltPond = repmat(contains(fname,'melt_pond','IgnoreCase',true), height(tbl),1);
end

function [lat,lon,t_gps] = getFirstGPS(coreFile,nRows)
    lat = nan(nRows,1);
    lon = nan(nRows,1);
    t_gps = NaT(nRows,1,'TimeZone','UTC');

    gpsFolder = fullfile(fileparts(coreFile),'GPSdata');
    if isfolder(gpsFolder)
        gpsFiles = dir(fullfile(gpsFolder,'Waypoints_*.gpx'));
        if ~isempty(gpsFiles)
            gpsFile = fullfile(gpsFiles(1).folder, gpsFiles(1).name);
            try
                xDoc = xmlread(gpsFile);
                wptNodes = xDoc.getElementsByTagName('wpt');
                if wptNodes.getLength>0
                    firstNode = wptNodes.item(0);
                    lat = repmat(str2double(firstNode.getAttribute('lat')), nRows,1);
                    lon = repmat(str2double(firstNode.getAttribute('lon')), nRows,1);
                    timeNode = firstNode.getElementsByTagName('time').item(0);
                    t_gps_val = datetime(char(timeNode.getTextContent),...
                        'InputFormat','yyyy-MM-dd''T''HH:mm:ss''Z''','TimeZone','UTC');
                    t_gps = repmat(t_gps_val, nRows,1);
                end
            catch
                warning('Failed to read GPX: %s', coreFile);
            end
        end
    end
end

%% Load imported -> process -> format/reorder/drop -> rename -> save (filename-key matching incl. lead)
% Output:
%   T_all_proc.rho          (original + processed columns)
%   T_all_proc.rho_out      (ONLY selected columns, preferred order, renamed)
%   T_all_proc.T_out        (ONLY selected columns, preferred order, renamed)
%   T_all_proc.SALO18_out   (ONLY selected columns, preferred order, renamed)

clear; close all; clc

scriptPath = which("import_all_cores.m");
if isempty(scriptPath)
    error("Cannot find import_all_cores.m on MATLAB path. Set Current Folder to the script folder or add it to path.");
end
scriptDir = fileparts(scriptPath);
exportFolder = fullfile(scriptDir, "Export");
if ~isfolder(exportFolder), mkdir(exportFolder); end

MAP = struct();

% Common/meta
MAP.EventLabel     = "EventLabel";
MAP.GPS_Time       = "GPS_Time";
MAP.GPS_Lat        = "GPS_Lat";
MAP.GPS_Lon        = "GPS_Lon";
MAP.Station        = "Station";         % internal only
MAP.StationNumber  = "StationNumber";   % new
MAP.StationVisit   = "StationVisit";    % new
MAP.IceAge         = "IceAge";
MAP.MeltPond       = "MeltPond";
MAP.Comments       = "";

% Core counters from import (new)
MAP.CoreID_RHO     = "CoreID_RHO";
MAP.CoreID_T       = "CoreID_T";
MAP.CoreID_SALO18  = "CoreID_SALO18";

% RHO-specific
MAP.Depth1       = "";
MAP.Depth2       = "";
MAP.Salinity_raw = "salinity";
MAP.Salinity_used= "Salinity_used";
MAP.Tlab         = "Tlab";
MAP.T_ice        = "Temperature_interp";
MAP.Rho_si       = "rho_si";

MAP.Vb_export    = "vb_rho_export";
MAP.Vg_pr        = "vg_pr";
MAP.Vg           = "vg";

% Geometry (optional, guessed if empty)
MAP.IceThickness = "";
MAP.IceDraft     = "";
MAP.CoreLength   = "";

% T-core specific (guessed if empty)
MAP.T_Depth      = "";
MAP.T_Temp       = "";

% SALO18 specific (guessed if empty)
MAP.S_Depth1     = "";
MAP.S_Depth2     = "";
MAP.S_Salinity   = "";

OUTFILE_PROCESSED = "Coring_data_processed.mat";

% 1) Load imported data
load('Coring_data_imported.mat','T_all');
fprintf('Imported data loaded from MAT-file\n');

% Normalize types
if isfield(T_all,'T') && ~isempty(T_all.T)
    T_all.T.SourceFile   = string(T_all.T.SourceFile);
    if ismember(MAP.Station, T_all.T.Properties.VariableNames), T_all.T.(MAP.Station) = string(T_all.T.(MAP.Station)); end
    if ismember("Core", T_all.T.Properties.VariableNames),      T_all.T.Core          = string(T_all.T.Core); end
end
if isfield(T_all,'rho') && ~isempty(T_all.rho)
    T_all.rho.SourceFile = string(T_all.rho.SourceFile);
    if ismember(MAP.Station, T_all.rho.Properties.VariableNames), T_all.rho.(MAP.Station) = string(T_all.rho.(MAP.Station)); end
    if ismember("Core", T_all.rho.Properties.VariableNames),      T_all.rho.Core          = string(T_all.rho.Core); end
end
if isfield(T_all,'SALO18') && ~isempty(T_all.SALO18)
    T_all.SALO18.SourceFile = string(T_all.SALO18.SourceFile);
    if ismember(MAP.Station, T_all.SALO18.Properties.VariableNames), T_all.SALO18.(MAP.Station) = string(T_all.SALO18.(MAP.Station)); end
    if ismember("Core", T_all.SALO18.Properties.VariableNames),      T_all.SALO18.Core          = string(T_all.SALO18.Core); end
end

% Ensure new columns exist (backward compatible)
T_all = ensureNewCols(T_all, MAP);

% 2) Make processed copy
T_all_proc = T_all;

% Ensure required processed columns exist in RHO
needCols = ["rho_si","rho_lab_kgm3","Salinity_used","Temperature_interp", ...
            "vb_rho_export","vg_pr","vg"];
for c = needCols
    if ~ismember(c, T_all_proc.rho.Properties.VariableNames)
        T_all_proc.rho.(c) = NaN(height(T_all_proc.rho),1);
    end
end

% Ensure lab density for ALL RHO rows (even if no matching T core)
T_all_proc.rho.rho_lab_kgm3 = T_all_proc.rho{:,3} * 1000;  % assumes col3 = g/cm^3

% 3) Match T and RHO cores within each folder by:
%    (a) same Station+Core folder
%    (b) same descriptor group (e.g., FYI vs FYI-lead)
%    (c) pair by sorted order of the numeric token after "...-SI_corer_?cm-"
Tmatch = table(strings(0,1), strings(0,1), strings(0,1), strings(0,1), ...
               'VariableNames', {'Station','Core','T_File','RHO_File'});

% Unique folders (Station+Core) present in T table
stCol = MAP.Station; % internal
coCol = "Core";

if ~ismember(stCol, T_all.T.Properties.VariableNames) || ~ismember(coCol, T_all.T.Properties.VariableNames)
    error("T table must contain Station and Core columns for folder grouping.");
end
if ~ismember(stCol, T_all.rho.Properties.VariableNames) || ~ismember(coCol, T_all.rho.Properties.VariableNames)
    error("RHO table must contain Station and Core columns for folder grouping.");
end

foldersT = unique(T_all.T(:, [stCol, coCol]), 'rows');

for f = 1:height(foldersT)
    st = string(foldersT.(stCol)(f));
    co = string(foldersT.(coCol)(f));

    Tfiles   = unique(T_all.T.SourceFile(  T_all.T.(stCol)==st & T_all.T.(coCol)==co ), 'stable');
    RHOfiles = unique(T_all.rho.SourceFile(T_all.rho.(stCol)==st & T_all.rho.(coCol)==co), 'stable');

    if isempty(Tfiles)
        continue
    end
    if isempty(RHOfiles)
        warning("No RHO files found in folder Station=%s / Core=%s", st, co);
        continue
    end

    % Build descriptor groups for each file (FYI vs FYI-lead etc.) and extract numbers
    [Tdesc, Tnum]     = arrayfun(@parseCoreDescriptorAndNumber, Tfiles,   'UniformOutput', false);
    [Rdesc, Rnum]     = arrayfun(@parseCoreDescriptorAndNumber, RHOfiles, 'UniformOutput', false);

    Tdesc = string(Tdesc);  Tnum = cell2mat(Tnum);
    Rdesc = string(Rdesc);  Rnum = cell2mat(Rnum);

    % For each descriptor group present in T, match to the same descriptor group in RHO
    groups = unique(Tdesc, 'stable');
    for g = 1:numel(groups)
        grp = groups(g);

        iT = find(Tdesc == grp);
        iR = find(Rdesc == grp);

        if isempty(iR)
            warning("No RHO group match in folder Station=%s / Core=%s for group '%s'", st, co, grp);
            continue
        end

        % Sort by the extracted number token
        [~, oT] = sort(Tnum(iT));
        [~, oR] = sort(Rnum(iR));

        Tsorted = Tfiles(iT(oT));
        Rsorted = RHOfiles(iR(oR));

        if numel(Tsorted) ~= numel(Rsorted)
            warning("Count mismatch in folder Station=%s / Core=%s group '%s': T=%d, RHO=%d. Pairing by min count.", ...
                st, co, grp, numel(Tsorted), numel(Rsorted));
        end

        nPair = min(numel(Tsorted), numel(Rsorted));
        for p = 1:nPair
            Tmatch = [Tmatch; table(st, co, Tsorted(p), Rsorted(p), ...
                'VariableNames', {'Station','Core','T_File','RHO_File'})];
        end
    end
end

fprintf('Core matching completed: %d pairs\n', height(Tmatch));

% 4) Compute rho_si + vb/vg and write back into T_all_proc.rho
nPairs = height(Tmatch);
rho_si_bulk = nan(nPairs,1);
T_bulk      = nan(nPairs,1);

for i = 1:nPairs

    T_T   = T_all.T(  T_all.T.SourceFile   == Tmatch.T_File(i), :);
    T_rho = T_all.rho(T_all.rho.SourceFile == Tmatch.RHO_File(i), :);

    if isempty(T_T) || isempty(T_rho)
        warning('Missing data for this pair — skipping');
        continue
    end

    depth_rho = mean(T_rho{:,1:2}, 2);

    depth_T = T_T{:,1};
    temp    = min(-0.1, T_T{:,2});
    ok = ~isnan(depth_T) & ~isnan(temp);
    depth_T = depth_T(ok);
    temp    = temp(ok);

    if numel(depth_T) < 2
        warning('Too few temperature points — skipping');
        continue
    end

    depth_T_rescaled = depth_T * (max(depth_rho) / max(depth_T));
    T_interp = interp1(depth_T_rescaled, temp, depth_rho, 'linear', 'extrap');

    % Salinity used
    if ismember(MAP.Salinity_raw, T_rho.Properties.VariableNames)
        Srho = T_rho.(MAP.Salinity_raw);
    else
        SrhoVar = guessVar(T_rho, ["salinity","S","SALO"], "");
        if SrhoVar ~= ""
            Srho = T_rho.(SrhoVar);
        else
            Srho = NaN(height(T_rho),1);
        end
    end

    % SALO18 fallback if all NaN
    if all(isnan(Srho))
        SALOcore = [];
        if isfield(T_all,'SALO18') && ~isempty(T_all.SALO18) && ismember(MAP.Station, T_all.SALO18.Properties.VariableNames)
            SALOcore = T_all.SALO18(T_all.SALO18.(MAP.Station) == Tmatch.Station(i), :);
        end
        if ~isempty(SALOcore)
            depth_SALO = mean(SALOcore{:,1:2}, 2);
            sal_SALO   = SALOcore{:,3};
            Srho = interp1(depth_SALO, sal_SALO, depth_rho, 'linear', 'extrap');
        else
            warning('No SALO18 core found — using NaN for salinity');
        end
    end

    % Lab temperature
    if ismember(MAP.Tlab, T_rho.Properties.VariableNames)
        T_lab = T_rho.(MAP.Tlab);
    else
        TlabVar = guessVar(T_rho, ["Tlab","lab"], "");
        if TlabVar ~= ""
            T_lab = T_rho.(TlabVar);
        else
            warning('No lab temperature column found — skipping');
            continue
        end
    end

    rho_meas = T_rho{:,3};
    rho_lab_kgm3 = rho_meas * 1000;
    rho = rho_lab_kgm3;

    % ===================== CALCULATIONS =====================
    % From Cox and Weeks (1983), Lepparanta and Manninen (1988)
    F1_pr_rho = -4.732 - 22.45*T_lab - 0.6397*T_lab.^2 - 0.01074*T_lab.^3;
    F2_pr_rho = 8.903e-2 - 1.763e-2*T_lab - 5.33e-4*T_lab.^2 - 8.801e-6*T_lab.^3;
    vb_pr_rho = rho .* Srho ./ F1_pr_rho; % Brine volume at laboratory temperature
    rhoi_pr   = 917 - 0.1403*T_lab; % Pure ice density at laboratory temperature
    vg_pr     = max(0, 1 - rho .* (F1_pr_rho - rhoi_pr .* Srho/1000 .* F2_pr_rho) ./ (rhoi_pr .* F1_pr_rho)); % Gas volume at laboratory temperature
    F3_pr     = rhoi_pr .* Srho/1000 ./ (F1_pr_rho - rhoi_pr .* Srho/1000 .* F2_pr_rho);

    T_insitu = T_interp; % In situ sea ice temperature
    rhoi_rho = 917 - 0.1403*T_insitu; % Pure ice density at in situ temperature
    F1_rho = -4.732 - 22.45*T_insitu - 0.6397*T_insitu.^2 - 0.01074*T_insitu.^3;
    F2_rho = 8.903e-2 - 1.763e-2*T_insitu - 5.33e-4*T_insitu.^2 - 8.801e-6*T_insitu.^3;

    idx_LM = T_insitu > -2;
    F1_rho(idx_LM) = -4.1221e-2 - 18.407*T_insitu(idx_LM) ...
                   + 0.58402*T_insitu(idx_LM).^2 ...
                   + 0.21454*T_insitu(idx_LM).^3;
    F2_rho(idx_LM) = 9.0312e-2 - 0.016111*T_insitu(idx_LM) ...
                   + 1.2291e-4*T_insitu(idx_LM).^2 ...
                   + 1.3603e-4*T_insitu(idx_LM).^3;

    F3_rho = rhoi_rho .* Srho/1000 ./ (F1_rho - rhoi_rho .* Srho/1000 .* F2_rho);

    vb_rho_raw = vb_pr_rho .* F1_pr_rho ./ F1_rho / 1000; % Brine volume at in situ temperature

    vb_rho_export = vb_rho_raw;
    vb_rho_export(vb_rho_export > 0.4 | vb_rho_export < 0) = NaN; 

    vb_rho = vb_rho_raw;
    vb_rho(vb_rho > 0.4 | vb_rho < 0) = 0.4; % Limits brine volume to 40% for innacurate sea ice temperatures

    vg = max(0, (1 - (1 - vg_pr) .* (rhoi_rho./rhoi_pr) .* (F3_pr.*F1_pr_rho./(F3_rho.*F1_rho))));  % Gas volume at in situ temperature

    rho_si = (1 - vg) .* rhoi_rho .* F1_rho ./ (F1_rho - rhoi_rho .* Srho/1000 .* F2_rho); % Sea ice density at in situ temperature
    rho_si(isnan(vb_rho)) = NaN;
    % =========================================================

    rho_si_bulk(i) = mean(rho_si,'omitnan');
    T_bulk(i)      = mean(T_insitu,'omitnan');

    idxAll = (T_all_proc.rho.SourceFile == Tmatch.RHO_File(i));
    if nnz(idxAll) ~= numel(rho_si)
        warning("Row count mismatch for %s (T_all has %d rows, processed has %d). Skipping write-back.", ...
            Tmatch.RHO_File(i), nnz(idxAll), numel(rho_si));
        continue
    end

    T_all_proc.rho.(MAP.Rho_si)(idxAll)            = rho_si;
    T_all_proc.rho.rho_lab_kgm3(idxAll)            = rho_lab_kgm3;
    T_all_proc.rho.(MAP.Salinity_used)(idxAll)     = Srho;
    T_all_proc.rho.(MAP.T_ice)(idxAll)             = T_interp;

    T_all_proc.rho.(MAP.Vb_export)(idxAll)         = vb_rho_export;
    T_all_proc.rho.(MAP.Vg_pr)(idxAll)             = vg_pr;
    T_all_proc.rho.(MAP.Vg)(idxAll)                = vg;
end

% 5) FORMAT / REORDER / DROP columns for final exports
rho = T_all_proc.rho;
rho = addTimeBest(rho, MAP.GPS_Time);
rho = addMeltPond01(rho, MAP.MeltPond);

d1 = pickOrDefault(rho, MAP.Depth1, rho.Properties.VariableNames{1});
d2 = pickOrDefault(rho, MAP.Depth2, rho.Properties.VariableNames{2});

salFinal = "";
if ismember(MAP.Salinity_used, rho.Properties.VariableNames)
    rho.Salinity_final = rho.(MAP.Salinity_used);
    salFinal = "Salinity_final";
elseif ismember(MAP.Salinity_raw, rho.Properties.VariableNames)
    rho.Salinity_final = rho.(MAP.Salinity_raw);
    salFinal = "Salinity_final";
else
    gsal = guessVar(rho, ["salinity","S"], "");
    rho.Salinity_final = NaN(height(rho),1);
    if gsal ~= "", rho.Salinity_final = rho.(gsal); end
    salFinal = "Salinity_final";
end

iceTh = pickOrDefault(rho, MAP.IceThickness, guessVar(rho, ["thick","thickness"], ""));
iceDr = pickOrDefault(rho, MAP.IceDraft,     guessVar(rho, ["draft"], ""));
coLen = pickOrDefault(rho, MAP.CoreLength,   guessVar(rho, ["length","corelength"], ""));

com = MAP.Comments;
if com == "", com = guessVar(rho, ["comment","remarks","notes"], ""); end

rhoLabVar = "rho_lab_kgm3";

preferredRHO = [
    string(MAP.EventLabel)
    "Time_best"
    string(MAP.GPS_Lat)
    string(MAP.GPS_Lon)
    iceTh
    iceDr
    coLen
    string(MAP.CoreID_RHO)
    string(MAP.StationNumber)
    string(MAP.StationVisit)
    string(MAP.IceAge)
    "MeltPond01"
    d1
    d2
    salFinal
    string(MAP.Tlab)
    string(MAP.T_ice)
    rhoLabVar
    string(MAP.Rho_si)
    string(MAP.Vb_export)
    string(MAP.Vg_pr)
    string(MAP.Vg)
    com
];
T_all_proc.rho_out = keepAndOrder(rho, preferredRHO);

TT = T_all_proc.T;
TT = addTimeBest(TT, MAP.GPS_Time);
TT = addMeltPond01(TT, MAP.MeltPond);

tDepth = pickOrDefault(TT, MAP.T_Depth, guessVar(TT, ["depth"], TT.Properties.VariableNames{1}));
tTemp  = pickOrDefault(TT, MAP.T_Temp,  guessVar(TT, ["temp","temperature"], TT.Properties.VariableNames{2}));

iceTh_T = pickOrDefault(TT, MAP.IceThickness, guessVar(TT, ["thick","thickness"], ""));
iceDr_T = pickOrDefault(TT, MAP.IceDraft,     guessVar(TT, ["draft"], ""));
coLen_T = pickOrDefault(TT, MAP.CoreLength,   guessVar(TT, ["length","corelength"], ""));

comT = MAP.Comments;
if comT == "", comT = guessVar(TT, ["comment","remarks","notes"], ""); end

preferredT = [
    string(MAP.EventLabel)
    "Time_best"
    string(MAP.GPS_Lat)
    string(MAP.GPS_Lon)
    iceTh_T
    iceDr_T
    coLen_T
    string(MAP.CoreID_T)
    string(MAP.StationNumber)
    string(MAP.StationVisit)
    string(MAP.IceAge)
    "MeltPond01"
    tDepth
    tTemp
    comT
];
T_all_proc.T_out = keepAndOrder(TT, preferredT);

if isfield(T_all_proc,'SALO18') && ~isempty(T_all_proc.SALO18)
    S = T_all_proc.SALO18;
    S = addTimeBest(S, MAP.GPS_Time);
    S = addMeltPond01(S, MAP.MeltPond);

    sD1  = pickOrDefault(S, MAP.S_Depth1, S.Properties.VariableNames{1});
    sD2  = pickOrDefault(S, MAP.S_Depth2, S.Properties.VariableNames{2});
    sSal = pickOrDefault(S, MAP.S_Salinity, guessVar(S, ["salinity","S"], ""));

    iceTh_S = pickOrDefault(S, MAP.IceThickness, guessVar(S, ["thick","thickness"], ""));
    iceDr_S = pickOrDefault(S, MAP.IceDraft,     guessVar(S, ["draft"], ""));
    coLen_S = pickOrDefault(S, MAP.CoreLength,   guessVar(S, ["length","corelength"], ""));

    comS = MAP.Comments;
    if comS == "", comS = guessVar(S, ["comment","remarks","notes"], ""); end

    preferredS = [
        string(MAP.EventLabel)
        "Time_best"
        string(MAP.GPS_Lat)
        string(MAP.GPS_Lon)
        iceTh_S
        iceDr_S
        coLen_S
        string(MAP.CoreID_SALO18)
        string(MAP.StationNumber)
        string(MAP.StationVisit)
        string(MAP.IceAge)
        "MeltPond01"
        sD1
        sD2
        sSal
        comS
    ];
    T_all_proc.SALO18_out = keepAndOrder(S, preferredS);
else
    warning('No SALO18 table found in T_all_proc');
end

% 6) Rename headers (PANGAEA convention) - export tables only
renameMap = {
    "EventLabel",          "Event label"
    "Time_best",           "DATE/TIME"
    "GPS_Lat",             "LATITUDE"
    "GPS_Lon",             "LONGITUDE"

    "CoreID_RHO",          "Core number (RHO)"
    "CoreID_T",            "Core number (T)"
    "CoreID_SALO18",       "Core number (SALO18)"
    "StationNumber",       "Ice station number"
    "StationVisit",        "Ice station visit"

    "rho_lab_kgm3",         "Density, ice, technical"
    "rho_si",              "Density, ice"
    "vb_rho_export",       "Volume, brine"
    "vg_pr",               "Volume, gas, technical"
    "vg",                  "Volume, gas"
    "Salinity_final",      "Sea ice salinity"
    "Tlab",                "Temperature, technical"
    "Temperature_interp",  "Temperature, ice/snow"
    "temperature",         "Temperature, ice/snow"
    "MeltPond01",          "Melt pond"
    "IceAge",              "Ice age"
    "CoreLength",          "Core length"
    "depth center",        "Depth, ice/snow"
    "depth 1",             "Depth, ice/snow, top/minimum"
    "depth 2",             "Depth, ice/snow, bottom/maximum"
    "Salinity",            "Sea ice salinity"
};

T_all_proc.rho_out = applyRenameMap(T_all_proc.rho_out, renameMap);
T_all_proc.T_out   = applyRenameMap(T_all_proc.T_out, renameMap);
if isfield(T_all_proc,'SALO18_out')
    T_all_proc.SALO18_out = applyRenameMap(T_all_proc.SALO18_out, renameMap);
end

% Round calculated export variables to 2 decimals
varsToRound_RHO = [
    "Density, ice, technical"
    "Density, ice"
    "Volume, brine"
    "Volume, gas, technical"
    "Volume, gas"
    "Temperature, ice/snow"
];
for k = 1:numel(varsToRound_RHO)
    vn = char(varsToRound_RHO(k));
    if ismember(vn, T_all_proc.rho_out.Properties.VariableNames)
        x = T_all_proc.rho_out.(vn);
        if isnumeric(x)
            T_all_proc.rho_out.(vn) = round(x, 2);
        end
    end
end

% 7) Save
save(OUTFILE_PROCESSED, 'T_all_proc','Tmatch','rho_si_bulk','T_bulk');
fprintf('Saved processed + formatted data to %s\n', OUTFILE_PROCESSED);
clearvars -except T_all_proc Tmatch rho_si_bulk T_bulk exportFolder

% % Export 3 final tables to one Excel workbook
% outXlsx = "Coring_data_export.xlsx";
% writetable(T_all_proc.rho_out,    outXlsx, "Sheet", "RHO",    "WriteMode", "overwritesheet");
% writetable(T_all_proc.T_out,      outXlsx, "Sheet", "T",      "WriteMode", "overwritesheet");
% writetable(T_all_proc.SALO18_out, outXlsx, "Sheet", "SALO18", "WriteMode", "overwritesheet");
% fprintf("Exported to %s (sheets: RHO, T, SALO18)\n", outXlsx);

% Export 3 final tables to one Excel workbook
outXlsx = fullfile(exportFolder, "Coring_data_export.xlsx");
writetable(T_all_proc.rho_out,    outXlsx, "Sheet", "RHO",    "WriteMode", "overwritesheet");
writetable(T_all_proc.T_out,      outXlsx, "Sheet", "T",      "WriteMode", "overwritesheet");
writetable(T_all_proc.SALO18_out, outXlsx, "Sheet", "SALO18", "WriteMode", "overwritesheet");
fprintf("Exported to %s (sheets: RHO, T, SALO18)\n", outXlsx);

%% ===== NetCDF export =====
%  Export Coring tables to NetCDF (with units + comments + categorical attrs)
%  - Writes one NetCDF4 file with groups: /RHO, /T, /SALO18
%  - Variable names are sanitized versions of your table headers
%  - Adds units/standard_name/long_name/comment + category/flag metadata

% Settings
clear; close all; clc

% Locate project folder + Export folder
scriptPath = which("import_all_cores.m");  % <-- change if needed
if isempty(scriptPath)
    error("Cannot find import_all_cores.m on MATLAB path. Set Current Folder to the script folder or add it to path.");
end
scriptDir = fileparts(scriptPath);
exportFolder = fullfile(scriptDir, "Export");
if ~isfolder(exportFolder), mkdir(exportFolder); end

INFILE = "Coring_data_processed.mat";   % your OUTFILE_PROCESSED
% ncFile = "Contrasts_physical_properties_coring.nc";
ncFile = fullfile(exportFolder, "Contrasts_physical_properties_coring.nc");

% Load 
load(INFILE, "T_all_proc");   % <-- one-line load of only T_all_proc

% Create NetCDF file (netcdf4)
if isfile(ncFile), delete(ncFile); end
nccreate(ncFile, "dummy", "Dimensions", {"dummy", 1}, "Format", "netcdf4");
ncwrite(ncFile, "dummy", 1);

% Global attributes
ncwriteatt(ncFile,"/","title","Sea ice physical properties from the Contrasts expedition");
ncwriteatt(ncFile,"/","Conventions","CF-1.7");
ncwriteatt(ncFile,"/","contributor_name","Dmitry Divine, Evgenii Salganik, David Clemens-Sewall, Emiliano Cimoli, Lena Eggers, Keigo Takahashi, Marcel Nicoalus");
ncwriteatt(ncFile,"/","contributor_email","evgenii.salganik@proton.me");
ncwriteatt(ncFile,"/","institution","Alfred Wegener Institute for Polar and Marine Research");
ncwriteatt(ncFile,"/","creator_name","Evgenii Salganik");
ncwriteatt(ncFile,"/","creator_email","evgenii.salganik@proton.me");
ncwriteatt(ncFile,"/","project","Arctic PASSION");
ncwriteatt(ncFile,"/","summary","First- and second-year sea-ice salinity, temperature, and density from the coring sites during the Contrasts expedition in July-September 2025");
ncwriteatt(ncFile,"/","license","CC-0");
ncwriteatt(ncFile,"/","time_coverage_start","2025-07-10 09:17:00");
ncwriteatt(ncFile,"/","time_coverage_end","2025-08-26 08:51:00");
ncwriteatt(ncFile,"/","keywords","arctic, polar, sea ice, salinity, temperature, density, coring");
ncwriteatt(ncFile,"/","geospatial_lat_min","82.4404");
ncwriteatt(ncFile,"/","geospatial_lat_max","84.9907");
ncwriteatt(ncFile,"/","geospatial_lon_min","-17.9440");
ncwriteatt(ncFile,"/","geospatial_lon_max","33.9612");
ncwriteatt(ncFile,"/","calendar","standard");
ncwriteatt(ncFile,"/","date_created",char(datetime("now","TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss'Z'")));
ncwriteatt(ncFile,"/","featureType","timeseries");
ncwriteatt(ncFile,"/","product_version","1");

% Build metadata maps (units + comments)
metaRHO  = buildMetaMap_RHO();
metaT    = buildMetaMap_T();
metaSALO = buildMetaMap_SALO18();

% Write groups
writeTableGroupNetCDF(ncFile, "/RHO",    T_all_proc.rho_out,    metaRHO);
writeTableGroupNetCDF(ncFile, "/T",      T_all_proc.T_out,      metaT);
writeTableGroupNetCDF(ncFile, "/SALO18", T_all_proc.SALO18_out, metaSALO);

fprintf("Exported NetCDF: %s (groups: /RHO, /T, /SALO18)\n", ncFile);

%% ===== Import NetCDF =====
clear; close all; clc

% Locate project + Export folder
scriptPath = which("import_all_cores.m");   % adjust if running from another script
scriptDir  = fileparts(scriptPath);

exportFolder = fullfile(scriptDir, "Export");
ncFile = fullfile(exportFolder, "Contrasts_physical_properties_coring.nc");

% Inspect structure
ncdisp(ncFile)

%% NetCDF -> MATLAB tables (RHO, T, SALO18)
clear; close all; clc

% ---- Locate NetCDF in Export folder (relative to your processing script) ----
scriptPath = which("import_all_cores.m");   % <-- change if your main script has a different name
if isempty(scriptPath)
    error("Cannot find import_all_cores.m on MATLAB path. Set Current Folder to the script folder or add it to path.");
end
scriptDir = fileparts(scriptPath);

exportFolder = fullfile(scriptDir, "Export");
ncFile = fullfile(exportFolder, "Contrasts_physical_properties_coring.nc");

if ~isfile(ncFile)
    error("NetCDF file not found: %s", ncFile);
end

% ---- Build tables from each group ----
T_RHO    = ncGroupToTable(ncFile, "/RHO");
T_T      = ncGroupToTable(ncFile, "/T");
T_SALO18 = ncGroupToTable(ncFile, "/SALO18");

% ---- Quick check ----
fprintf("Loaded tables:\n");
fprintf("  RHO:    %d rows x %d vars\n", height(T_RHO),    width(T_RHO));
fprintf("  T:      %d rows x %d vars\n", height(T_T),      width(T_T));
fprintf("  SALO18: %d rows x %d vars\n", height(T_SALO18), width(T_SALO18));

% ========================= FUNCTIONS =========================
function T = ncGroupToTable(ncFile, grp)
ginfo = ncinfo(ncFile, grp);

% Find obs length from group dimensions
nObs = getDimLength(ginfo, "obs");
if isempty(nObs)
    error("Group %s has no 'obs' dimension. Cannot build table.", grp);
end

T = table();
Tobs = (1:nObs).'; %#ok<NASGU> % just to make intent clear

for k = 1:numel(ginfo.Variables)
    vname = string(ginfo.Variables(k).Name);

    % Skip helper variables
    if vname == "dummy" || startsWith(vname, "strlen_")
        continue
    end

    vpath = grp + "/" + vname;

    % Read variable
    x = ncread(ncFile, vpath);

    % ---- CHAR/string variables ----
    if ischar(x)
        s = char(x);

        sz = size(s);
        if ~any(sz == nObs)
            warning("Skipping %s%s (char) because no dimension matches obs=%d. Size=%s", ...
                grp, "/" + vname, nObs, mat2str(sz));
            continue
        end

        if sz(1) == nObs
            svec = string(strtrim(cellstr(s)));
        else
            svec = string(strtrim(cellstr(s')));
        end

        if numel(svec) ~= nObs
            warning("Skipping %s%s (char->string) because length mismatch: %d vs obs=%d", ...
                grp, "/" + vname, numel(svec), nObs);
            continue
        end

        T.(vname) = svec(:);
        continue
    end

    % ---- NUMERIC/LOGICAL variables ----
    if isnumeric(x) || islogical(x)
        x = x(:);

        if numel(x) ~= nObs
            warning("Skipping %s%s because numel=%d != obs=%d (size=%s)", ...
                grp, "/" + vname, numel(x), nObs, mat2str(size(ncread(ncFile, vpath))));
            continue
        end

        units = "";
        try
            units = ncreadatt(ncFile, vpath, "units");
        catch
        end

        if (ischar(units) || isstring(units)) && contains(lower(string(units)), "days since 1979-01-01")
            ref = datetime(1979,1,1,0,0,0,"TimeZone","UTC");
            T.(vname) = ref + days(double(x));
        else
            T.(vname) = double(x);
        end

        continue
    end

    warning("Skipping %s%s due to unsupported type: %s", grp, "/" + vname, class(x));
end
end

function n = getDimLength(ginfo, dimName)
n = [];
if ~isfield(ginfo, "Dimensions") || isempty(ginfo.Dimensions)
    return
end
for i = 1:numel(ginfo.Dimensions)
    if string(ginfo.Dimensions(i).Name) == string(dimName)
        n = ginfo.Dimensions(i).Length;
        return
    end
end
end

%% Plot: Per-core avg density vs ice thickness (LAB + IN SITU)
% clear; close all; clc

load("Coring_data_processed.mat","T_all_proc");
rho = T_all_proc.rho;

% ---------- Core grouping ----------
if ismember("CoreID_RHO", rho.Properties.VariableNames)
    coreID = double(rho.CoreID_RHO);
else
    warning("CoreID_RHO not found. Falling back to SourceFile grouping.");
    [~,~,coreID] = unique(string(rho.SourceFile), "stable");
    coreID = double(coreID);
end

% ---------- Densities ----------
% Lab density (kg/m^3): prefer rho_lab_kgm3; else column 3 * 1000
if ismember("rho_lab_kgm3", rho.Properties.VariableNames)
    rhoLab = rho.rho_lab_kgm3;
else
    rhoLab = rho{:,3} * 1000;
end

% In situ density (kg/m^3): rho_si (computed); may be NaN for unmatched cores
if ismember("rho_si", rho.Properties.VariableNames)
    rhoIns = rho.rho_si;
else
    error("rho_si not found in T_all_proc.rho");
end

rhoLab = toNum(rhoLab);
rhoIns = toNum(rhoIns);

% ---------- Ice thickness ----------
% Try common metadata names (your import usually creates something like IceThickness or similar)
thkVar = guessVarName(string(rho.Properties.VariableNames), ...
    ["IceThickness","Ice thickness","thickness","Thick","Ice_thickness","Ice_Thickness"]);

if thkVar == ""
    error("Could not find an ice thickness column in T_all_proc.rho. Add/rename it in import or set thkVar manually.");
end

thk = toNum(rho.(thkVar));

% ---------- Aggregate per core ----------
coreList = unique(coreID(~isnan(coreID) & coreID>0), "stable");
n = numel(coreList);

CoreID = nan(n,1);
IceThickness = nan(n,1);
AvgRhoLab = nan(n,1);
AvgRhoIns = nan(n,1);
N_lab = zeros(n,1);
N_ins = zeros(n,1);

for i = 1:n
    cid = coreList(i);
    idx = coreID == cid;

    CoreID(i) = cid;
    IceThickness(i) = firstNonNaN(thk(idx));

    xLab = rhoLab(idx);
    xIns = rhoIns(idx);

    AvgRhoLab(i) = mean(xLab, "omitnan");
    AvgRhoIns(i) = mean(xIns, "omitnan");

    N_lab(i) = sum(~isnan(xLab));
    N_ins(i) = sum(~isnan(xIns));
end

% Keep only cores with thickness and at least one density point
ok = ~isnan(IceThickness) & (N_lab>0 | N_ins>0);
CoreID = CoreID(ok);
IceThickness = IceThickness(ok);
AvgRhoLab = AvgRhoLab(ok);
AvgRhoIns = AvgRhoIns(ok);
N_lab = N_lab(ok);
N_ins = N_ins(ok);

% ---------- Plot ----------
figure("Color","w"); hold on; grid on; box on

idxLab = ~isnan(AvgRhoLab);
idxIns = ~isnan(AvgRhoIns);

scatter(IceThickness(idxLab), AvgRhoLab(idxLab), 70);
scatter(IceThickness(idxIns), AvgRhoIns(idxIns), 70, "filled");

xlabel("Ice thickness");
ylabel("Average density (kg m^{-3})");
title("Per-core average density vs ice thickness");

legend(["Lab density (avg per core)", "In situ density (avg per core)"], ...
    "Location","best");

set(gcf,"Units","inches","Position",[4 4 6.5 4.5])
exportgraphics(gcf, "Density_vs_Thickness.png", "Resolution", 300);
fprintf("Saved: Density_vs_Thickness.png\n");

%% Plot: Salinity, density, temperature vs time
clear; close all; clc

load("Coring_data_processed.mat","T_all_proc");

rho = T_all_proc.rho;

if ~ismember("Date", rho.Properties.VariableNames)
    error("T_all_proc.rho does not contain a 'Date' column.");
end
tRowR = rho.Date;
if isa(tRowR,'datetime') && ~isempty(tRowR.TimeZone), tRowR.TimeZone = ''; end

if ~ismember("StationNumber", rho.Properties.VariableNames)
    error("T_all_proc.rho does not contain 'StationNumber'.");
end
stRowR = double(rho.StationNumber);

if ismember("CoreID_RHO", rho.Properties.VariableNames)
    coreID_R = rho.CoreID_RHO;
else
    [~,~,coreID_R] = unique(string(rho.SourceFile), "stable");
end
coreID_R = double(coreID_R);

if ismember("Salinity_used", rho.Properties.VariableNames)
    salRowR = toNum(rho.Salinity_used);
elseif ismember("salinity", rho.Properties.VariableNames)
    salRowR = toNum(rho.salinity);
else
    salRowR = NaN(height(rho),1);
end

if ismember("rho_lab_kgm3", rho.Properties.VariableNames)
    denLabRow = toNum(rho.rho_lab_kgm3);
else
    denLabRow = toNum(rho{:,3}) * 1000;
end

if ismember("rho_si", rho.Properties.VariableNames)
    denInsRow = toNum(rho.rho_si);
else
    denInsRow = NaN(height(rho),1);
end

if ismember("Temperature_interp", rho.Properties.VariableNames)
    tempRow = toNum(rho.Temperature_interp);
else
    tempRow = NaN(height(rho),1);
end

coreListR = unique(coreID_R(~isnan(coreID_R) & coreID_R>0), "stable");
nR = numel(coreListR);

CoreTimeR = NaT(nR,1);
CoreStationR = NaN(nR,1);
AvgSalR = NaN(nR,1);
AvgLab  = NaN(nR,1);
AvgIns  = NaN(nR,1);
AvgTemp = NaN(nR,1);

for i = 1:nR
    idx = coreID_R == coreListR(i);
    CoreTimeR(i)    = earliestNonMissing(tRowR(idx));
    CoreStationR(i) = firstNonNaN(stRowR(idx));
    AvgSalR(i)  = mean(salRowR(idx),'omitnan');
    AvgLab(i)   = mean(denLabRow(idx),'omitnan');
    AvgIns(i)   = mean(denInsRow(idx),'omitnan');
    AvgTemp(i)  = mean(tempRow(idx),'omitnan');
end

hasSALO = isfield(T_all_proc,"SALO18") && ~isempty(T_all_proc.SALO18);
CoreTimeS = NaT(0,1);
CoreStationS = NaN(0,1);
AvgSalS = NaN(0,1);

if hasSALO
    salo = T_all_proc.SALO18;

    if ismember("Date", salo.Properties.VariableNames)
        tRowS = salo.Date;
        if isa(tRowS,'datetime') && ~isempty(tRowS.TimeZone), tRowS.TimeZone=''; end
    else
        tRowS = NaT(height(salo),1);
        if ismember("Time_best", salo.Properties.VariableNames) && isa(salo.Time_best,'datetime')
            tRowS = salo.Time_best;
            if ~isempty(tRowS.TimeZone), tRowS.TimeZone=''; end
        end
        if ismember("GPS_Time", salo.Properties.VariableNames) && isa(salo.GPS_Time,'datetime')
            gg = salo.GPS_Time;
            if ~isempty(gg.TimeZone), gg.TimeZone=''; end
            miss = ismissing(tRowS) & ~ismissing(gg);
            tRowS(miss) = gg(miss);
        end
    end

    if ~ismember("StationNumber", salo.Properties.VariableNames)
        warning("SALO18 has no StationNumber. SALO18 salinity will be skipped.");
        hasSALO = false;
    else
        stRowS = double(salo.StationNumber);

        if ismember("Salinity", salo.Properties.VariableNames)
            salRowS = toNum(salo.Salinity);
        else
            salRowS = toNum(salo{:,3});
        end

        if ismember("CoreID_SALO18", salo.Properties.VariableNames)
            coreID_S = double(salo.CoreID_SALO18);
        else
            [~,~,coreID_S] = unique(string(salo.SourceFile), "stable");
            coreID_S = double(coreID_S);
        end

        coreListS = unique(coreID_S(~isnan(coreID_S) & coreID_S>0), "stable");
        nS = numel(coreListS);

        CoreTimeS = NaT(nS,1);
        CoreStationS = NaN(nS,1);
        AvgSalS = NaN(nS,1);

        for i = 1:nS
            idx = coreID_S == coreListS(i);
            CoreTimeS(i)    = earliestNonMissing(tRowS(idx));
            CoreStationS(i) = firstNonNaN(stRowS(idx));
            AvgSalS(i)      = mean(salRowS(idx),'omitnan');
        end
    end
end

stations = unique([CoreStationR; CoreStationS]);
stations = stations(~isnan(stations));
stations = stations(ismember(stations,[1 2 3]));
stations = sort(stations);

cols = lines(3);

figure('Color','w');
tiledlayout(3,1,'Padding','compact','TileSpacing','compact');

ax1 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
    idxR = CoreStationR==st & ~isnan(AvgSalR) & ~ismissing(CoreTimeR);
    scatter(CoreTimeR(idxR), AvgSalR(idxR), 70, 'o', ...
        'MarkerFaceColor',c,'MarkerEdgeColor',c);

    if hasSALO
        idxS = CoreStationS==st & ~isnan(AvgSalS) & ~ismissing(CoreTimeS);
        scatter(CoreTimeS(idxS), AvgSalS(idxS), 70, '^', ...
            'MarkerFaceColor','none','MarkerEdgeColor',c,'LineWidth',1.2);
    end
end
ylabel("Ice salinity");
title("Salinity vs time");

ax2 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
% Lab density → OPEN circle
idxL = CoreStationR==st & ~isnan(AvgLab) & ~ismissing(CoreTimeR);
scatter(CoreTimeR(idxL), AvgLab(idxL), 70, 'o', ...
    'MarkerFaceColor','none','MarkerEdgeColor',c,'LineWidth',1.2);

% In situ density → FILLED circle
idxI = CoreStationR==st & ~isnan(AvgIns) & ~ismissing(CoreTimeR);
scatter(CoreTimeR(idxI), AvgIns(idxI), 70, 'o', ...
    'MarkerFaceColor',c,'MarkerEdgeColor',c);
end
ylabel("Ice density (kg m^{-3})");
title("Density vs time");

ax3 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
    idxT = CoreStationR==st & ~isnan(AvgTemp) & ~ismissing(CoreTimeR);
    scatter(CoreTimeR(idxT), AvgTemp(idxT), 70, 's', ...
        'MarkerFaceColor',c,'MarkerEdgeColor',c);
end
ylabel("Ice temperature (°C)");
title("Temperature vs time");

% ---------- Legend (OLD placement: attached to panel 1, shown above it) ----------
hSt = gobjects(0,1);
labSt = strings(0,1);
for st = stations'
    hSt(end+1,1) = scatter(ax1, nan,nan,70,'s','filled', ...
        'MarkerFaceColor',cols(st,:), 'MarkerEdgeColor',cols(st,:));
    labSt(end+1,1) = "Station " + string(st);
end

hRHO   = scatter(ax1, nan,nan,70,'o','filled','MarkerFaceColor','k','MarkerEdgeColor','k');
hSALO  = scatter(ax1, nan,nan,70,'^','MarkerFaceColor','none','MarkerEdgeColor','k','LineWidth',1.2);
hLAB   = scatter(ax1, nan,nan,70,'o','MarkerFaceColor','none','MarkerEdgeColor','k','LineWidth',1.2);
hINS   = scatter(ax1, nan,nan,70,'o','filled','MarkerFaceColor','k','MarkerEdgeColor','k');
hTEMP  = scatter(ax1, nan,nan,70,'s','filled','MarkerFaceColor','k','MarkerEdgeColor','k');

hAll  = [hSt; hRHO; hSALO; hLAB; hINS; hTEMP];
labAll = [labSt; "RHO"; "SALO18"; "Lab density"; "In situ density"; "Temperature"];

legend(ax1, hAll, labAll, 'Location','northoutside','NumColumns',5);

set(gcf,'Units','inches','Position',[4 4 8 8])
exportgraphics(gcf,"Density_salinity_temperature_vs_time.png","Resolution",300);
fprintf("Saved: Density_salinity_temperature_vs_time.png\n");

%% Counter of data sheets
rootFolder = "C:\Users\evsalg001\Documents\MATLAB\Contrasts coring\Data";
files = dir(fullfile(rootFolder, '**', '*.xlsx'));
names = string({files.name});
n_RHO    = sum(contains(names, "-RHO"));
n_SALO18 = sum(contains(names, "-SALO18"));
n_T      = sum(contains(names, "-T.xlsx"));
fprintf('RHO files: %d\n', n_RHO)
fprintf('SALO18 files: %d\n', n_SALO18)
fprintf('T files: %d\n', n_T)

%% =========================== Helpers ===========================

function T_all = ensureNewCols(T_all, MAP)
    if isfield(T_all,'rho') && ~isempty(T_all.rho)
        if ~ismember(MAP.StationNumber, T_all.rho.Properties.VariableNames), T_all.rho.(MAP.StationNumber) = NaN(height(T_all.rho),1); end
        if ~ismember(MAP.StationVisit,  T_all.rho.Properties.VariableNames), T_all.rho.(MAP.StationVisit)  = repmat("", height(T_all.rho),1); end
        T_all.rho.(MAP.StationVisit) = string(T_all.rho.(MAP.StationVisit));
        if ~ismember(MAP.CoreID_RHO, T_all.rho.Properties.VariableNames),     T_all.rho.(MAP.CoreID_RHO)     = NaN(height(T_all.rho),1); end
    end
    if isfield(T_all,'T') && ~isempty(T_all.T)
        if ~ismember(MAP.StationNumber, T_all.T.Properties.VariableNames),   T_all.T.(MAP.StationNumber) = NaN(height(T_all.T),1); end
        if ~ismember(MAP.StationVisit,  T_all.T.Properties.VariableNames),   T_all.T.(MAP.StationVisit)  = repmat("", height(T_all.T),1); end
        T_all.T.(MAP.StationVisit) = string(T_all.T.(MAP.StationVisit));
        if ~ismember(MAP.CoreID_T, T_all.T.Properties.VariableNames),        T_all.T.(MAP.CoreID_T)       = NaN(height(T_all.T),1); end
    end
    if isfield(T_all,'SALO18') && ~isempty(T_all.SALO18)
        if ~ismember(MAP.StationNumber, T_all.SALO18.Properties.VariableNames), T_all.SALO18.(MAP.StationNumber) = NaN(height(T_all.SALO18),1); end
        if ~ismember(MAP.StationVisit,  T_all.SALO18.Properties.VariableNames), T_all.SALO18.(MAP.StationVisit)  = repmat("", height(T_all.SALO18),1); end
        T_all.SALO18.(MAP.StationVisit) = string(T_all.SALO18.(MAP.StationVisit));
        if ~ismember(MAP.CoreID_SALO18, T_all.SALO18.Properties.VariableNames), T_all.SALO18.(MAP.CoreID_SALO18) = NaN(height(T_all.SALO18),1); end
    end
end

function key = filePairKey(sourceFile, kind)
    % kind: "T" or "RHO"
    f = string(sourceFile);

    % Strip extension
    f = regexprep(f, "\.xlsx$", "", "ignorecase");

    % Remove trailing type token
    if strcmpi(kind,"T")
        f = regexprep(f, "-T$", "", "ignorecase");
    else
        % handles "...-RHO" and "...-RHO-TXT"
        f = regexprep(f, "-RHO(-TXT)?$", "", "ignorecase");
    end

    % Remove the numeric token immediately before the ice-age/label chunk.
    % This fixes 016 vs 019 and 020 vs 022 while keeping "lead" distinction.
    % Example:
    %   ...-016-FYI          -> ...-FYI
    %   ...-022-FYI-lead     -> ...-FYI-lead
    f = regexprep(f, "-\d+(?=-[A-Za-z])", "", "once");

    % Clean accidental double dashes
    f = regexprep(f, "--+", "-");

    key = f;
end

function tbl = keepAndOrder(tbl, preferredVars)
    preferredVars = preferredVars(preferredVars ~= "");
    keep = preferredVars(ismember(preferredVars, tbl.Properties.VariableNames));
    tbl = tbl(:, keep);
end

function tbl = applyRenameMap(tbl, renameMap)
    for i = 1:size(renameMap,1)
        oldName = renameMap{i,1};
        newName = renameMap{i,2};
        if ismember(oldName, tbl.Properties.VariableNames)
            tbl = renamevars(tbl, oldName, newName);
        end
    end
end

function vname = pickOrDefault(tbl, preferredName, defaultName)
    vname = "";
    if preferredName ~= "" && ismember(preferredName, tbl.Properties.VariableNames)
        vname = preferredName;
        return
    end
    if defaultName ~= "" && ismember(defaultName, tbl.Properties.VariableNames)
        vname = defaultName;
        return
    end
end

function vname = guessVar(tbl, patterns, fallback)
    names = string(tbl.Properties.VariableNames);
    vname = fallback;
    for p = patterns
        idx = find(contains(lower(names), lower(p)), 1, 'first');
        if ~isempty(idx)
            vname = names(idx);
            return
        end
    end
end

function tbl = addMeltPond01(tbl, meltVar)
    if ismember(meltVar, tbl.Properties.VariableNames)
        try
            tbl.MeltPond01 = double(tbl.(meltVar));
        catch
            tbl.MeltPond01 = double(contains(string(tbl.(meltVar)), "true", 'IgnoreCase', true) | ...
                                    contains(string(tbl.(meltVar)), "1"));
        end
    else
        tbl.MeltPond01 = NaN(height(tbl),1);
    end
end

function tbl = addTimeBest(tbl, gpsTimeVar)
    n = height(tbl);
    tbl.Time_best = NaT(n,1);

    if ismember(gpsTimeVar, tbl.Properties.VariableNames) && isa(tbl.(gpsTimeVar),'datetime')
        t = tbl.(gpsTimeVar);
        if ~isempty(t.TimeZone), t.TimeZone = ''; end
        tbl.Time_best = t;
    end

    isDT   = varfun(@(x) isa(x,'datetime'), tbl, 'OutputFormat','uniform');
    dtVars = string(tbl.Properties.VariableNames(isDT));
    dtVars = setdiff(dtVars, ["Time_best", string(gpsTimeVar)], 'stable');

    if ~isempty(dtVars)
        fb = tbl.(dtVars(1));
        if isa(fb,'datetime') && ~isempty(fb.TimeZone), fb.TimeZone = ''; end
        miss = ismissing(tbl.Time_best);
        tbl.Time_best(miss) = fb(miss);
    end
end

function [desc, num] = parseCoreDescriptorAndNumber(fname)
    % Example:
    % 20250726-PS149_21-1-SI_corer_9cm-020-FYI-lead-T.xlsx
    % -> num = 20
    % -> desc = "FYI-lead"
    f = string(fname);

    % remove extension
    f = regexprep(f, "\.xlsx$", "", "ignorecase");

    % remove trailing type
    f = regexprep(f, "-(T|RHO(-TXT)?)$", "", "ignorecase");

    % capture number after SI_corer_?cm-
    tokNum = regexp(f, "SI_corer_\d+cm-(\d+)", "tokens", "once");
    if isempty(tokNum)
        num = NaN;
    else
        num = str2double(tokNum{1});
    end

    % descriptor = everything after that number + dash
    % e.g. "...SI_corer_9cm-020-FYI-lead" -> "FYI-lead"
    tokDesc = regexp(f, "SI_corer_\d+cm-\d+-(.+)$", "tokens", "once");
    if isempty(tokDesc)
        desc = "";
    else
        desc = string(tokDesc{1});
    end
end

function x = toNum(x)
    if iscell(x) || isstring(x)
        x = str2double(string(x));
    end
    x = double(x);
end

function y = firstNonNaN(x)
    x = double(x);
    k = find(~isnan(x), 1, "first");
    if isempty(k), y = NaN; else, y = x(k); end
end

function v = guessVarName(vars, patterns)
    v = "";
    low = lower(vars);
    for p = patterns
        hit = find(contains(low, lower(p)), 1, "first");
        if ~isempty(hit)
            v = vars(hit);
            return
        end
    end
end

function y = earliestNonMissing(x)
    x = x(~ismissing(x));
    if isempty(x), y = NaT; else, y = min(x); end
end

%% FUNCTION: Write one table into a NetCDF group
function writeTableGroupNetCDF(ncFile, grp, T, metaMap)

if isempty(T) || height(T)==0
    warning("Group %s: table empty, skipping.", grp);
    return
end

nObs  = height(T);
obsDim = "obs";
vars  = string(T.Properties.VariableNames);

for i = 1:numel(vars)
    origName = vars(i);
    ncName   = sanitizeVarName(origName);
    vpath    = grp + "/" + ncName;

    x = T.(origName);

    % ---- Time handling ----
    if isTimeColumn(origName)
        tDays = convertTimeToDaysSince1979(x);

        nccreate(ncFile, vpath, "Dimensions", {obsDim, nObs}, "Datatype", "double");
        ncwrite(ncFile, vpath, tDays);

        ncwriteatt(ncFile, vpath, "standard_name", "time");
        ncwriteatt(ncFile, vpath, "long_name", char(origName));
        ncwriteatt(ncFile, vpath, "units", "days since 1979-01-01 00:00:00");
        ncwriteatt(ncFile, vpath, "calendar", "standard");
        continue
    end

    % ---- Write data ----
    if isnumeric(x) || islogical(x)
        x = x(:);
        nccreate(ncFile, vpath, "Dimensions", {obsDim, nObs}, "Datatype", "double");
        ncwrite(ncFile, vpath, double(x));
    else
        % Strings -> char(strlen, obs)
        s = string(x);
        s(ismissing(s)) = "";
        maxLen = max(strlength(s));
        if maxLen < 1, maxLen = 1; end

        strlenDim = "strlen_" + ncName;

        % create group-local strlen dim placeholder (shows up in file)
        nccreate(ncFile, grp + "/" + strlenDim, "Dimensions", {strlenDim, maxLen}, "Datatype", "int32");

        nccreate(ncFile, vpath, "Dimensions", {strlenDim, maxLen, obsDim, nObs}, "Datatype", "char");

        C = char(pad(s, maxLen));  % nObs x maxLen
        C = permute(C, [2 1]);     % maxLen x nObs
        ncwrite(ncFile, vpath, C);
    end

    % Default long_name = table header
    ncwriteatt(ncFile, vpath, "long_name", char(origName));

% ---- Apply metadata (units/standard_name/comment) from map ----
key1 = char(origName);   % table header (long_name)
key2 = char(ncName);     % sanitized netcdf variable name

hasMeta = false;
if isKey(metaMap, key1)
    m = metaMap(key1); hasMeta = true;
elseif isKey(metaMap, key2)
    m = metaMap(key2); hasMeta = true;
end

if hasMeta
    if isfield(m, "standard_name"), ncwriteatt(ncFile, vpath, "standard_name", m.standard_name); end
    if isfield(m, "units"),         ncwriteatt(ncFile, vpath, "units",         m.units); end
    if isfield(m, "comment"),       ncwriteatt(ncFile, vpath, "comment",       m.comment); end
else
    if isnumeric(x) || islogical(x)
        ncwriteatt(ncFile, vpath, "units", "1");
    end
end

    % ---- Extra CF-style categorical attributes ----
    if origName == "Ice age"
        ncwriteatt(ncFile, vpath, "category_values", "FYI SYI SMYI");
        ncwriteatt(ncFile, vpath, "category_meanings", ...
            "first_year_ice second_year_ice second_or_multiyear_ice");
    elseif origName == "Melt pond"
        ncwriteatt(ncFile, vpath, "flag_values", int8([0 1]));
        ncwriteatt(ncFile, vpath, "flag_meanings", "no yes");
    elseif origName == "Ice station visit"
        ncwriteatt(ncFile, vpath, "category_values", "a b c d");
        ncwriteatt(ncFile, vpath, "category_meanings", "visit_a visit_b visit_c visit_d");
    end
end
end

% FUNCTION: Identify time column names
function tf = isTimeColumn(varName)
vn = string(varName);
tf = any(vn == ["DATE/TIME","Time_best","time","TIME","DateTime","datetime"]);
end

% FUNCTION: Convert time column to days since 1979-01-01 00:00:00 UTC
function tDays = convertTimeToDaysSince1979(x)
ref = datetime(1979,1,1,0,0,0,"TimeZone","UTC");

if isdatetime(x)
    t = x;
else
    s = string(x);
    s(ismissing(s)) = "";
    % Adjust InputFormat if needed:
    try
        t = datetime(s, "TimeZone","UTC", "InputFormat","yyyy-MM-dd HH:mm:ss");
    catch
        t = datetime(s, "TimeZone","UTC");
    end
end

t = datetime(t, "TimeZone","UTC");
tDays = double(days(t - ref));
end

% FUNCTION: Sanitize table header to a NetCDF-safe variable name
function ncName = sanitizeVarName(name)
s = string(name);
s = regexprep(s, "\s+", "_");           % spaces -> _
s = regexprep(s, "[^A-Za-z0-9_]", "_"); % commas/slashes/etc -> _
s = regexprep(s, "_+", "_");            % collapse
s = regexprep(s, "^_+|_+$", "");        % trim

if s == "", s = "var"; end
if ~isempty(regexp(s, "^[0-9]", "once"))
    s = "v_" + s;
end
ncName = char(s);
end

% FUNCTION: Metadata map for RHO table (units + comments)
function M = buildMetaMap_RHO()
M = containers.Map();

% core coords / time
M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east");
M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");
M("IceThickness") = struct("standard_name","sea_ice_thickness","units","m");
M("Draft")        = struct("standard_name","sea_ice_draft","units","m");
M("Core_length")  = struct("standard_name","core_length","units","m");
M("Depth_ice_snow") = struct("standard_name","depth","units","m");
M("Depth, ice/snow, top/minimum")    = struct("standard_name","depth","units","m");
M("Depth, ice/snow, bottom/maximum") = struct("standard_name","depth","units","m");
M("Sea ice salinity") = struct("standard_name","sea_ice_salinity","units","PSU");
M("Temperature, technical") = struct("standard_name","temperature","units","degree_Celsius","comment","Air temperature in laboratory.");
M("Temperature, ice/snow") = struct("standard_name","sea_ice_temperature","units","degree_Celsius","comment","In situ ice temperature.");
M("Density, ice, technical") = struct("standard_name","sea_ice_density","units","kg/m3","comment","Sea ice density measured in laboratory.");
M("Density, ice") = struct("standard_name","sea_ice_density","units","kg/m3","comment","Sea ice density estimated for in situ temperature.");
M("Volume, gas, technical") = struct("standard_name","gas_volume_fraction","units","1","comment","Gas volume fraction estimated for laboratory temperature.");
M("Volume, gas") = struct("standard_name","gas_volume_fraction","units","1","comment","Gas volume fraction estimated for in situ temperature.");
end

% FUNCTION: Metadata map for T table (units + comments)
function M = buildMetaMap_T()
M = containers.Map();
M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east");
M("IceThickness") = struct("standard_name","sea_ice_thickness","units","m");
M("Draft")        = struct("standard_name","sea_ice_draft","units","m");
M("Core_length")     = struct("standard_name","core_length","units","m");
M("Depth_ice_snow")  = struct("standard_name","depth","units","m");
M("Ice age") = struct("standard_name","ice_age", "comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment", "Ice stations 1, 2, 3.");
M("Temperature, ice/snow") = struct("standard_name","sea_ice_temperature","units","degree_Celsius","comment", ...
    "In situ ice temperature.");
end

function M = buildMetaMap_SALO18()
M = containers.Map();
M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east");
M("IceThickness") = struct("standard_name","sea_ice_thickness","units","m");
M("Draft")        = struct("standard_name","sea_ice_draft","units","m");
M("Core_length")  = struct("standard_name","core_length","units","m");
M("Depth_ice_snow")  = struct("standard_name","depth","units","m");
M("Depth, ice/snow, top/minimum")    = struct("standard_name","depth","units","m");
M("Depth, ice/snow, bottom/maximum") = struct("standard_name","depth","units","m");
M("Sea ice salinity") = struct("standard_name","sea_ice_salinity","units","PSU");
M("Ice age") = struct("standard_name","ice_age", "comment", ...
    "First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment", ...
    "Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment", ...
    "Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment", ...
    "Ice stations 1, 2, 3.");
end