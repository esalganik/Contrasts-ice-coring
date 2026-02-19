%% Mass Import of RHO, T, and SALO18 cores with GPS
clear; close all; clc
tStart = tic;

rootFolder = "C:\Users\evsalg001\Documents\MATLAB\Contrasts coring\Data"; % Folder with coring protocols
filePatterns = ["*-RHO.xlsx","*-RHO-TXT.xlsx","*-SALO18.xlsx","*-T.xlsx"]; % Names of imported cores
allFiles = [];

for fp = filePatterns
    files = dir(fullfile(rootFolder,"**",fp));
    allFiles = [allFiles; files];
end

% Initialize empty tables for each core type
T_all = struct('rho', table(), 'T', table(), 'SALO18', table());

for k = 1:numel(allFiles)

    filePath = fullfile(allFiles(k).folder, allFiles(k).name);

    pathParts = split(allFiles(k).folder, filesep);
    station   = pathParts{end-1};
    coreName  = pathParts{end};
    extraMeta = struct("Station",station,"Core",coreName);

    % RHO cores
    if contains(filePath,'-RHO.xlsx') || contains(filePath,'-RHO-TXT.xlsx')

        opts = detectImportOptions(filePath,'Sheet','Density-densimetry','VariableNamingRule','preserve');
        opts.SelectedVariableNames = opts.VariableNames([1 2 7 8]);
        T = readtable(filePath, opts);

        % Convert numeric columns
        varNames = T.Properties.VariableNames;
        for c = 1:numel(varNames)
            if c==4, continue, end
            v = T.(varNames{c});
            if iscell(v) || isstring(v)
                T.(varNames{c}) = str2double(v);
            end
        end

        % Parse comments
        comments = string(T{:,4});
        rho_all = str2double(regexp(comments, '(?i)(?<=parafine[^=]*=\s*)[-+]?\d*\.?\d+', 'match','once'));
        Tlab_all = str2double(regexp(comments, '(?i)(?<=lab[^=]*=\s*)[-+]?\d*\.?\d+', 'match','once'));
        rho_all  = fillmissing(rho_all,'previous');
        Tlab_all = fillmissing(Tlab_all,'previous');
        f = find(~isnan(rho_all),1,'first'); if ~isempty(f), rho_all(1:f-1)=rho_all(f); end
        f = find(~isnan(Tlab_all),1,'first'); if ~isempty(f), Tlab_all(1:f-1)=Tlab_all(f); end
        T.rho_parafine = rho_all;
        T.Tlab         = Tlab_all;
        T = T(~isnan(T{:,1}), :);

        % SALO18 safe import
        try
            T_salo = readtable(filePath,'Sheet','SALO18','VariableNamingRule','preserve');
            if size(T_salo,2)>=5
                col5 = T_salo{:,5};
                if iscell(col5) || isstring(col5)
                    col5 = str2double(col5);
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

        % --- Add GPS (first waypoint) ---
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        if ~isempty(newTbl)
            T_all.rho = [T_all.rho; newTbl];
        end

    % T cores
    elseif contains(filePath,'-T.xlsx')

        T = readtable(filePath,'Sheet','TEMP','Range','A2:B1000','ReadVariableNames',true,'VariableNamingRule','preserve');
        T = forceNumericTable(T);
        T = T(~isnan(T{:,1}), :);

        newTbl = add_metadata(T,filePath,{'A7','A8','A10'},{'C7','C8','C10'},extraMeta);

        % --- Add GPS (first waypoint) ---
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        if ~isempty(newTbl)
            T_all.T = [T_all.T; newTbl];
        end

    % SALO18 cores
    elseif contains(filePath,'-SALO18.xlsx')
        opts_S = detectImportOptions(filePath,'Sheet','SALO18','VariableNamingRule','preserve');
        opts_S.SelectedVariableNames = opts_S.VariableNames([1 2 5]);
        T = readtable(filePath, opts_S);
        T = forceNumericTable(T);
        T = T(~isnan(T{:,1}), :);
        if ~isempty(T)
            T.Salinity = T{:,3}; T(:,3) = [];
        end
        newTbl = add_metadata(T,filePath,{'A7','A8','A10'},{'C7','C8','C10'},extraMeta);

        % --- Add GPS (first waypoint) ---
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        if ~isempty(newTbl)
            T_all.SALO18 = [T_all.SALO18; newTbl];
        end
    end
end

% --- Helper functions ---
function T = forceNumericTable(T)
    vars = T.Properties.VariableNames;
    for i = 1:numel(vars)
        v = T.(vars{i});
        if iscell(v) || isstring(v)
            T.(vars{i}) = str2double(v);
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
        if ischar(val) || isstring(val), val = str2double(val); end
        if isnumeric(val) && contains(lower(varname),'date')
            val = datetime(val,'ConvertFrom','excel');
        end
        if isempty(val) || (~isnumeric(val) && ~isdatetime(val)), val = NaN; end
        tbl.(varname) = repmat(val,height(tbl),1);
    end
    % Extra metadata
    fn = fieldnames(extraMeta);
    for i = 1:numel(fn)
        tbl.(fn{i}) = repmat(string(extraMeta.(fn{i})),height(tbl),1);
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
    % --- Event label ---
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
    % Default outputs
    lat = nan(nRows,1);
    lon = nan(nRows,1);
    t_gps = NaT(nRows,1,'TimeZone','UTC');  % <-- set UTC timezone even if empty

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
                warning('Failed to read GPX: %s', gpsFile);
            end
        end
    end
end

% --- Finish ---
clear newTbl varNames v T_salo filePatterns c v rootFolder allFiles k fp filePath files pathParts station coreName extraMeta T opts opts_S col5 comments rho_all Tlab_all f
elapsed = toc(tStart);
fprintf('Mass import completed in %.0f seconds.\n', elapsed);
clear elapsed tStart
save('Coring_data_imported.mat','T_all');
fprintf('Imported data saved to Coring_data_imported.mat\n');

%% Load imported -> process -> format/reorder/drop -> rename -> save
% Output:
%   T_all_proc.rho          (original + processed columns)
%   T_all_proc.rho_out      (ONLY selected columns, preferred order, renamed)
%   T_all_proc.T_out        (ONLY selected columns, preferred order, renamed)
%   T_all_proc.SALO18_out   (ONLY selected columns, preferred order, renamed)

clear; close all; clc

% ============ USER SETTINGS ============
MAP = struct();

% Common/meta
MAP.EventLabel   = "EventLabel";
MAP.GPS_Time     = "GPS_Time";
MAP.GPS_Lat      = "GPS_Lat";
MAP.GPS_Lon      = "GPS_Lon";
MAP.Station      = "Station";
MAP.IceAge       = "IceAge";
MAP.MeltPond     = "MeltPond";
MAP.Comments     = "";

% RHO-specific
MAP.Depth1       = "";
MAP.Depth2       = "";
MAP.Salinity_raw = "salinity";
MAP.Salinity_used= "Salinity_used";
MAP.Tlab         = "Tlab";
MAP.T_ice        = "Temperature_interp";
MAP.Rho_si       = "rho_si"; % computed density (kg/m3)

MAP.Vb_export    = "vb_rho_export"; % NaN if raw >0.4 or <0
MAP.Vg_pr        = "vg_pr";
MAP.Vg           = "vg";

% Geometry (optional, guessed if empty)
MAP.IceThickness = "";              % metadata column if exists
MAP.IceDraft     = "";              % metadata column if exists
MAP.CoreLength   = "";              % metadata column if exists

% T-core specific (guessed if empty)
MAP.T_Depth      = "";              % depth column in T table
MAP.T_Temp       = "";              % temperature column in T table

% SALO18 specific (guessed if empty)
MAP.S_Depth1     = "";              % if SALO18 has two depth columns, else depth
MAP.S_Depth2     = "";
MAP.S_Salinity   = "";              % salinity column in SALO18 table

OUTFILE_PROCESSED = "Coring_data_processed.mat";

% 1) Load imported data
load('Coring_data_imported.mat','T_all');
fprintf('Imported data loaded from MAT-file\n');

% Normalize to strings where relevant
T_all.T.SourceFile   = string(T_all.T.SourceFile);
T_all.rho.SourceFile = string(T_all.rho.SourceFile);

T_all.T.Station      = string(T_all.T.Station);
T_all.T.Core         = string(T_all.T.Core);
T_all.rho.Station    = string(T_all.rho.Station);
T_all.rho.Core       = string(T_all.rho.Core);

if isfield(T_all,'SALO18') && ~isempty(T_all.SALO18)
    if ismember('Station', T_all.SALO18.Properties.VariableNames)
        T_all.SALO18.Station = string(T_all.SALO18.Station);
    end
    if ismember('Core', T_all.SALO18.Properties.VariableNames)
        T_all.SALO18.Core = string(T_all.SALO18.Core);
    end
end

% 2) Make processed copy
T_all_proc = T_all;

needCols = ["rho_si","rho_lab_kgm3","Salinity_used","Temperature_interp", ...
            "vb_rho_export","vg_pr","vg"];

for c = needCols
    if ~ismember(c, T_all_proc.rho.Properties.VariableNames)
        T_all_proc.rho.(c) = NaN(height(T_all_proc.rho),1);
    end
end

% 3) Match T and RHO cores
Tmatch = table(strings(0,1), strings(0,1), strings(0,1), strings(0,1), ...
               'VariableNames', {'Station','Core','T_File','RHO_File'});

uniqueTFiles = unique(T_all.T.SourceFile);

for i = 1:numel(uniqueTFiles)
    tFile = uniqueTFiles(i);

    tIdx  = find(T_all.T.SourceFile == tFile, 1);
    tCore = T_all.T(tIdx,:);
    st    = string(tCore.Station);
    co    = string(tCore.Core);

    idxR = find(T_all.rho.Station == st & T_all.rho.Core == co);
    if isempty(idxR)
        warning("No RHO core found for Station %s / Core %s", st, co);
        continue
    end

    baseT = regexprep(tFile, "^\d{8}-PS\d+_\d+-\d+-SI_corer_\d+cm-\d+-", "");
    baseT = regexprep(baseT, "-T.*\.xlsx$", "");

    chosenR = [];
    for r = idxR'
        rFile = T_all.rho.SourceFile(r);
        baseR = regexprep(rFile, "^\d{8}-PS\d+_\d+-\d+-SI_corer_\d+cm-\d+-", "");
        baseR = regexprep(baseR, "-RHO.*\.xlsx$", "");
        if strcmpi(baseT, baseR)
            chosenR = r;
            break
        end
    end

    if isempty(chosenR)
        chosenR = idxR(1);
        warning("Fallback used for Station %s / Core %s", st, co);
    end

    Tmatch = [Tmatch; table(st, co, tFile, T_all.rho.SourceFile(chosenR), ...
        'VariableNames', {'Station','Core','T_File','RHO_File'})];
end

fprintf('Core matching completed: %d pairs\n', height(Tmatch));

% 4) Compute rho_si + vb/vg and write back into T_all_proc.rho
nPairs = height(Tmatch);
rho_si_bulk = nan(nPairs,1);
T_bulk      = nan(nPairs,1);

for i = 1:nPairs
    % fprintf('\n--- Processing pair %d / %d ---\n', i, nPairs);
    % fprintf('T   : %s\n', Tmatch.T_File(i));
    % fprintf('RHO : %s\n', Tmatch.RHO_File(i));

    T_T   = T_all.T(  T_all.T.SourceFile   == Tmatch.T_File(i), :);
    T_rho = T_all.rho(T_all.rho.SourceFile == Tmatch.RHO_File(i), :);

    if isempty(T_T) || isempty(T_rho)
        warning('Missing data for this pair — skipping');
        continue
    end

    % --- Depth midpoints for RHO ---
    depth_rho = mean(T_rho{:,1:2}, 2);

    % --- Temperature profile ---
    depth_T = T_T{:,1};
    temp    = min(-0.1, T_T{:,2});
    ok = ~isnan(depth_T) & ~isnan(temp);
    depth_T = depth_T(ok);
    temp    = temp(ok);

    if numel(depth_T) < 2
        warning('Too few temperature points — skipping');
        continue
    end

    % --- Rescale + interpolate temperature to RHO depths ---
    depth_T_rescaled = depth_T * (max(depth_rho) / max(depth_T));
    T_interp = interp1(depth_T_rescaled, temp, depth_rho, 'linear', 'extrap');

    % --- Salinity used: from RHO salinity column ---
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

    % --- SALO18 fallback if all NaN ---
    if all(isnan(Srho))
        SALOcore = [];
        if isfield(T_all,'SALO18') && ~isempty(T_all.SALO18)
            SALOcore = T_all.SALO18(T_all.SALO18.Station == Tmatch.Station(i), :);
        end
        if ~isempty(SALOcore)
            depth_SALO = mean(SALOcore{:,1:2}, 2);
            sal_SALO   = SALOcore{:,3};
            Srho = interp1(depth_SALO, sal_SALO, depth_rho, 'linear', 'extrap');
            % fprintf('  -> Using SALO18 salinity fallback\n');
        else
            warning('No SALO18 core found — using NaN for salinity');
        end
    end

    % --- Lab temperature ---
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

    % --- Measured density at lab state (assume column 3 is g/cm^3) ---
    rho_meas = T_rho{:,3};
    rho_lab_kgm3 = rho_meas * 1000;  % kg/m^3 used for calcs + saved
    rho = rho_lab_kgm3;

    % ===================== CALCULATIONS (Cox and Weeks (1983) and Leppäranta and Manninen (1988)) =====================
    F1_pr_rho = -4.732 - 22.45*T_lab - 0.6397*T_lab.^2 - 0.01074*T_lab.^3;
    F2_pr_rho = 8.903e-2 - 1.763e-2*T_lab - 5.33e-4*T_lab.^2 - 8.801e-6*T_lab.^3;
    vb_pr_rho = rho .* Srho ./ F1_pr_rho;
    rhoi_pr   = 917 - 0.1403*T_lab;
    vg_pr     = max(0, 1 - rho .* (F1_pr_rho - rhoi_pr .* Srho/1000 .* F2_pr_rho) ./ (rhoi_pr .* F1_pr_rho));
    F3_pr     = rhoi_pr .* Srho/1000 ./ (F1_pr_rho - rhoi_pr .* Srho/1000 .* F2_pr_rho);

    T_insitu = T_interp;
    rhoi_rho = 917 - 0.1403*T_insitu;
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

    % --- Compute raw vb (no clipping yet) ---
    vb_rho_raw = vb_pr_rho .* F1_pr_rho ./ F1_rho / 1000;

    % --- Export version: NaN if raw >0.4 or raw <0 ---
    vb_rho_export = vb_rho_raw;
    vb_rho_export(vb_rho_export > 0.4 | vb_rho_export < 0) = NaN;

    % --- Calculation version: clip to 0.4 as before ---
    vb_rho = vb_rho_raw;
    vb_rho(vb_rho > 0.4 | vb_rho < 0) = 0.4;

    vg = max(0, (1 - (1 - vg_pr) .* (rhoi_rho./rhoi_pr) .* (F3_pr.*F1_pr_rho./(F3_rho.*F1_rho))));

    rho_si = (1 - vg) .* rhoi_rho .* F1_rho ./ (F1_rho - rhoi_rho .* Srho/1000 .* F2_rho);
    rho_si(isnan(vb_rho)) = NaN;
    % ================================================================================

    rho_si_bulk(i) = mean(rho_si,'omitnan');
    T_bulk(i)      = mean(T_insitu,'omitnan');

    % fprintf('Bulk rho_si = %.1f kg/m3\n', rho_si_bulk(i));
    % fprintf('Bulk T      = %.2f C\n', T_bulk(i));

    % --- Write back into exact rows in T_all_proc.rho ---
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

% fprintf('\nPair processing completed for %d cores\n', nPairs);

% 5) FORMAT / REORDER / DROP columns for final exports
% ---------- RHO OUT ----------
rho = T_all_proc.rho;

rho = addTimeBest(rho, MAP.GPS_Time);
rho = addMeltPond01(rho, MAP.MeltPond);

d1 = pickOrDefault(rho, MAP.Depth1, rho.Properties.VariableNames{1});
d2 = pickOrDefault(rho, MAP.Depth2, rho.Properties.VariableNames{2});

% Salinity: prefer Salinity_used if exists; else raw salinity; else guess
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

% Geometry guessed if not specified
iceTh = pickOrDefault(rho, MAP.IceThickness, guessVar(rho, ["thick","thickness"], ""));
iceDr = pickOrDefault(rho, MAP.IceDraft,     guessVar(rho, ["draft"], ""));
coLen = pickOrDefault(rho, MAP.CoreLength,   guessVar(rho, ["length","corelength"], ""));

% Comments guessed if not specified
com = MAP.Comments;
if com == ""
    com = guessVar(rho, ["comment","remarks","notes"], "");
end

% Use processed lab density in kg/m3 ALWAYS
rhoLabVar = "rho_lab_kgm3";

preferredRHO = [
    string(MAP.EventLabel)
    "Time_best"
    string(MAP.GPS_Lat)
    string(MAP.GPS_Lon)
    iceTh
    iceDr
    coLen
    string(MAP.Station)
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

% ---------- T OUT ----------
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
    string(MAP.Station)
    string(MAP.IceAge)
    "MeltPond01"
    tDepth
    tTemp
    comT
];

T_all_proc.T_out = keepAndOrder(TT, preferredT);

% ---------- SALO18 OUT ----------
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
        string(MAP.Station)
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

% 6) Rename headers (PANGAEA convention) - applied to export tables only
renameMap = {
    "EventLabel",          "Event label"
    "Time_best",           "DATE/TIME"
    "GPS_Lat",             "LATITUDE"
    "GPS_Lon",             "LONGITUDE"
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
    "Station",             "Ice station"
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

% ---- Round calculated export variables to 2 decimals ----
% Helper inline: round only if numeric
roundVar = @(tbl,varname) setfield(tbl, varname, round(tbl.(varname),2));

% --- RHO export table ---
varsToRound_RHO = [
    "Density, ice, technical"
    "Density, ice"
    "Volume, brine"
    "Volume, gas, technical"
    "Volume, gas"
    "Temperature, ice/snow"
];

for k = 1:numel(varsToRound_RHO)
    vn = char(varsToRound_RHO(k)); % <-- important
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
clearvars -except T_all_proc T_all Tmatch

% ---- Export 3 final tables to one Excel workbook ----
outXlsx = "Coring_data_export.xlsx";

% RHO
writetable(T_all_proc.rho_out, outXlsx, "Sheet", "RHO", "WriteMode", "overwritesheet");

% T
writetable(T_all_proc.T_out,   outXlsx, "Sheet", "T",   "WriteMode", "overwritesheet");

% SALO18 
writetable(T_all_proc.SALO18_out, outXlsx, "Sheet", "SALO18", "WriteMode", "overwritesheet");

fprintf("Exported to %s (sheets: RHO, T, SALO18)\n", outXlsx);

%% Vertical profiles of salinity, density and temperature
% Panels: Salinity / Density / Temperature vs depth
% Color = date
% Marker = station number (1,2,3)

clear; close all; clc
load('Coring_data_processed.mat','T_all_proc');

rho = T_all_proc.rho_out;

% ---- Extract columns ----
t   = rho{:,2};
st_raw  = string(rho{:,8});
z1  = rho{:,11};
z2  = rho{:,12};
S   = rho{:,13};
Temp = rho{:,15};
Den  = rho{:,16};

zmid = mean([z1 z2], 2);

% ---- Station number from "IceStation1a" -> 1 ----
stationGroup = str2double(regexp(st_raw, '\d+', 'match', 'once'));

% ---- Clean rows ----
ok = ~isnan(zmid) & ~isnan(S) & ~isnan(Temp) & ~isnan(Den) & ...
     ~isnat(t) & ismember(stationGroup,[1 2 3]);

zmid = zmid(ok);
S    = S(ok);
Temp = Temp(ok);
Den  = Den(ok);
t    = t(ok);
stationGroup = stationGroup(ok);

% ---- Plot setup ----
figure('Color','w');
tiledlayout(1,3,'Padding','compact','TileSpacing','compact');

markers = containers.Map([1 2 3], {'o','s','^'});
c = datenum(t);
load("managua.mat"); colormap(flipud(managua));

% ===== Panel 1: Salinity =====
nexttile; hold on; box on;
for st = [1 2 3]
    idx = stationGroup == st;
    if any(idx)
        scatter(S(idx), zmid(idx), 25, c(idx), markers(st), 'filled');
    end
end
set(gca,'YDir','reverse');
xlabel('Salinity');
ylabel('Depth');
title('Salinity');

% ===== Panel 2: Density =====
nexttile; hold on; box on;
for st = [1 2 3]
    idx = stationGroup == st;
    if any(idx)
        scatter(Den(idx), zmid(idx), 25, c(idx), markers(st), 'filled');
    end
end
set(gca,'YDir','reverse');
xlabel('Density (kg/m^3)');
title('Density');

% ===== Panel 3: Temperature =====
nexttile; hold on; box on;
for st = [1 2 3]
    idx = stationGroup == st;
    if any(idx)
        scatter(Temp(idx), zmid(idx), 25, c(idx), markers(st), 'filled');
    end
end
set(gca,'YDir','reverse');
xlabel('Temperature (°C)');
title('Temperature');

% ---- Colorbar ----
cb = colorbar;
cb.Layout.Tile = 'east';
cb.Label.String = 'Date';
if ~isempty(c)
    clim([min(c) max(c)]);
    ticks = linspace(min(c), max(c), 6);
    cb.Ticks = ticks;
    cb.TickLabels = cellstr(datestr(ticks,'mmm dd'));
end

% ---- Marker legend ----
h1 = scatter(nan,nan,10,[0 0 0],'o','filled');
h2 = scatter(nan,nan,10,[0 0 0],'s','filled');
h3 = scatter(nan,nan,10,[0 0 0],'^','filled');
lgd = legend([h1 h2 h3], {'Station 1','Station 2','Station 3'}, ...
       'Location','northoutside','Orientation','horizontal');
lgd.Layout.Tile = 'north';

set(gcf,'Units','inches','Position',[5 5 6 4])
exportgraphics(gcf,'Vertical_profiles_SDT.png','Resolution',300)

%% 3-panel plot (RHO+SALO18 salinity, RHO density, T temperature)
clear; close all; clc
load('Coring_data_processed.mat','T_all_proc','Tmatch');   % contains T_all_proc.(rho,T,SALO18) + Tmatch

% ---------- 1) Standardize types ----------
T_all_proc.rho.SourceFile   = string(T_all_proc.rho.SourceFile);
T_all_proc.T.SourceFile     = string(T_all_proc.T.SourceFile);
T_all_proc.SALO18.SourceFile= string(T_all_proc.SALO18.SourceFile);

T_all_proc.rho.Station      = string(T_all_proc.rho.Station);
T_all_proc.T.Station        = string(T_all_proc.T.Station);
T_all_proc.SALO18.Station   = string(T_all_proc.SALO18.Station);

Tmatch.T_File   = string(Tmatch.T_File);
Tmatch.RHO_File = string(Tmatch.RHO_File);
Tmatch.Station  = string(Tmatch.Station);

% ---------- 2) Add CoreID (one integer per unique file, repeated for each row) ----------
[~,~,T_all_proc.rho.CoreID]    = unique(T_all_proc.rho.SourceFile,    'stable');
[~,~,T_all_proc.T.CoreID]      = unique(T_all_proc.T.SourceFile,      'stable');
[~,~,T_all_proc.SALO18.CoreID] = unique(T_all_proc.SALO18.SourceFile, 'stable');

% ---------- 3) Pick the columns ----------
rhoLabVar = "rho_lab_kgm3";     % preferred lab density variable (kg/m^3) in RHO table
rhoSalVar = "Salinity_used";    % preferred salinity variable in RHO table (can be "salinity")
tTempVar  = string(T_all_proc.T.Properties.VariableNames{2}); % temperature column in T table (or set explicitly)

% If rhoLabVar / rhoSalVar do not exist, fallback to sensible defaults
if ~ismember(rhoLabVar, T_all_proc.rho.Properties.VariableNames)
    rhoLabVar = string(T_all_proc.rho.Properties.VariableNames{3});  % old assumption
end
if ~ismember(rhoSalVar, T_all_proc.rho.Properties.VariableNames)
    if ismember("salinity", T_all_proc.rho.Properties.VariableNames)
        rhoSalVar = "salinity";
    else
        rhoSalVar = ""; % no RHO salinity available
    end
end

% ---------- 4) Make timezone-naive time columns for plotting (prevents tz errors) ----------
rhoTime = NaT(height(T_all_proc.rho),1);
if ismember("GPS_Time", T_all_proc.rho.Properties.VariableNames) && isa(T_all_proc.rho.GPS_Time,'datetime')
    rhoTime = T_all_proc.rho.GPS_Time;
    if ~isempty(rhoTime.TimeZone), rhoTime.TimeZone = ''; end
elseif ismember("Time_best", T_all_proc.rho.Properties.VariableNames) && isa(T_all_proc.rho.Time_best,'datetime')
    rhoTime = T_all_proc.rho.Time_best;
    if ~isempty(rhoTime.TimeZone), rhoTime.TimeZone = ''; end
end

saloTime = NaT(height(T_all_proc.SALO18),1);
if ismember("GPS_Time", T_all_proc.SALO18.Properties.VariableNames) && isa(T_all_proc.SALO18.GPS_Time,'datetime')
    saloTime = T_all_proc.SALO18.GPS_Time;
    if ~isempty(saloTime.TimeZone), saloTime.TimeZone = ''; end
else
    % fallback: first datetime column if exists
    isDT = varfun(@(x) isa(x,'datetime'), T_all_proc.SALO18, 'OutputFormat','uniform');
    dtVars = string(T_all_proc.SALO18.Properties.VariableNames(isDT));
    if ~isempty(dtVars)
        saloTime = T_all_proc.SALO18.(dtVars(1));
        if ~isempty(saloTime.TimeZone), saloTime.TimeZone = ''; end
    end
end

% ---------- 5) Precompute per-core aggregates (fast, clean) ----------
nR = max(T_all_proc.rho.CoreID);
nT = max(T_all_proc.T.CoreID);

rhoCoreTime      = splitapply(@firstOrNaT, rhoTime,                         T_all_proc.rho.CoreID);
rhoCoreDenLab    = splitapply(@(x) mean(x,'omitnan'), T_all_proc.rho.(rhoLabVar), T_all_proc.rho.CoreID);
rhoCoreSal       = NaN(nR,1);
if rhoSalVar ~= ""
    rhoCoreSal   = splitapply(@(x) mean(x,'omitnan'), T_all_proc.rho.(rhoSalVar), T_all_proc.rho.CoreID);
end

tCoreTemp         = splitapply(@(x) mean(x,'omitnan'), T_all_proc.T.(tTempVar),   T_all_proc.T.CoreID);

% density unit safety: if looks like g/cm^3, convert to kg/m^3
mDen = mean(rhoCoreDenLab,'omitnan');
if ~isnan(mDen) && mDen < 5
    rhoCoreDenLab = rhoCoreDenLab * 1000;
end

% ---------- 6) SALO18 salinity as one point per SALO18 core file ----------
% (mean salinity col 3 + time = first non-missing time in that file)
saloSal = splitapply(@(x) mean(x,'omitnan'), T_all_proc.SALO18{:,3}, T_all_proc.SALO18.CoreID);
saloCoreTime = splitapply(@firstOrNaT, saloTime, T_all_proc.SALO18.CoreID);

% station per SALO18 core (first row station)
saloStation = splitapply(@(x) string(x(1)), T_all_proc.SALO18.Station, T_all_proc.SALO18.CoreID);

saloSummary = table(saloCoreTime, saloStation, saloSal, ...
    'VariableNames', ["Time","Station","AvgSal_SALO18"]);
saloSummary = saloSummary(~ismissing(saloSummary.Time) & ~isnan(saloSummary.AvgSal_SALO18), :);
saloSummary = sortrows(saloSummary,"Time");

% ---------- 7) Build matched-core summary (one row per matched pair) ----------
rhoFileList = unique(T_all_proc.rho.SourceFile, 'stable');
tFileList   = unique(T_all_proc.T.SourceFile,   'stable');
rhoID = containers.Map(rhoFileList, 1:numel(rhoFileList));
tID   = containers.Map(tFileList,   1:numel(tFileList));

RHO_ID = arrayfun(@(s) rhoID(s), Tmatch.RHO_File);
T_ID   = arrayfun(@(s) tID(s),   Tmatch.T_File);

CoreSummary = table;
CoreSummary.Time         = rhoCoreTime(RHO_ID);
CoreSummary.Station      = Tmatch.Station;
CoreSummary.StationGroup = str2double(regexp(CoreSummary.Station,'\d+','match','once'));
CoreSummary.AvgSal_RHO   = rhoCoreSal(RHO_ID);
CoreSummary.AvgDenLab    = rhoCoreDenLab(RHO_ID);
CoreSummary.AvgTemp      = tCoreTemp(T_ID);

CoreSummary = CoreSummary(~ismissing(CoreSummary.Time), :);
CoreSummary = sortrows(CoreSummary,"Time");

% Optional: keep only station groups 1/2/3
CoreSummary = CoreSummary(ismember(CoreSummary.StationGroup,[1 2 3]) & ~isnan(CoreSummary.StationGroup), :);
saloSummary.StationGroup = str2double(regexp(saloSummary.Station,'\d+','match','once'));
saloSummary = saloSummary(ismember(saloSummary.StationGroup,[1 2 3]) & ~isnan(saloSummary.StationGroup), :);

% ---------- 8) Plot ----------
stations = unique(CoreSummary.StationGroup,'stable');
colors   = lines(max(3,numel(stations)));

figure('Color','w');
tiledlayout(3,1,'Padding','compact','TileSpacing','compact');

% Panel 1: Salinity from BOTH datasets (RHO circles, SALO18 triangles)
nexttile; hold on; box on; grid on;
for s = 1:numel(stations)
    st = stations(s);
    c  = colors(min(s,size(colors,1)),:);

    idxR = CoreSummary.StationGroup==st & ~isnan(CoreSummary.AvgSal_RHO);
    scatter(CoreSummary.Time(idxR), CoreSummary.AvgSal_RHO(idxR), 60, c, 'o', 'filled');

    idxS = saloSummary.StationGroup==st;
    scatter(saloSummary.Time(idxS), saloSummary.AvgSal_SALO18(idxS), 60, c, '^');
end
ylabel('Ice salinity');
title('Salinity vs time');
% legend (simple)
hSt = gobjects(0); labSt = strings(0);
for s = 1:numel(stations)
    c = colors(min(s,size(colors,1)),:);
    hSt(end+1) = scatter(nan,nan,60,c,'s','filled');
    labSt(end+1)= "Station " + string(stations(s));
end
hR = scatter(nan,nan,60,[0 0 0],'o','filled');
hS = scatter(nan,nan,60,[0 0 0],'^');
legend([hSt hR hS],[labSt "RHO" "SALO18"],'Location','northoutside','NumColumns',5);

% Panel 2: Lab density
nexttile; hold on; box on; grid on;
for s = 1:numel(stations)
    st = stations(s);
    c  = colors(min(s,size(colors,1)),:);
    idx = CoreSummary.StationGroup==st & ~isnan(CoreSummary.AvgDenLab);
    scatter(CoreSummary.Time(idx), CoreSummary.AvgDenLab(idx), 60, c, 'o', 'filled');
end
ylabel('Ice density (kg/m^3)');
title('Lab density vs time');

% Panel 3: Temperature
nexttile; hold on; box on; grid on;
for s = 1:numel(stations)
    st = stations(s);
    c  = colors(min(s,size(colors,1)),:);
    idx = CoreSummary.StationGroup==st & ~isnan(CoreSummary.AvgTemp);
    scatter(CoreSummary.Time(idx), CoreSummary.AvgTemp(idx), 60, c, 'o', 'filled');
end
ylabel('Ice temperature (°C)');
title('Temperature vs time');

set(gcf,'Units','inches','Position',[5 5 6 6])
exportgraphics(gcf,'Density_salinity_temperature_vs_time.png','Resolution',300)

% ---------- local helper (tiny, to keep the rest clean) ----------
function t = firstOrNaT(x)
    k = find(~ismissing(x), 1, 'first');
    if isempty(k), t = NaT; else, t = x(k); end
end

%% Plot avg lab density + avg in-situ density vs ice thickness (station = color, lab/in-situ = marker)
% Uses ONLY saved raw tables: T_all_proc + Tmatch from Coring_data_processed.mat
% Station grouping: "1a/1b/1c" -> station group 1 (same color)
% Different markers:
%   Lab density    = 'o' (filled circle)
%   In-situ density= '^' (filled triangle)

clear; close all; clc
load('Coring_data_processed.mat','T_all_proc','Tmatch')

Tmatch.RHO_File = string(Tmatch.RHO_File);
if ~ismember('Station', Tmatch.Properties.VariableNames)
    error("Tmatch has no 'Station' column.");
end
Tmatch.Station = string(Tmatch.Station);

T_all_proc.rho.SourceFile = string(T_all_proc.rho.SourceFile);

nPairs = height(Tmatch);

avgLabDensity    = nan(nPairs,1);
avgInSituDensity = nan(nPairs,1);
iceThickness     = nan(nPairs,1);
stationGroup     = nan(nPairs,1);

for i = 1:nPairs
    idx = (T_all_proc.rho.SourceFile == Tmatch.RHO_File(i));
    RHOcore = T_all_proc.rho(idx, :);
    if isempty(RHOcore), continue, end

    % Station group from Tmatch.Station (works for "1a", "Station 2b", "PS122-3c", etc.)
    token = regexp(Tmatch.Station(i), '\d+', 'match', 'once');
    if ~isempty(token)
        stationGroup(i) = str2double(token);
    end

    % Lab density
    if ismember('rho_lab_kgm3', RHOcore.Properties.VariableNames)
        avgLabDensity(i) = mean(RHOcore.rho_lab_kgm3, 'omitnan');
    else
        v = RHOcore{:,3};
        m = mean(v,'omitnan');
        if ~isnan(m) && m < 5, v = v * 1000; end % g/cm3 -> kg/m3 if needed
        avgLabDensity(i) = mean(v,'omitnan');
    end

    % In-situ density
    if ismember('rho_si', RHOcore.Properties.VariableNames)
        avgInSituDensity(i) = mean(RHOcore.rho_si, 'omitnan');
    end

    % Ice thickness
    if width(RHOcore) >= 8 && isnumeric(RHOcore{:,8})
        iceThickness(i) = mean(RHOcore{:,8}, 'omitnan');
    else
        vn = string(RHOcore.Properties.VariableNames);
        cand = vn(contains(lower(vn),'thick'));
        if ~isempty(cand)
            iceThickness(i) = mean(RHOcore.(cand(1)), 'omitnan');
        end
    end
end

% Keep only usable rows
ok = ~isnan(iceThickness) & ~isnan(stationGroup) & (~isnan(avgLabDensity) | ~isnan(avgInSituDensity));
avgLabDensity    = avgLabDensity(ok);
avgInSituDensity = avgInSituDensity(ok);
iceThickness     = iceThickness(ok);
stationGroup     = stationGroup(ok);

if isempty(iceThickness)
    error("Nothing to plot after filtering. Check thickness column and station labels.");
end

% Colors by station group
stations = unique(stationGroup,'stable');
colors = lines(max(3, numel(stations)));
colorByStation = containers.Map('KeyType','double','ValueType','any');
for s = 1:numel(stations)
    colorByStation(stations(s)) = colors(s,:);
end

% Marker styles for the two density types
mkLab    = 'o';
mkInSitu = '^';
sz = 80;

figure('Color','w'); hold on; box on; grid on;

% Plot per station group (same color), different markers for lab vs in-situ
for s = 1:numel(stations)
    st = stations(s);
    c  = colorByStation(st);
    idxS = (stationGroup == st);

    % Lab density points
    idxLab = idxS & ~isnan(avgLabDensity);
    if any(idxLab)
        scatter(iceThickness(idxLab), avgLabDensity(idxLab), sz, c, mkLab);
    end

    % In-situ density points
    idxIS = idxS & ~isnan(avgInSituDensity);
    if any(idxIS)
        scatter(iceThickness(idxIS), avgInSituDensity(idxIS), sz, c, mkInSitu, 'filled');
    end
end

xlabel('Ice thickness (m)');
ylabel('Average sea-ice density (kg/m^3)');
title('Average sea-ice density vs ice thickness');

% Legend: station colors + marker meaning (lab vs in-situ)
legendHandles = gobjects(0);
legendLabels  = strings(0);

% Station color legend (dummy)
for s = 1:numel(stations)
    st = stations(s);
    c  = colorByStation(st);
    h = scatter(nan,nan,sz,c,'s','filled');  % square dummy
    legendHandles(end+1) = h;
    legendLabels(end+1)  = "Station " + string(st);
end

% Marker meaning (dummy black)
h1 = scatter(nan,nan,sz,[0 0 0],mkLab);
h2 = scatter(nan,nan,sz,[0 0 0],mkInSitu,'filled');
legendHandles(end+1:end+2) = [h1 h2];
legendLabels(end+1:end+2)  = ["Lab density", "In-situ density"];

legend(legendHandles, legendLabels, 'Location','southeast');

set(gcf,'Units','inches','Position',[5 5 6 4])
exportgraphics(gcf,'Density_vs_Thickness.png','Resolution',300)

%% Collect unique names + count occurrences
clear; close all; clc

rootFolder = "C:\Users\evsalg001\Documents\MATLAB\Contrasts coring\Data";

files = [ ...
    dir(fullfile(rootFolder, "**", "*.xlsx")); ...
    dir(fullfile(rootFolder, "**", "*.xlsm"))  ...
];

if isempty(files)
    error("No Excel files found under: %s", rootFolder);
end

sheetName = "metadata-coring";
rangeName = "C32:L32";

allNames = strings(0,1);

for k = 1:numel(files)

    fname = files(k).name;

    if startsWith(fname, "~$")
        continue
    end

    filePath = fullfile(files(k).folder, fname);

    try
        vals = readcell(filePath, "Sheet", sheetName, "Range", rangeName);
    catch
        continue   % skip if sheet doesn't exist or read fails
    end

    vals = vals(:);

    for i = 1:numel(vals)
        v = vals{i};

        if isempty(v)
            continue
        elseif ismissing(v)
            continue
        elseif isnumeric(v) && isscalar(v) && isnan(v)
            continue
        else
            name = strtrim(string(v));
            if name ~= "" && lower(name) ~= "nan"
                allNames(end+1,1) = name;
            end
        end
    end
end

% Remove blanks
allNames = allNames(allNames ~= "");

% ---- Count occurrences (case-insensitive) ----
namesLower = lower(allNames);
[uniqueLower, ~, idx] = unique(namesLower);
counts = accumarray(idx, 1);

% Keep original capitalization from first occurrence
uniqueNames = strings(size(uniqueLower));
for i = 1:numel(uniqueLower)
    firstMatch = find(namesLower == uniqueLower(i), 1, 'first');
    uniqueNames(i) = allNames(firstMatch);
end

% Sort alphabetically
[uniqueNames, sortIdx] = sort(uniqueNames);
counts = counts(sortIdx);

% Create result table
result = table(uniqueNames, counts, ...
    'VariableNames', ["Name","Count"]);

fprintf("\nUnique names found: %d\n\n", height(result));
disp(result);

%% ============================ HELPER FUNCTIONS ============================

function tbl = keepAndOrder(tbl, preferredVars)
    preferredVars = preferredVars(preferredVars ~= "");
    keep = preferredVars(ismember(preferredVars, tbl.Properties.VariableNames));
    tbl = tbl(:, keep);
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

function tbl = applyRenameMap(tbl, renameMap)
    for i = 1:size(renameMap,1)
        oldName = renameMap{i,1};
        newName = renameMap{i,2};
        if ismember(oldName, tbl.Properties.VariableNames)
            tbl = renamevars(tbl, oldName, newName);
        end
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
