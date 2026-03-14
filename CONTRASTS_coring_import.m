%% IMPORT ONLY: RHO, T, SALO18 (with GPS + metadata)
clear; close all; clc
tStart = tic;

projectRoot = fileparts(which("import_all_cores_only.m"));
if isempty(projectRoot); projectRoot = pwd; end
rootFolder = fullfile(projectRoot, "Data");

filePatterns = ["*-RHO.xlsx","*-RHO-TXT.xlsx","*-SALO18.xlsx","*-T.xlsx"];
allFiles = [];

for fp = filePatterns
    files = dir(fullfile(rootFolder,"**",fp));
    allFiles = [allFiles; files];
end

T_all = struct('rho', table(), 'T', table(), 'SALO18', table());

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

        try
% --- FAST fixed-range read (no detectImportOptions) ---
% Read enough columns to cover: depth1, depth2, density (~col7), comments (~col8)
Traw = readtable(filePath, ...
    "Sheet","Density-densimetry", ...
    "Range","A1:H150", ...   % header + up to ~149 rows (safe for 100 measurements)
    "ReadVariableNames", true, ...
    "VariableNamingRule","preserve", 'UseExcel', false);

nCols = width(Traw);
cDepth1 = 1;
cDepth2 = 2;
cDens   = min(7, nCols); if cDens < 3, cDens = 3; end
cCom    = min(8, nCols); if cCom < 3, cCom = nCols; end

T = Traw(:, [cDepth1 cDepth2 cDens cCom]);
T.Properties.VariableNames = {'Depth1','Depth2','Density_gcm3','Comments'};
        catch ME
            warning("Skipping RHO file (readtable failed): %s\n  %s", filePath, ME.message);
            continue
        end

        % Convert numeric columns (skip comment col 4)
        T = forceNumericTableExcept(T, 4);

        % Parse comments (parafine and lab)
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

        % -------- Import SALO18 values from SAME file (fixed E5:G) --------
        n = height(T);

        % One read instead of 3 (faster + no shifting)
        block = readcell(filePath, "Sheet","SALO18", "Range", sprintf("E5:G%d", 4+n), 'UseExcel', false);
        salE  = cellfun(@toScalarNum, block(:,1));
        condF = cellfun(@toScalarNum, block(:,2));
        tempG = cellfun(@toScalarNum, block(:,3));

        T.salinity      = salE;
        T.cond_mScm     = condF;
        T.tempC_SALO18  = tempG;

        % Add metadata (FAST: read metadata-core once per file)
        newTbl = add_metadata_fast(T,filePath,{'A7','A8','A10','A2'},{'C7','C8','C10','C2'},extraMeta);
        mc = readMetaCoringOnce(filePath);
        newTbl = add_metadata_coring_fields(newTbl, mc);

        % GPS from folder (unchanged)
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));

        % Fallback GPS from metadata-coring ONLY if folder GPS missing
missingGPS = all(isnan(newTbl.GPS_Lat)) || all(isnan(newTbl.GPS_Lon)) || all(isnat(newTbl.GPS_Time));

if missingGPS && mc.hasAny && ~isnan(mc.lat) && ~isnan(mc.lon) && ~isnat(mc.date)
    newTbl.GPS_Lat  = repmat(mc.lat, height(newTbl), 1);
    newTbl.GPS_Lon  = repmat(mc.lon, height(newTbl), 1);
    newTbl.GPS_Time = repmat(mc.date + mc.tod, height(newTbl), 1);
end

        % If still no GPS_Time, build it from metadata-core date + metadata-coring time (default 09:00)
        newTbl = fillGPSTimeFromMetaCoreDate(newTbl, mc.tod);

        % Ensure timezone consistency for concatenation
        if ismember("GPS_Time", newTbl.Properties.VariableNames) && isdatetime(newTbl.GPS_Time)
            newTbl.GPS_Time.TimeZone = "UTC";
        end

        if ~isempty(newTbl)
            rhoCoreCounter = rhoCoreCounter + 1;
            newTbl.CoreID_RHO = repmat(rhoCoreCounter, height(newTbl), 1);
            T_all.rho = [T_all.rho; newTbl];
        end

    % ---------------- T cores ----------------
    elseif contains(filePath,'-T.xlsx')

        try
            T = readtable(filePath,'Sheet','TEMP','Range','A2:B1000', ...
                'ReadVariableNames',true,'VariableNamingRule','preserve', 'UseExcel', false);
        catch ME
            warning("Skipping T file (cannot read TEMP): %s\n  %s", filePath, ME.message);
            continue
        end

        T = forceNumericTable(T);
        T = T(~isnan(T{:,1}), :);

        newTbl = add_metadata_fast(T,filePath,{'A7','A8','A10'},{'C7','C8','C10'},extraMeta);
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));
        mc = readMetaCoringOnce(filePath);
        newTbl = add_metadata_coring_fields(newTbl, mc);
        newTbl = fillGPSTimeFromMetaCoreDate(newTbl, mc.tod);

        if ismember("GPS_Time", newTbl.Properties.VariableNames) && isdatetime(newTbl.GPS_Time)
            newTbl.GPS_Time.TimeZone = "UTC";
        end

        if ~isempty(newTbl)
            tCoreCounter = tCoreCounter + 1;
            newTbl.CoreID_T = repmat(tCoreCounter, height(newTbl), 1);
            T_all.T = [T_all.T; newTbl];
        end

    % ---------------- SALO18 cores ----------------
elseif contains(filePath,'-SALO18.xlsx')

    % ---- FAST fixed-range read (no detectImportOptions) ----
    Traw = readtable(filePath, ...
        "Sheet","SALO18", ...
        "Range","A1:G120", ...      % safe for <=100 rows
        "ReadVariableNames",true, ...
        "VariableNamingRule","preserve", 'UseExcel', false);

    % Remove empty depth rows
    Traw = Traw(~isnan(Traw{:,1}), :);

    % Build clean table with expected columns
    T = table;

    T.Depth1 = Traw{:,1};
    T.Depth2 = Traw{:,2};

    if width(Traw) >= 5
        T.Salinity = Traw{:,5};
    else
        T.Salinity = NaN(height(Traw),1);
    end

    if width(Traw) >= 6
        T.Conductivity_mScm = Traw{:,6};
    else
        T.Conductivity_mScm = NaN(height(Traw),1);
    end

    if width(Traw) >= 7
        T.TempC = Traw{:,7};
    else
        T.TempC = NaN(height(Traw),1);
    end

    T = forceNumericTable(T);

        newTbl = add_metadata_fast(T,filePath,{'A7','A8','A10'},{'C7','C8','C10'},extraMeta);
        [newTbl.GPS_Lat,newTbl.GPS_Lon,newTbl.GPS_Time] = getFirstGPS(filePath,height(newTbl));
        newTbl = fillGPSTimeFromMetaCoreDate(newTbl, filePath);
        mc = readMetaCoringOnce(filePath);
        newTbl = add_metadata_coring_fields(newTbl, mc);
newTbl = fillGPSTimeFromMetaCoreDate(newTbl, mc.tod);

        if ismember("GPS_Time", newTbl.Properties.VariableNames) && isdatetime(newTbl.GPS_Time)
            newTbl.GPS_Time.TimeZone = "UTC";
        end

        if ~isempty(newTbl)
            saloCoreCounter = saloCoreCounter + 1;
            newTbl.CoreID_SALO18 = repmat(saloCoreCounter, height(newTbl), 1);
            T_all.SALO18 = [T_all.SALO18; newTbl];
        end
    end
end

elapsed = toc(tStart);
fprintf('IMPORT completed in %.0f seconds.\n', elapsed);
fprintf("Imported cores: RHO=%d, T=%d, SALO18=%d\n", rhoCoreCounter, tCoreCounter, saloCoreCounter);

save('Coring_data_imported.mat','T_all');
fprintf('Imported data saved to Coring_data_imported.mat\n');

%% ========================= FUNCTIONS =========================

function T = forceNumericTable(T)
vars = T.Properties.VariableNames;
for i = 1:numel(vars)
    v = T.(vars{i});
    if iscell(v) || isstring(v)
        T.(vars{i}) = str2double(string(v));
    end
end
end

function T = forceNumericTableExcept(T, skipColIdx)
vars = T.Properties.VariableNames;
for i = 1:numel(vars)
    if i == skipColIdx, continue; end
    v = T.(vars{i});
    if iscell(v) || isstring(v)
        T.(vars{i}) = str2double(string(v));
    end
end
end

function x = toScalarNum(v)
if isnumeric(v) && isscalar(v)
    x = double(v);
else
    x = str2double(string(v));
end
end

function tbl = add_metadata_fast(tbl,file,name_cells,value_cells,extraMeta)

if isempty(tbl), return; end

% Read one block; covers A2,C2,A7,C7,A8,C8,A10,C10, etc.
meta = readcell(file,'Sheet','metadata-core','Range','A1:C20', 'UseExcel', false);

for j = 1:numel(name_cells)
    nm = meta{cellRefRow(name_cells{j}), cellRefCol(name_cells{j})};
    varname = matlab.lang.makeValidName(string(nm));

    val = meta{cellRefRow(value_cells{j}), cellRefCol(value_cells{j})};

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

fn = fieldnames(extraMeta);
for i = 1:numel(fn)
    tbl.(fn{i}) = repmat(string(extraMeta.(fn{i})),height(tbl),1);
end

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

[~,fname,ext] = fileparts(file);
tbl.SourceFile = repmat(string(strcat(fname, ext)), height(tbl),1);

iceAge = repmat("Unknown", height(tbl),1);
if contains(fname,'FYI','IgnoreCase',true), iceAge(:)="FYI";
elseif contains(fname,'SMYI','IgnoreCase',true), iceAge(:)="SMYI";
elseif contains(fname,'SYI','IgnoreCase',true), iceAge(:)="SYI";
end
tbl.IceAge = iceAge;

expr = "PS\d+.*?(?=-SI_corer)";
tokens = regexp(fname, expr, 'match', 'once');
if isempty(tokens)
    tbl.EventLabel = repmat("Unknown", height(tbl), 1);
else
    tbl.EventLabel = repmat(string(tokens), height(tbl), 1);
end

tbl.MeltPond = repmat(contains(fname,'melt_pond','IgnoreCase',true), height(tbl),1);
end

function r = cellRefRow(a1)
r = sscanf(regexprep(a1,'[A-Z]',''),'%d');
end

function c = cellRefCol(a1)
letters = regexprep(upper(a1),'[0-9]','');
c = 0;
for k=1:numel(letters)
    c = c*26 + (letters(k)-'A'+1);
end
end

function tbl = fillGPSTimeFromMetaCoreDate(tbl, tod)
% If GPS_Time is all-NaT, set it to (date from metadata-core columns in tbl) + tod.
% tod is a duration; default is 09:00 if empty/invalid.

    if isempty(tbl) || ~ismember("GPS_Time", tbl.Properties.VariableNames)
        return
    end
    if ~isdatetime(tbl.GPS_Time) || ~all(isnat(tbl.GPS_Time))
        return
    end

    vn = string(tbl.Properties.VariableNames);
    dateCandidates = vn(contains(lower(vn), "date"));
    if isempty(dateCandidates), return, end

    d = NaT(height(tbl),1);
    found = false;
    for i = 1:numel(dateCandidates)
        v = tbl.(dateCandidates(i));
        if isdatetime(v) && ~all(isnat(v))
            d = v; found = true; break
        end
    end
    if ~found, return, end

    d0 = dateshift(d, "start", "day");

    if nargin < 2 || ~isduration(tod) || isnan(seconds(tod))
        tod = hours(9);
    end

    tbl.GPS_Time = d0 + tod;
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

function mc = readMetaCoringOnce(filePath)
% Read metadata-coring block once (C8:C16)
% mc.lat, mc.lon, mc.date, mc.tod (duration), mc.snowHeight, mc.hasAny
    mc = struct('lat',NaN,'lon',NaN,'date',NaT,'tod',hours(9), ...
                'snowHeight',NaN,'hasAny',false);

    try
        block = readcell(filePath,'Sheet','metadata-coring','Range','C8:C16','UseExcel',false);
    catch
        return
    end

    % C8 lat, C9 lon, C12 date, C13 time, C16 snow height
    lat0  = str2double(strrep(strtrim(string(block{1})),",","."));
    lon0  = str2double(strrep(strtrim(string(block{2})),",","."));
    dRaw  = block{5};
    tRaw  = block{6};
    snow0 = str2double(strrep(strtrim(string(block{9})),",","."));

    if ~isnan(lat0),  mc.lat = lat0; end
    if ~isnan(lon0),  mc.lon = lon0; end
    if ~isnan(snow0), mc.snowHeight = snow0; end

    % date
    if isdatetime(dRaw)
        mc.date = dateshift(dRaw,'start','day');
    elseif isnumeric(dRaw) && isscalar(dRaw) && ~isnan(dRaw)
        mc.date = dateshift(datetime(dRaw,'ConvertFrom','excel'),'start','day');
    else
        ds = strtrim(string(dRaw));
        if strlength(ds)>0
            try
                mc.date = dateshift(datetime(ds,'InputFormat','yyyy-MM-dd'),'start','day');
            catch
                try mc.date = dateshift(datetime(ds),'start','day'); catch, end
            end
        end
    end

    % time-of-day (default 09:00)
    if isnumeric(tRaw) && isscalar(tRaw) && ~isnan(tRaw)
        mc.tod = seconds(tRaw*24*3600);
    else
        ts = strtrim(string(tRaw));
        if strlength(ts)>0
            try
                mc.tod = duration(ts,'InputFormat','hh:mm');
            catch
                try mc.tod = duration(ts,'InputFormat','hh:mm:ss'); catch, end
            end
        end
    end

    mc.hasAny = (~isnan(mc.lat) || ~isnan(mc.lon) || ~isnat(mc.date));
end

function tbl = add_metadata_coring_fields(tbl, mc)
if isempty(tbl), return; end
tbl.SnowHeight = repmat(mc.snowHeight, height(tbl), 1);
end