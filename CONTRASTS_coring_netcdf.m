%% ===== NetCDF export (PROCESS SCRIPT ADD-ON) =====
% Exports Coring tables to NetCDF (with units + comments + categorical attrs)
% - One NetCDF4 file with groups: /RHO, /T, /SALO18
% - Variable names are sanitized versions of table headers
% - Adds units/standard_name/long_name/comment + category/flag metadata

ncFile = fullfile(exportFolder, "Contrasts_physical_properties_coring.nc");

% Create NetCDF file (netcdf4)
if isfile(ncFile), delete(ncFile); end
nccreate(ncFile, "dummy", "Dimensions", {"dummy", 1}, "Format", "netcdf4");
ncwrite(ncFile, "dummy", 1);

% -------- Global attributes (edit emails if needed) --------
ncwriteatt(ncFile,"/","title","Sea ice physical properties from the Contrasts expedition");
ncwriteatt(ncFile,"/","Conventions","CF-1.7");
ncwriteatt(ncFile,"/","contributor_name","Dmitry Divine, Evgenii Salganik, David Clemens-Sewall, Emiliano Cimoli, Lena Eggers, Keigo Takahashi, Marcel Nicolaus");
ncwriteatt(ncFile,"/","contributor_email","evgenii.salganik@awi.de");
ncwriteatt(ncFile,"/","institution","Alfred Wegener Institute for Polar and Marine Research");
ncwriteatt(ncFile,"/","creator_name","Evgenii Salganik");
ncwriteatt(ncFile,"/","creator_email","evgenii.salganik@awi.de");
ncwriteatt(ncFile,"/","project","Arctic PASSION");
ncwriteatt(ncFile,"/","summary","First- and second-year sea-ice salinity, temperature, and density from the coring sites during the Contrasts expedition in July-August 2025");
ncwriteatt(ncFile,"/","license","CC-0");

% Time coverage (from DATE/TIME if present)
[tStartISO, tEndISO] = inferTimeCoverageISO(T_all_proc);
if tStartISO ~= "", ncwriteatt(ncFile,"/","time_coverage_start", tStartISO); end
if tEndISO   ~= "", ncwriteatt(ncFile,"/","time_coverage_end",   tEndISO); end

% Geospatial coverage (from LAT/LON if present)
[latMin, latMax, lonMin, lonMax] = inferGeoCoverage(T_all_proc);
if ~isnan(latMin), ncwriteatt(ncFile,"/","geospatial_lat_min", num2str(latMin, '%.6f')); end
if ~isnan(latMax), ncwriteatt(ncFile,"/","geospatial_lat_max", num2str(latMax, '%.6f')); end
if ~isnan(lonMin), ncwriteatt(ncFile,"/","geospatial_lon_min", num2str(lonMin, '%.6f')); end
if ~isnan(lonMax), ncwriteatt(ncFile,"/","geospatial_lon_max", num2str(lonMax, '%.6f')); end

ncwriteatt(ncFile,"/","keywords","arctic, polar, sea ice, salinity, temperature, density, coring");
ncwriteatt(ncFile,"/","calendar","standard");
ncwriteatt(ncFile,"/","date_created",char(datetime("now","TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss'Z'")));
ncwriteatt(ncFile,"/","featureType","timeseries");
ncwriteatt(ncFile,"/","product_version","1");

% -------- Metadata maps (units + comments) --------
metaRHO  = buildMetaMap_RHO();
metaT    = buildMetaMap_T();
metaSALO = buildMetaMap_SALO18();

% -------- Write groups --------
writeTableGroupNetCDF(ncFile, "/RHO",    T_all_proc.rho_out,    metaRHO);
writeTableGroupNetCDF(ncFile, "/T",      T_all_proc.T_out,      metaT);

if isfield(T_all_proc,'SALO18_out') && ~isempty(T_all_proc.SALO18_out)
    writeTableGroupNetCDF(ncFile, "/SALO18", T_all_proc.SALO18_out, metaSALO);
else
    warning("No SALO18_out table found — skipping /SALO18 group.");
end

fprintf("Exported NetCDF: %s (groups: /RHO, /T, /SALO18)\n", ncFile);

%% ========================= FUNCTIONS (NetCDF export) =========================

function writeTableGroupNetCDF(ncFile, grp, T, metaMap)

if isempty(T) || height(T)==0
    warning("Group %s: table empty, skipping.", grp);
    return
end

nObs   = height(T);
obsDim = "obs";
vars   = string(T.Properties.VariableNames);

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

        % create group-local strlen dim placeholder
        nccreate(ncFile, grp + "/" + strlenDim, "Dimensions", {strlenDim, maxLen}, "Datatype", "int32");

        nccreate(ncFile, vpath, "Dimensions", {strlenDim, maxLen, obsDim, nObs}, "Datatype", "char");

        C = char(pad(s, maxLen));  % nObs x maxLen
        C = permute(C, [2 1]);     % maxLen x nObs
        ncwrite(ncFile, vpath, C);
    end

    % Default long_name = table header
    ncwriteatt(ncFile, vpath, "long_name", char(origName));

    % ---- Apply metadata (units/standard_name/comment) from map ----
    key1 = char(origName);   % table header
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

function tf = isTimeColumn(varName)
vn = string(varName);
tf = any(vn == ["DATE/TIME","Time_best","time","TIME","DateTime","datetime"]);
end

function tDays = convertTimeToDaysSince1979(x)
ref = datetime(1979,1,1,0,0,0,"TimeZone","UTC");

if isdatetime(x)
    t = x;
else
    s = string(x);
    s(ismissing(s)) = "";
    try
        t = datetime(s, "TimeZone","UTC", "InputFormat","yyyy-MM-dd HH:mm:ss");
    catch
        t = datetime(s, "TimeZone","UTC");
    end
end

t = datetime(t, "TimeZone","UTC");
tDays = double(days(t - ref));
end

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

% -------- Metadata map for RHO table (units + comments)
function M = buildMetaMap_RHO()
M = containers.Map();

M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east");

M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");

% If your output headers are exactly these after rename:
M("Core length") = struct("standard_name","core_length","units","m");
M("Depth, ice/snow") = struct("standard_name","depth","units","m");
M("Depth, ice/snow, top/minimum")    = struct("standard_name","depth","units","m");
M("Depth, ice/snow, bottom/maximum") = struct("standard_name","depth","units","m");

M("Sea ice salinity") = struct("standard_name","sea_ice_salinity","units","1e-3","comment","Practical salinity (PSU).");
M("Temperature, technical") = struct("standard_name","temperature","units","degree_Celsius","comment","Air temperature in laboratory.");
M("Temperature, ice/snow") = struct("standard_name","sea_ice_temperature","units","degree_Celsius","comment","In situ ice temperature.");

M("Density, ice, technical") = struct("standard_name","sea_ice_density","units","kg/m3","comment","Sea ice density measured in laboratory.");
M("Density, ice") = struct("standard_name","sea_ice_density","units","kg/m3","comment","Sea ice density estimated for in situ temperature.");

M("Volume, brine") = struct("standard_name","brine_volume_fraction","units","1","comment","Brine volume fraction estimated for in situ temperature.");
M("Volume, gas, technical") = struct("standard_name","gas_volume_fraction","units","1","comment","Gas volume fraction estimated for laboratory temperature.");
M("Volume, gas") = struct("standard_name","gas_volume_fraction","units","1","comment","Gas volume fraction estimated for in situ temperature.");
end

% -------- Metadata map for T table
function M = buildMetaMap_T()
M = containers.Map();
M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east");

M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");

M("Core length") = struct("standard_name","core_length","units","m");
M("Depth, ice/snow") = struct("standard_name","depth","units","m");
M("Temperature, ice/snow") = struct("standard_name","sea_ice_temperature","units","degree_Celsius","comment","In situ ice temperature.");
end

% -------- Metadata map for SALO18 table
function M = buildMetaMap_SALO18()
M = containers.Map();
M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east");

M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");

M("Core length") = struct("standard_name","core_length","units","m");
M("Depth, ice/snow, top/minimum")    = struct("standard_name","depth","units","m");
M("Depth, ice/snow, bottom/maximum") = struct("standard_name","depth","units","m");
M("Sea ice salinity") = struct("standard_name","sea_ice_salinity","units","1e-3","comment","Practical salinity (PSU).");
end

% -------- Helpers to infer time/geo coverage from output tables
function [tStartISO, tEndISO] = inferTimeCoverageISO(T_all_proc)
tStartISO = ""; tEndISO = "";

cand = {};
if isfield(T_all_proc,"rho_out"), cand{end+1} = T_all_proc.rho_out; end
if isfield(T_all_proc,"T_out"),   cand{end+1} = T_all_proc.T_out;   end
if isfield(T_all_proc,"SALO18_out"), cand{end+1} = T_all_proc.SALO18_out; end

tAll = NaT(0,1);
for k = 1:numel(cand)
    T = cand{k};
    if isempty(T), continue, end
    if ismember("DATE/TIME", T.Properties.VariableNames) && isdatetime(T.("DATE/TIME"))
        tAll = [tAll; T.("DATE/TIME")];
    end
end

tAll = tAll(~ismissing(tAll));
if isempty(tAll), return, end

tmin = min(tAll);
tmax = max(tAll);

if isempty(tmin.TimeZone), tmin.TimeZone = "UTC"; end
if isempty(tmax.TimeZone), tmax.TimeZone = "UTC"; end

tStartISO = char(datetime(tmin,"TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss"));
tEndISO   = char(datetime(tmax,"TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss"));
end

function [latMin, latMax, lonMin, lonMax] = inferGeoCoverage(T_all_proc)
latMin = NaN; latMax = NaN; lonMin = NaN; lonMax = NaN;

cand = {};
if isfield(T_all_proc,"rho_out"), cand{end+1} = T_all_proc.rho_out; end
if isfield(T_all_proc,"T_out"),   cand{end+1} = T_all_proc.T_out;   end
if isfield(T_all_proc,"SALO18_out"), cand{end+1} = T_all_proc.SALO18_out; end

latAll = [];
lonAll = [];
for k = 1:numel(cand)
    T = cand{k};
    if isempty(T), continue, end
    if ismember("LATITUDE", T.Properties.VariableNames)
        latAll = [latAll; double(T.LATITUDE)];
    end
    if ismember("LONGITUDE", T.Properties.VariableNames)
        lonAll = [lonAll; double(T.LONGITUDE)];
    end
end

latAll = latAll(~isnan(latAll));
lonAll = lonAll(~isnan(lonAll));

if ~isempty(latAll)
    latMin = min(latAll); latMax = max(latAll);
end
if ~isempty(lonAll)
    lonMin = min(lonAll); lonMax = max(lonAll);
end
end