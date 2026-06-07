%% ===== NetCDF export =====
% Exports Coring tables to separate NetCDF files (with units + comments + categorical attrs)
% - One NetCDF4 file per table:
%     Contrasts_coring_density.nc
%     Contrasts_coring_temperature.nc
%     Contrasts_coring_salinity.nc
% - Variable names are sanitized versions of table headers
% - Adds units/standard_name/long_name/comment + category/flag metadata

clear; close all; clc

projectRoot = pwd;

inputFolder  = fullfile(projectRoot, "data", "final");
outputFolder = fullfile(projectRoot, "data", "final", "netcdf");

if ~isfolder(outputFolder)
    mkdir(outputFolder);
end

ncFileRHO  = fullfile(outputFolder, "Contrasts_coring_density.nc");
ncFileT    = fullfile(outputFolder, "Contrasts_coring_temperature.nc");
ncFileSALO = fullfile(outputFolder, "Contrasts_coring_salinity.nc");

INFILE_PROCESSED = fullfile(inputFolder, "core_data_processed.mat");

if ~isfile(INFILE_PROCESSED)
    error("Processed MAT not found: %s", INFILE_PROCESSED);
end

load(INFILE_PROCESSED, "T_all_proc");

doiRHO  = "https://doi.org/10.1594/PANGAEA.993687";
doiT    = "https://doi.org/10.1594/PANGAEA.993704";
doiSALO = "https://doi.org/10.1594/PANGAEA.993703";

% Remove old files
if isfile(ncFileRHO),  delete(ncFileRHO);  end
if isfile(ncFileT),    delete(ncFileT);    end
if isfile(ncFileSALO), delete(ncFileSALO); end

% -------- Metadata maps (units + comments) --------
metaRHO  = buildMetaMap_RHO();
metaT    = buildMetaMap_T();
metaSALO = buildMetaMap_SALO18();

% -------- Write separate files --------
writeSingleNetCDF( ...
    ncFileRHO, ...
    T_all_proc.rho_out, ...
    metaRHO, ...
    "Sea ice density from the Contrasts expedition", ...
    "First- and second-year sea-ice density from the coring sites during the Contrasts expedition in July-August 2025", ...
    doiRHO);

writeSingleNetCDF( ...
    ncFileT, ...
    T_all_proc.T_out, ...
    metaT, ...
    "Sea ice temperature from the Contrasts expedition", ...
    "First- and second-year sea-ice temperature from the coring sites during the Contrasts expedition in July-August 2025", ...
    doiT);

if isfield(T_all_proc,'SALO18_out') && ~isempty(T_all_proc.SALO18_out)
writeSingleNetCDF( ...
    ncFileSALO, ...
    T_all_proc.SALO18_out, ...
    metaSALO, ...
    "Sea ice salinity from the Contrasts expedition", ...
    "First- and second-year sea-ice salinity from the coring sites during the Contrasts expedition in July-August 2025", ...
    doiSALO);
else
    warning("No SALO18_out table found — skipping salinity export.");
end

fprintf("Exported NetCDF files:\n");
fprintf("  %s\n", ncFileRHO);
fprintf("  %s\n", ncFileT);
if isfile(ncFileSALO)
    fprintf("  %s\n", ncFileSALO);
end

%% NetCDF import check
clear; clc; close all;
outputFolder = fullfile(pwd, "data", "final", "netcdf");

filename = fullfile(exportFolder, 'Contrasts_coring_density.nc');
if isfile(filename), ncdisp(filename); end

filename = fullfile(exportFolder, 'Contrasts_coring_temperature.nc');
if isfile(filename), ncdisp(filename); end

filename = fullfile(exportFolder, 'Contrasts_coring_salinity.nc');
if isfile(filename), ncdisp(filename); end

%% ========================= FUNCTIONS (NetCDF export) =========================

function writeSingleNetCDF(ncFile, T, metaMap, titleText, summaryText, doi)

if isempty(T) || height(T) == 0
    warning("Table for file %s is empty, skipping.", ncFile);
    return
end

% First create the real variables. The first nccreate call will create the file.
writeTableRootNetCDF(ncFile, T, metaMap);

% -------- Global attributes --------
ncwriteatt(ncFile,"/","title",titleText);
ncwriteatt(ncFile,"/","Conventions","CF-1.7");
ncwriteatt(ncFile,"/","id",char(doi));
ncwriteatt(ncFile,"/","contributor_name","Dmitry Divine, Evgenii Salganik, David Clemens-Sewall, Emiliano Cimoli, Lena Eggers, Keigo Takahashi, Marcel Nicolaus");
ncwriteatt(ncFile,"/","contributor_email","evgenii.salganik@awi.de");
ncwriteatt(ncFile,"/","institution","Alfred Wegener Institute for Polar and Marine Research");
ncwriteatt(ncFile,"/","creator_name","Evgenii Salganik");
ncwriteatt(ncFile,"/","creator_email","evgenii.salganik@awi.de");
ncwriteatt(ncFile,"/","project","Arctic PASSION");
ncwriteatt(ncFile,"/","summary",summaryText);
ncwriteatt(ncFile,"/","license","CC-0");
ncwriteatt(ncFile,"/","keywords","arctic, polar, sea ice, salinity, temperature, density, coring");
ncwriteatt(ncFile,"/","calendar","standard");
ncwriteatt(ncFile,"/","date_created",char(datetime("now","TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss'Z'")));
ncwriteatt(ncFile,"/","featureType","timeseries");
ncwriteatt(ncFile,"/","product_version","1");

% Time coverage
[tStartISO, tEndISO] = inferTimeCoverageISO_single(T);
if tStartISO ~= "", ncwriteatt(ncFile,"/","time_coverage_start", tStartISO); end
if tEndISO   ~= "", ncwriteatt(ncFile,"/","time_coverage_end",   tEndISO); end

% Geospatial coverage
[latMin, latMax, lonMin, lonMax] = inferGeoCoverage_single(T);
if ~isnan(latMin), ncwriteatt(ncFile,"/","geospatial_lat_min", num2str(latMin, '%.6f')); end
if ~isnan(latMax), ncwriteatt(ncFile,"/","geospatial_lat_max", num2str(latMax, '%.6f')); end
if ~isnan(lonMin), ncwriteatt(ncFile,"/","geospatial_lon_min", num2str(lonMin, '%.6f')); end
if ~isnan(lonMax), ncwriteatt(ncFile,"/","geospatial_lon_max", num2str(lonMax, '%.6f')); end

end

function writeTableRootNetCDF(ncFile, T, metaMap)

if isempty(T) || height(T) == 0
    warning("Table empty, skipping.");
    return
end

nObs   = height(T);
obsDim = "obs";
vars   = string(T.Properties.VariableNames);

for i = 1:numel(vars)
    origName = vars(i);
    ncName   = sanitizeVarName(origName);
    vpath    = char("/" + ncName);

    x = T.(origName);

    % ---- Time handling ----
    if isTimeColumn(origName)
        tDays = convertTimeToDaysSince1979(x);

        nccreate(ncFile, vpath, ...
            "Dimensions", {char(obsDim), nObs}, ...
            "Datatype", "double", ...
            "Format", "netcdf4");
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
        nccreate(ncFile, vpath, ...
            "Dimensions", {char(obsDim), nObs}, ...
            "Datatype", "double", ...
            "Format", "netcdf4");
        ncwrite(ncFile, vpath, double(x));
    else
        s = string(x);
        s(ismissing(s)) = "";
        maxLen = max(strlength(s));
        if maxLen < 1, maxLen = 1; end

        strlenDim = "strlen_" + string(ncName);

        nccreate(ncFile, vpath, ...
            "Dimensions", {char(strlenDim), maxLen, char(obsDim), nObs}, ...
            "Datatype", "char", ...
            "Format", "netcdf4");

        C = char(pad(s, maxLen));   % nObs x maxLen
        C = permute(C, [2 1]);      % maxLen x nObs
        ncwrite(ncFile, vpath, C);
    end

    % Default long_name = table header
    ncwriteatt(ncFile, vpath, "long_name", char(origName));

    % ---- Apply metadata (units/standard_name/comment) from map ----
    key1 = char(origName);
    key2 = char(ncName);

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
s = regexprep(s, "\s+", "_");
s = regexprep(s, "[^A-Za-z0-9_]", "_");
s = regexprep(s, "_+", "_");
s = regexprep(s, "^_+|_+$", "");
if s == "", s = "var"; end
if ~isempty(regexp(s, "^[0-9]", "once"))
    s = "v_" + s;
end
ncName = char(s);
end

% -------- Metadata map for RHO table
function M = buildMetaMap_RHO()
M = containers.Map();

M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north","comment","Coordinates of the ice coring site at the time of sampling");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east","comment","Coordinates of the ice coring site at the time of sampling");
M("Snow height") = struct("standard_name","Snow_thickness","units","m","comment","Average snow or Surface Scattering Layer depth at coring site");

M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");

M("Sea_ice_thickness") = struct("standard_name","sea_ice_thickness","units","m","comment","Distance from the top of ice to the bottom of the ice");
M("Sea_ice_draft") = struct("standard_name","sea_ice_draft","units","m","comment","Distance from water level to bottom of the ice");
M("Core length") = struct("standard_name","core_length","units","m","comment","Total length of the extracted ice core");
M("Depth, ice/snow") = struct("standard_name","depth","units","m","comment","Middle of ice layer sampled, measured from the ice surface");
M("Depth, ice/snow, top/minimum") = struct("standard_name","depth","units","m","comment","Top of ice layer sampled, measured from the ice surface");
M("Depth, ice/snow, bottom/maximum") = struct("standard_name","depth","units","m","comment","Bottom of ice layer sampled, measured from the ice surface");

M("Sea ice salinity") = struct("standard_name","sea_ice_salinity","units","1","comment","Practical salinity measured after melting");
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
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north","comment","Coordinates of the ice coring site at the time of sampling");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east","comment","Coordinates of the ice coring site at the time of sampling");
M("Snow height") = struct("standard_name","Snow_thickness","units","m","comment","Average snow or Surface Scattering Layer depth at coring site");
M("Sea_ice_thickness") = struct("standard_name","sea_ice_thickness","units","m","comment","Distance from the top of ice to the bottom of the ice");
M("Sea_ice_draft") = struct("standard_name","sea_ice_draft","units","m","comment","Distance from water level to bottom of the ice");

M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");

M("Core length") = struct("standard_name","core_length","units","m","comment","Total length of the extracted ice core");
M("Depth, ice/snow") = struct("standard_name","depth","units","m","comment","Middle of ice layer sampled, measured from the ice surface");
M("Temperature, ice/snow") = struct("standard_name","sea_ice_temperature","units","degree_Celsius","comment","In situ ice temperature.");
end

% -------- Metadata map for SALO18 table
function M = buildMetaMap_SALO18()
M = containers.Map();

M("DATE/TIME")  = struct("standard_name","time","units","days since 1979-01-01 00:00:00");
M("LATITUDE")   = struct("standard_name","latitude","units","degree_north","comment","Coordinates of the ice coring site at the time of sampling");
M("LONGITUDE")  = struct("standard_name","longitude","units","degree_east","comment","Coordinates of the ice coring site at the time of sampling");
M("Snow height") = struct("standard_name","Snow_thickness","units","m","comment","Average snow or Surface Scattering Layer depth at coring site");
M("Sea_ice_thickness") = struct("standard_name","sea_ice_thickness","units","m","comment","Distance from the top of ice to the bottom of the ice");
M("Sea_ice_draft") = struct("standard_name","sea_ice_draft","units","m","comment","Distance from water level to bottom of the ice");

M("Ice age") = struct("standard_name","ice_age","comment","First-year ice (FYI), second-year ice (SYI), (second- or multiyear ice) SMYI.");
M("Melt pond") = struct("standard_name","melt_pond","units","1","comment","Sea ice covered (1) or not covered (0) with melt pond.");
M("Ice station visit") = struct("standard_name","station_visit","units","1","comment","Visits a, b, c, d.");
M("Ice station number") = struct("standard_name","station_number","units","1","comment","Ice stations 1, 2, 3.");

M("Core length") = struct("standard_name","core_length","units","m","comment","Total length of the extracted ice core");
M("Depth, ice/snow, top/minimum") = struct("standard_name","depth","units","m","comment","Top of ice layer sampled, measured from the ice surface");
M("Depth, ice/snow, bottom/maximum") = struct("standard_name","depth","units","m","comment","Bottom of ice layer sampled, measured from the ice surface");
M("Sea ice salinity") = struct("standard_name","sea_ice_salinity","units","1","comment","Practical salinity measured after melting");
end

% -------- Helpers to infer time/geo coverage from one table
function [tStartISO, tEndISO] = inferTimeCoverageISO_single(T)
tStartISO = "";
tEndISO   = "";

if isempty(T), return, end
if ~ismember("DATE/TIME", T.Properties.VariableNames), return, end
if ~isdatetime(T.("DATE/TIME")), return, end

tAll = T.("DATE/TIME");
tAll = tAll(~ismissing(tAll));
if isempty(tAll), return, end

tmin = min(tAll);
tmax = max(tAll);

if isempty(tmin.TimeZone), tmin.TimeZone = "UTC"; end
if isempty(tmax.TimeZone), tmax.TimeZone = "UTC"; end

tStartISO = char(datetime(tmin,"TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss"));
tEndISO   = char(datetime(tmax,"TimeZone","UTC","Format","yyyy-MM-dd HH:mm:ss"));
end

function [latMin, latMax, lonMin, lonMax] = inferGeoCoverage_single(T)
latMin = NaN; latMax = NaN; lonMin = NaN; lonMax = NaN;

if isempty(T), return, end

if ismember("LATITUDE", T.Properties.VariableNames)
    latAll = double(T.LATITUDE);
    latAll = latAll(~isnan(latAll));
    if ~isempty(latAll)
        latMin = min(latAll);
        latMax = max(latAll);
    end
end

if ismember("LONGITUDE", T.Properties.VariableNames)
    lonAll = double(T.LONGITUDE);
    lonAll = lonAll(~isnan(lonAll));
    if ~isempty(lonAll)
        lonMin = min(lonAll);
        lonMax = max(lonAll);
    end
end
end