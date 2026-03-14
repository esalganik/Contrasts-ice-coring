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

% ---------- Legend ----------
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

%% NetCDF import
clear; clc; close all;
filename = fullfile(pwd, 'Export', 'Contrasts_physical_properties_coring.nc');
% ncdisp(filename)
date_time_raw = ncread(filename, '/RHO/DATE_TIME');
temperature = ncread(filename, '/RHO/Temperature_ice_snow');
Core_number_RHO = ncread(filename, '/RHO/Core_number_RHO');
Station = ncread(filename, '/RHO/Ice_station_number');
Melt_pond = ncread(filename, '/RHO/Melt_pond');
IceThickness = ncread(filename, '/RHO/Sea_ice_thickness');
Depth = ncread(filename, '/RHO/Depth_ice_snow_top_minimum');
info = ncinfo(filename, '/RHO/DATE_TIME');
units = ncreadatt(filename, '/RHO/DATE_TIME', 'units');
refDate = datetime(extractAfter(units, 'since '), 'InputFormat','yyyy-MM-dd HH:mm:ss');
date_time = refDate + days(date_time_raw);
T = table(Core_number_RHO, Station, date_time, temperature, Melt_pond,IceThickness);
T = sortrows(T, {'Core_number_RHO','date_time'});
[~, ~, idx] = unique(T.Core_number_RHO);
T_bot = accumarray(idx, T.temperature, [], @(x) x(end));
% Handle datetime via datenum trick
tNum = datenum(T.date_time);
time_bot_num = accumarray(idx, tNum, [], @(x) x(end));
time_bot = datetime(time_bot_num, 'ConvertFrom','datenum');
% Get station of each core (last entry per core)
station_bot = accumarray(idx, T.Station, [], @(x) x(end));
Melt_pond_bot = accumarray(idx, T.Melt_pond, [], @(x) x(end));
IceThickness_bot = accumarray(idx, T.IceThickness, [], @(x) x(end));

figure
stations_unique = unique(station_bot);

% marker types for stations
markerTypes = {'o','^','s','d','v','>','<','p','h'};
markerTypes = markerTypes(1:length(stations_unique));

% colormap for ice thickness
load("lipari.mat")
cmap = lipari;
colormap(cmap)

% normalize thickness for coloring
th_min = min(IceThickness_bot);
th_max = max(IceThickness_bot);

hold on; box on;
for k = 1:length(stations_unique)

    ind = station_bot == stations_unique(k);

    if any(ind)

        scatter(time_bot(ind), T_bot(ind), ...
            60, IceThickness_bot(ind), ...
            markerTypes{k}, ...
            'filled');

    end
end

ylabel('Bottom Temperature')
title('Bottom Temperature vs Time')
grid on
colorbar
clim([th_min th_max])

legend("Station " + string(stations_unique), ...
       'Location','northoutside','numcolumns',3,'box','off')

%% NetCDF import + profile plotting (3 panels by Station)
% Depth vs Temperature for each core
% Color = colormap (e.g., batlow/lipari)
% LineStyle = Visit (a,b,c,d)
clear; clc; close all;

filename = fullfile(pwd, 'Export', 'Contrasts_physical_properties_coring.nc');

% --- Read NetCDF variables
date_time_raw    = ncread(filename, '/RHO/DATE_TIME');
temperature      = ncread(filename, '/RHO/Temperature_ice_snow');
Core_number_RHO  = ncread(filename, '/RHO/Core_number_RHO');
Station          = ncread(filename, '/RHO/Ice_station_number');
Visit            = ncread(filename, '/RHO/Ice_station_visit');   % a,b,c,d (often char)
Depth            = ncread(filename, '/RHO/Depth_ice_snow_top_minimum');

units   = ncreadatt(filename, '/RHO/DATE_TIME', 'units');
refDate = datetime(extractAfter(units, 'since '), 'InputFormat','yyyy-MM-dd HH:mm:ss');
date_time = refDate + days(date_time_raw);

% --- Make Visit a string vector
Visit = lower(strtrim(string(Visit(:))));   % e.g. "a","b","c","d"

% Load colormap from MAT (example: batlow.mat contains variable "batlow")
% If you want lipari instead:
%   load("lipari.mat","lipari"); cmap = squeeze(lipari);
load("batlow.mat","batlow")
cmap = squeeze(batlow);          % ensure Nx3
if size(cmap,2) ~= 3
    error('Colormap in MAT must be Nx3.');
end

% Build long table for profiles
Tfull = table(Core_number_RHO(:), Station(:), Visit(:), date_time(:), temperature(:), Depth(:), ...
    'VariableNames', {'Core_number_RHO','Station','Visit','date_time','temperature','Depth'});

% Remove invalid rows
good = isfinite(Tfull.Core_number_RHO) & isfinite(Tfull.Station) & isfinite(Tfull.temperature) & isfinite(Tfull.Depth);
Tfull = Tfull(good,:);

% Sort so "end" corresponds to last time per core (same as your bottom-temp approach)
Tfull = sortrows(Tfull, {'Core_number_RHO','date_time'});

% Per-core indexing
[core_list, ~, idxC] = unique(Tfull.Core_number_RHO);

% Station per core (last)
station_core = accumarray(idxC, Tfull.Station, [], @(x) x(end));

% Date per core (last)
date_core_num = accumarray(idxC, datenum(Tfull.date_time), [], @(x) x(end));
date_core = datetime(date_core_num, 'ConvertFrom','datenum');

% --- Visit per core (last) using numeric codes + accumarray
visitTypes = ["a","b","c","d"];
lineStyles = ["-","--",":","-."];

visit_code = zeros(height(Tfull),1);  % 0 = unknown
for j = 1:numel(visitTypes)
    visit_code(Tfull.Visit == visitTypes(j)) = j;
end
visit_code_core = accumarray(idxC, visit_code, [], @(x) x(end));  % numeric per core

% Stations to plot (first 3 found). If you want fixed: stations_to_plot = [1 2 3];
stations_unique = unique(station_core(isfinite(station_core)));
stations_unique = stations_unique(:)';
nPanels = min(3, numel(stations_unique));
if nPanels == 0
    error('No stations found after filtering. Check Station data.');
end
stations_to_plot = stations_unique(1:nPanels);

% Plot
figure
tiledlayout(1,nPanels,"TileSpacing","compact","Padding","compact")

for i = 1:nPanels
    st = stations_to_plot(i);

    nexttile
    hold on; box on;

    cores_here = core_list(station_core == st);
    nCores = numel(cores_here);

    % Colors from colormap (evenly spaced)
    ic = round(linspace(1, size(cmap,1), max(nCores,1)));
    colors = cmap(ic,:);

    legendHandles = gobjects(0);
    legendLabels  = strings(0);

    for k = 1:nCores
        c = cores_here(k);

        ind = Tfull.Core_number_RHO == c;

        z = Tfull.Depth(ind);
        t = Tfull.temperature(ind);

        ok = isfinite(z) & isfinite(t);
        if ~any(ok), continue; end

        % Sort by depth for nice profile
        [zS, ii] = sort(z(ok));
        tS = t(ok);
        tS = tS(ii);

        % Line style based on visit code
        vcode = visit_code_core(core_list == c);
        style = "-";
        vtxt  = "unk";
        if vcode >= 1 && vcode <= numel(lineStyles)
            style = lineStyles(vcode);
            vtxt  = visitTypes(vcode);
        end

        h = plot(tS, zS, ...
            'LineWidth', 1.5, ...
            'Color', colors(k,:), ...
            'LineStyle', style);

        % Legend label includes date and visit
        legendHandles(end+1) = h;
        legendLabels(end+1)  = string(date_core(core_list==c), "MMM dd") + " (visit " + vtxt + ")";
    end

    set(gca,'YDir','reverse')
    grid on
    xlabel('Temperature')
    ylabel('Depth (m)')
    title("Station " + string(st))

    if ~isempty(legendHandles)
        legend(legendHandles, legendLabels, 'Location','best', 'Box','off')
    end
end

%% NetCDF import
clear; clc; close all;
% ncdisp(filename)
filename = fullfile(pwd, 'Export', 'Contrasts_physical_properties_coring.nc');

date_time_raw    = ncread(filename, '/RHO/DATE_TIME');
temperature      = ncread(filename, '/RHO/Temperature_ice_snow');
Core_number_RHO  = ncread(filename, '/RHO/Core_number_RHO');
Station          = ncread(filename, '/RHO/Ice_station_number');
Visit            = ncread(filename, '/RHO/Ice_station_visit'); string(Visit(:))
Melt_pond        = ncread(filename, '/RHO/Melt_pond');
IceThickness     = ncread(filename, '/RHO/Sea_ice_thickness');
Depth            = ncread(filename, '/RHO/Depth_ice_snow_top_minimum');

units   = ncreadatt(filename, '/RHO/DATE_TIME', 'units');
refDate = datetime(extractAfter(units, 'since '), 'InputFormat','yyyy-MM-dd HH:mm:ss');
date_time = refDate + days(date_time_raw);

load("batlow.mat","batlow")
cmap = batlow;
Tfull = table(Core_number_RHO(:), Station(:), Visit(:), date_time(:), temperature(:), Depth(:), ...
    'VariableNames', {'Core_number_RHO','Station','Visit','date_time','temperature','Depth'});

% Remove invalid rows
good = isfinite(Tfull.Core_number_RHO) & isfinite(Tfull.Station) & isfinite(Tfull.temperature) & isfinite(Tfull.Depth);
Tfull = Tfull(good,:);

% Sort so "last" corresponds to last date_time for each core (like your bottom-temp approach)
Tfull = sortrows(Tfull, {'Core_number_RHO','date_time'});

% Per-core indices
[core_list, ~, idxC] = unique(Tfull.Core_number_RHO);

% Station per core (last)
station_core = accumarray(idxC, Tfull.Station, [], @(x) x(end));

% Date per core (last)
date_core_num = accumarray(idxC, datenum(Tfull.date_time), [], @(x) x(end));
date_core = datetime(date_core_num, 'ConvertFrom','datenum');

% Choose up to 3 stations to plot (edit if you want fixed [1 2 3])
stations_unique = unique(station_core(isfinite(station_core)));
stations_unique = stations_unique(:)';
nPanels = min(3, numel(stations_unique));
if nPanels == 0
    error('No stations found after filtering. Check Station data.');
end

% Plot: 3 panels (by station), each core = one profile, legend = date, colors = lipari
figure
tiledlayout(1,3,"TileSpacing","compact","Padding","compact")

for i = 1:3
    st = stations_unique(i);

    nexttile
    hold on; box on;

    cores_here = core_list(station_core == st);
    nCores = numel(cores_here);

    % Lipari colors for this station panel
    idx = round(linspace(1, size(cmap,1), max(nCores,1)));
    colors = cmap(idx,:);

    legendHandles = gobjects(0);
    legendLabels  = strings(0);

    for k = 1:nCores
        c = cores_here(k);

        ind = Tfull.Core_number_RHO == c;

        z = Tfull.Depth(ind);
        t = Tfull.temperature(ind);

        good = isfinite(z) & isfinite(t);
        if ~any(good)
            continue
        end

        % sort by depth so profile draws nicely
        [zS, ii] = sort(z(good));
        tS = t(good);
        tS = tS(ii);

        h = plot(tS, zS, '-', ...
            'LineWidth', 1.5, ...
            'Color', colors(k,:));

        legendHandles(end+1) = h;
        legendLabels(end+1)  = string(date_core(core_list==c), "MMM dd");
    end

    set(gca,'YDir','reverse')
    grid on
    xlabel('Temperature')
    ylabel('Depth (m)')
    title("Station " + string(st))

    if ~isempty(legendHandles)
        legend(legendHandles, legendLabels, 'Location','best', 'Box','off')
    end
end
set(gcf,'Units','inches','Position',[2 5 7 4.5]); box on

%% Functions
function x = toNum(x)
    if iscell(x) || isstring(x)
        x = str2double(string(x));
    end
    x = double(x);
end

function y = earliestNonMissing(x)
    x = x(~ismissing(x));
    if isempty(x), y = NaT; else, y = min(x); end
end

function y = firstNonNaN(x)
    x = double(x);
    k = find(~isnan(x), 1, "first");
    if isempty(k), y = NaN; else, y = x(k); end
end