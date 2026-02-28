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
tiledlayout(2,1)

stations_unique = unique(station_bot);
colors = lines(length(stations_unique));

markerTypes = {'o','^'};   % 0 = no melt pond, 1 = melt pond

% Panel 1: Temp vs Time
nexttile
hold on; box on;
legendHandles = [];
for k = 1:length(stations_unique)
    for mp = 0:1

        ind = station_bot == stations_unique(k) & ...
              Melt_pond_bot == mp;

        if any(ind)

            h = plot(time_bot(ind), T_bot(ind), ...
                markerTypes{mp+1}, ...
                'LineStyle','none', ...
                'Color', colors(k,:), ...
                'MarkerFaceColor', colors(k,:));

            % Only create legend entry once per station
            if mp == 0
                legendHandles(k) = h;
            end
        end
    end
end

ylabel('Bottom Temperature')
title('Bottom Temperature vs Time')
grid on

% Panel 2: Temp vs Ice Thickness
nexttile
hold on; box on;
for k = 1:length(stations_unique)
    for mp = 0:1

        ind = station_bot == stations_unique(k) & ...
              Melt_pond_bot == mp;

        if any(ind)
            plot(T_bot(ind),IceThickness_bot(ind), ...
                markerTypes{mp+1}, ...
                'LineStyle','none', ...
                'Color', colors(k,:), ...
                'MarkerFaceColor', colors(k,:));
        end
    end
end

xlabel('Bottom Temperature')
ylabel('Ice Thickness')
title('Bottom Temperature vs Ice Thickness')
grid on

legend(legendHandles, ...
       "Station " + string(stations_unique), ...
       'Location','northoutside','numcolumns',3,'box','off')

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