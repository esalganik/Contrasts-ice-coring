clear; close all; clc

projectRoot = pwd;

inputFolder  = fullfile(projectRoot, "data", "final");
outputFolder = fullfile(projectRoot, "figures");

if ~isfolder(outputFolder)
    mkdir(outputFolder);
end

INFILE_PROCESSED = fullfile(inputFolder, "core_data_processed.mat");

if ~isfile(INFILE_PROCESSED)
    error("Processed MAT not found: %s", INFILE_PROCESSED);
end

load(INFILE_PROCESSED, "T_all_proc");

rho = T_all_proc.rho;

tRowR = rho.Date;
tRowR.TimeZone = '';

stRowR = double(rho.StationNumber);

if ismember("CoreID_RHO", rho.Properties.VariableNames)
    coreID_R = double(rho.CoreID_RHO);
else
    [~,~,coreID_R] = unique(string(rho.SourceFile), "stable");
    coreID_R = double(coreID_R);
end

salRowR   = toNum(rho.Salinity_used);
denInsRow = toNum(rho.rho_si);
tempRow   = toNum(rho.Temperature_interp);
thickRow  = toNum(rho{:,10});

coreListR = unique(coreID_R(coreID_R>0), "stable");
nR = numel(coreListR);

CoreTimeR    = NaT(nR,1);
CoreStationR = NaN(nR,1);
AvgSalR      = NaN(nR,1);
AvgIns       = NaN(nR,1);
AvgTemp      = NaN(nR,1);
AvgThick     = NaN(nR,1);

for i = 1:nR
    idx = coreID_R == coreListR(i);
    CoreTimeR(i)    = min(tRowR(idx),[],'omitmissing');
    CoreStationR(i) = stRowR(find(idx,1,'first'));
    AvgSalR(i)      = mean(salRowR(idx),'omitnan');
    AvgIns(i)       = mean(denInsRow(idx),'omitnan');
    AvgTemp(i)      = mean(tempRow(idx),'omitnan');
    AvgThick(i)     = thickRow(find(idx,1,'first'));
end

hasSALO = isfield(T_all_proc,"SALO18");

if hasSALO
    salo = T_all_proc.SALO18;

    if ismember("Date", salo.Properties.VariableNames)
        tRowS = salo.Date;
    elseif ismember("Time_best", salo.Properties.VariableNames)
        tRowS = salo.Time_best;
    else
        tRowS = salo.GPS_Time;
    end
    tRowS.TimeZone = '';

    stRowS  = double(salo.StationNumber);
    salRowS = toNum(salo.Salinity);

    if ismember("CoreID_SALO18", salo.Properties.VariableNames)
        coreID_S = double(salo.CoreID_SALO18);
    else
        [~,~,coreID_S] = unique(string(salo.SourceFile), "stable");
        coreID_S = double(coreID_S);
    end

    coreListS = unique(coreID_S(coreID_S>0), "stable");
    nS = numel(coreListS);

    CoreTimeS    = NaT(nS,1);
    CoreStationS = NaN(nS,1);
    AvgSalS      = NaN(nS,1);

    for i = 1:nS
        idx = coreID_S == coreListS(i);
        CoreTimeS(i)    = min(tRowS(idx),[],'omitmissing');
        CoreStationS(i) = stRowS(find(idx,1,'first'));
        AvgSalS(i)      = mean(salRowS(idx),'omitnan');
    end
else
    CoreTimeS = NaT(0,1);
    CoreStationS = NaN(0,1);
    AvgSalS = NaN(0,1);
end

stations = unique([CoreStationR; CoreStationS]);
stations = stations(ismember(stations,[1 2 3]));

cols = [
    1, 61, 115;    % Station 1  (#013D73)
    58, 174, 140;  % Station 2  (#3AAE8C)
    245, 174, 16   % Station 3  (#F5AE10)
] / 255;

figure('Color','w');
tiledlayout(2,2,'Padding','compact','TileSpacing','compact');

ax1 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
    idxR = CoreStationR==st;
    scatter(CoreTimeR(idxR), AvgSalR(idxR), 70, 'o', ...
        'MarkerFaceColor',c,'MarkerEdgeColor',c);

    if hasSALO
        idxS = CoreStationS==st;
        scatter(CoreTimeS(idxS), AvgSalS(idxS), 70, 'o', ...
            'MarkerFaceColor','none','MarkerEdgeColor',c,'LineWidth',1.2);
    end
end
ylabel("Ice salinity (g/kg)");
title("(a) Ice salinity");

ax2 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
    idx = CoreStationR==st;
    scatter(CoreTimeR(idx), AvgIns(idx), 70, 'o', ...
        'MarkerFaceColor',c,'MarkerEdgeColor',c);
end
ylabel("Ice density (kg m^{-3})");
title("(b) In situ density");

ax3 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
    idx = CoreStationR==st;
    scatter(CoreTimeR(idx), AvgTemp(idx), 70, 'o', ...
        'MarkerFaceColor',c,'MarkerEdgeColor',c);
end
ylabel("Ice temperature (°C)");
title("(c) Ice temperature");

ax4 = nexttile; hold on; grid on; box on;
for st = stations'
    c = cols(st,:);
    idx = CoreStationR==st;
    scatter(CoreTimeR(idx), AvgThick(idx), 70, 'o', ...
        'MarkerFaceColor',c,'MarkerEdgeColor',c);
end
ylabel("Ice thickness (m)");
title("(d) Ice thickness");

linkaxes([ax1 ax2 ax3 ax4],'x');
ax1.XTickLabel = [];
ax2.XTickLabel = [];

set([ax1 ax2 ax3 ax4],'FontSize',10)
set(ax1,'YLim',[1 4],'YLimMode','manual')
set(ax2,'YLim',[840 920],'YLimMode','manual')
set(ax3,'YLim',[-1.5 -0.4],'YLimMode','manual')

hSt = gobjects(0,1);
labSt = strings(0,1);

for st = stations'
    hSt(end+1,1) = plot(ax1,nan,nan,'o', ...
        'MarkerSize',12, ...
        'MarkerFaceColor',cols(st,:), ...
        'MarkerEdgeColor',cols(st,:), ...
        'LineStyle','none');
    labSt(end+1,1) = "Station " + string(st);
end

hFilled = plot(ax1,nan,nan,'o', ...
    'MarkerSize',12, ...
    'MarkerFaceColor','k', ...
    'MarkerEdgeColor','k', ...
    'LineStyle','none');

hOpen = plot(ax1,nan,nan,'o', ...
    'MarkerSize',12, ...
    'MarkerFaceColor','none', ...
    'MarkerEdgeColor','k', ...
    'LineStyle','none');

hAll   = [hSt; hFilled; hOpen];
labAll = [labSt; "Density core"; "Salinity core"];

lgd = legend(hAll,labAll,'Location','northoutside','NumColumns',numel(hAll));
lgd.Layout.Tile = 'north';

set(gcf,'Units','inches','Position',[4 4 12 5])
outFigure = fullfile(outputFolder, "coring_CONTRASTS_overview.png");
exportgraphics(gcf, outFigure, "Resolution", 300);
fprintf("Figure saved to %s\n", outFigure);

function x = toNum(x)
    x = double(x);
end