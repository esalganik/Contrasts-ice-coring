% Analysis script.
% Not part of the published data-processing workflow.
% Uses exported density NetCDF products to estimate
% effective latent heat of the lowermost sea-ice layers.

% e_bottom_n_sections_latent_density_thickness_vs_time.m
% Uses the deepest n sections of each density core.
%
% Plot:
%   (1) effective latent heat of the lowermost ice vs time
%   (2) mean in situ density of the lowermost ice vs time
%   (3) total ice thickness vs time
%
% Run from repository root.

clear; close all; clc

projectRoot = pwd;

nBottomSections = 3;   % 1 = bottom 5 cm, 2 = bottom 10 cm

ncFile = fullfile(projectRoot, "data", "final", "netcdf", ...
    "Contrasts_coring_density.nc");

if ~isfile(ncFile)
    error("NetCDF file not found: %s", ncFile);
end

figFolder = fullfile(projectRoot, "figures");
if ~isfolder(figFolder)
    mkdir(figFolder);
end

cols = [
    1, 61, 115;     % Station 1 (#013D73)
    58, 174, 140;   % Station 2 (#3AAE8C)
    245, 174, 16    % Station 3 (#F5AE10)
] / 255;

% ========================= Import variables =========================

timeDays = ncread(ncFile, "DATE_TIME");
timeAll = datetime(1979,1,1,0,0,0,"TimeZone","UTC") + days(timeDays);
timeAll.TimeZone = "";

station  = ncread(ncFile, "Ice_station_number");
coreID   = ncread(ncFile, "Core_number_RHO");

depthTop = ncread(ncFile, "Depth_ice_snow_top_minimum");
depthBot = ncread(ncFile, "Depth_ice_snow_bottom_maximum");

Vb       = ncread(ncFile, "Volume_brine");
Vg       = ncread(ncFile, "Volume_gas");
rhoBulk  = ncread(ncFile, "Density_ice");
Tice     = ncread(ncFile, "Temperature_ice_snow");
iceThick = ncread(ncFile, "Sea_ice_thickness");

% ========================= Deepest n sections per core =========================

coreList = unique(coreID(~isnan(coreID) & coreID > 0), "stable");
nCores = numel(coreList);

CoreTime      = NaT(nCores,1);
CoreID        = NaN(nCores,1);
CoreStation   = NaN(nCores,1);
CoreDepthMean = NaN(nCores,1);
CoreThickness = NaN(nCores,1);
CoreBrine     = NaN(nCores,1);
CoreGas       = NaN(nCores,1);
CoreDensity   = NaN(nCores,1);
CoreTemp      = NaN(nCores,1);
CoreNSections = NaN(nCores,1);

for i = 1:nCores

    idx = coreID == coreList(i);
    rows = find(idx);

    if isempty(rows)
        continue
    end

    [~, order] = sort(depthBot(rows), ...
        "descend", ...
        "MissingPlacement", "last");

    nUse = min(nBottomSections, numel(rows));
    useRows = rows(order(1:nUse));

    dz = depthBot(useRows) - depthTop(useRows);
    dz(dz <= 0 | isnan(dz)) = NaN;

    CoreTime(i)      = min(timeAll(useRows), [], "omitmissing");
    CoreID(i)        = coreList(i);
    CoreStation(i)   = station(useRows(1));
    CoreNSections(i) = nUse;

    CoreDepthMean(i) = weightedMean(depthBot(useRows), dz);
    CoreThickness(i) = weightedMean(iceThick(useRows), dz);
    CoreBrine(i)     = weightedMean(Vb(useRows), dz);
    CoreGas(i)       = weightedMean(Vg(useRows), dz);
    CoreDensity(i)   = weightedMean(rhoBulk(useRows), dz);
    CoreTemp(i)      = weightedMean(Tice(useRows), dz);

end

% Remove empty rows, if any
valid = ~isnat(CoreTime) & ~isnan(CoreID);

CoreTime      = CoreTime(valid);
CoreID        = CoreID(valid);
CoreStation   = CoreStation(valid);
CoreDepthMean = CoreDepthMean(valid);
CoreThickness = CoreThickness(valid);
CoreBrine     = CoreBrine(valid);
CoreGas       = CoreGas(valid);
CoreDensity   = CoreDensity(valid);
CoreTemp      = CoreTemp(valid);
CoreNSections = CoreNSections(valid);

% Sort by time
[CoreTime, I] = sort(CoreTime);

CoreID        = CoreID(I);
CoreStation   = CoreStation(I);
CoreDepthMean = CoreDepthMean(I);
CoreThickness = CoreThickness(I);
CoreBrine     = CoreBrine(I);
CoreGas       = CoreGas(I);
CoreDensity   = CoreDensity(I);
CoreTemp      = CoreTemp(I);
CoreNSections = CoreNSections(I);

% ========================= Simple effective latent heat =========================

Lf = 333.4e3;                        % J kg^-1
rhoIce = 916.8 - 0.1403 .* CoreTemp; % kg m^-3

SolidFraction = 1 - CoreBrine - CoreGas;
SolidFraction(SolidFraction < 0 | SolidFraction > 1) = NaN;

EffectiveLatentHeat_Jm3 = rhoIce .* SolidFraction .* Lf;
EffectiveLatentHeat_MJm3 = EffectiveLatentHeat_Jm3 / 1e6;

BottomNTable = table( ...
    CoreTime, CoreID, CoreStation, CoreNSections, CoreDepthMean, ...
    CoreThickness, CoreBrine, CoreGas, SolidFraction, ...
    CoreDensity, CoreTemp, EffectiveLatentHeat_Jm3);

% disp(BottomNTable)

% outTable = fullfile(figFolder, ...
%     sprintf("bottom_%d_sections_latent_density_thickness_vs_time.csv", nBottomSections));
% writetable(BottomNTable, outTable);
% fprintf("Saved table to %s\n", outTable);

fprintf("Mean effective latent heat = %.1f MJ m^-3\n", ...
    mean(EffectiveLatentHeat_MJm3, "omitnan"));
fprintf("Std effective latent heat  = %.1f MJ m^-3\n", ...
    std(EffectiveLatentHeat_MJm3, "omitnan"));
fprintf("Mean bottom density        = %.1f kg m^-3\n", ...
    mean(CoreDensity, "omitnan"));

% ========================= Plot =========================

figure("Color","w");
tiledlayout(3,1,"Padding","compact","TileSpacing","compact");

ax1 = nexttile; hold on; grid on; box on
plotByStation(CoreTime, EffectiveLatentHeat_MJm3, CoreStation, cols)
ylabel("Effective latent heat (MJ m^{-3})")
title(sprintf("(a) Lowermost %d section(s): effective latent heat", nBottomSections))

ax2 = nexttile; hold on; grid on; box on
plotByStation(CoreTime, CoreDensity, CoreStation, cols)
ylabel("In situ density (kg m^{-3})")
title(sprintf("(b) Lowermost %d section(s): mean sea-ice density", nBottomSections))

ax3 = nexttile; hold on; grid on; box on
plotByStation(CoreTime, CoreThickness, CoreStation, cols)
ylabel("Ice thickness (m)")
title("(c) Core-site ice thickness")
xlabel("Date")

linkaxes([ax1 ax2 ax3], "x")

ticks = dateshift(min(CoreTime),"start","day"): ...
        caldays(7): ...
        dateshift(max(CoreTime),"start","day");

set([ax1 ax2 ax3], "XTick", ticks)
xtickformat(ax1, "dd MMM")
xtickformat(ax2, "dd MMM")
xtickformat(ax3, "dd MMM")

h = gobjects(3,1);
for st = 1:3
    h(st) = plot(ax1, NaN, NaN, "o", ...
        "MarkerSize", 8, ...
        "MarkerFaceColor", cols(st,:), ...
        "MarkerEdgeColor", cols(st,:), ...
        "LineStyle", "none");
end

lgd = legend(h, ["Station 1","Station 2","Station 3"], ...
    "Location","northoutside", ...
    "NumColumns",3);
lgd.Layout.Tile = "north";

set([ax1 ax2 ax3], "FontSize", 10)
set(gcf, "Units", "inches", "Position", [4 4 9 8])

outFig = fullfile(figFolder, ...
    sprintf("bottom_%d_sections_latent_density_thickness_vs_time.png", nBottomSections));

exportgraphics(gcf, outFig, "Resolution", 300);
fprintf("Saved figure to %s\n", outFig);

%% ========================= Helper functions =========================

function mu = weightedMean(x, w)
    x = double(x);
    w = double(w);

    ok = ~isnan(x) & ~isnan(w) & w > 0;

    if ~any(ok)
        mu = NaN;
    else
        mu = sum(x(ok) .* w(ok)) ./ sum(w(ok));
    end
end

function plotByStation(x, y, st, cols)
    for s = 1:3
        idx = st == s & ~isnan(y);

        scatter(x(idx), y(idx), 70, "o", ...
            "MarkerFaceColor", cols(s,:), ...
            "MarkerEdgeColor", cols(s,:));
    end
end