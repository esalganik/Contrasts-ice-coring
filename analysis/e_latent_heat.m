% e_bottom_n_sections_latent_density_thickness_vs_time.m
% Analysis script.
% Not part of the published data-processing workflow.
% Uses exported density NetCDF products to estimate
% effective latent heat of the lowermost sea-ice layers.
%
% Uses the deepest n sections of each density core.
%
% Plot:
%   (1) effective latent heat of the lowermost ice vs time
%       - filled markers: latent-only estimate
%       - open markers: enthalpy-style estimate
%   (2) mean in situ density of the lowermost ice vs time
%   (3) total ice thickness vs time
%
% Run from repository root.

clear; close all; clc

projectRoot = pwd;

nBottomSections = 3;   % CHANGE THIS: 1 = bottom 5 cm, 2 = bottom 10 cm, etc.

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

% Remove empty rows
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

% ========================= Effective latent heat estimates =========================

Lf = 333.4e3;                        % J kg^-1

T = CoreTemp;                        % deg C

% Temperature-dependent properties
rhoIce = 916.8 - 0.1403 .* T;        % kg m^-3, Pounder (1965)

ci = 2112.2 + 7.6973 .* T;           % J kg^-1 K^-1, Weast (1971)

Sb = -18.7 .* T ...
     - 0.519 .* T.^2 ...
     - 0.00535 .* T.^3;             % brine salinity, Vancoppenolle et al.

cb = 4208.8 ...
     + 111.71 .* T ...
     + 3.5611 .* T.^2 ...
     + 0.052168 .* T.^3;            % J kg^-1 K^-1, Fofonoff and Millard Jr. (1983)

rhoBrine = 1000.3 ...
         + 0.78237 .* Sb ...
         + 2.8008e-4 .* Sb.^2;      % kg m^-3, Schwerdtfeger (1963)

SolidFraction = 1 - CoreBrine - CoreGas;
SolidFraction(SolidFraction < 0 | SolidFraction > 1) = NaN;

% Method 1: latent heat only, based on solid ice volume fraction
EffectiveLatentHeat_simple_Jm3 = rhoIce .* SolidFraction .* Lf;
EffectiveLatentHeat_simple_MJm3 = EffectiveLatentHeat_simple_Jm3 / 1e6;

% Method 2: enthalpy-style estimate including sensible heat of ice and brine
EffectiveLatentHeat_enthalpy_Jm3 = ...
      rhoIce .* SolidFraction .* Lf ...
    + rhoIce .* SolidFraction .* ci .* abs(T) ...
    + rhoBrine .* CoreBrine .* cb .* abs(T);

EffectiveLatentHeat_enthalpy_MJm3 = EffectiveLatentHeat_enthalpy_Jm3 / 1e6;

EnthalpyCorrection_MJm3 = ...
    EffectiveLatentHeat_enthalpy_MJm3 - EffectiveLatentHeat_simple_MJm3;

BottomNTable = table( ...
    CoreTime, CoreID, CoreStation, CoreNSections, CoreDepthMean, ...
    CoreThickness, CoreBrine, CoreGas, SolidFraction, ...
    CoreDensity, CoreTemp, Sb, rhoIce, rhoBrine, ci, cb, ...
    EffectiveLatentHeat_simple_Jm3, ...
    EffectiveLatentHeat_enthalpy_Jm3, ...
    EnthalpyCorrection_MJm3);

% disp(BottomNTable)

% outTable = fullfile(figFolder, ...
%     sprintf("bottom_%d_sections_latent_density_thickness_vs_time.csv", nBottomSections));
% writetable(BottomNTable, outTable);
% fprintf("Saved table to %s\n", outTable);

fprintf("Mean simple effective latent heat   = %.1f MJ m^-3\n", ...
    mean(EffectiveLatentHeat_simple_MJm3, "omitnan"));
fprintf("Std simple effective latent heat    = %.1f MJ m^-3\n", ...
    std(EffectiveLatentHeat_simple_MJm3, "omitnan"));
fprintf("Mean enthalpy effective latent heat = %.1f MJ m^-3\n", ...
    mean(EffectiveLatentHeat_enthalpy_MJm3, "omitnan"));
fprintf("Std enthalpy effective latent heat  = %.1f MJ m^-3\n", ...
    std(EffectiveLatentHeat_enthalpy_MJm3, "omitnan"));
fprintf("Mean enthalpy correction            = %.1f MJ m^-3\n", ...
    mean(EnthalpyCorrection_MJm3, "omitnan"));
fprintf("Mean bottom density                 = %.1f kg m^-3\n", ...
    mean(CoreDensity, "omitnan"));

% ========================= Plot =========================

figure("Color","w");
tiledlayout(3,1,"Padding","compact","TileSpacing","compact");

ax1 = nexttile; hold on; grid on; box on

% Filled markers: simple latent-only estimate
plotByStation(CoreTime, EffectiveLatentHeat_simple_MJm3, CoreStation, cols, true)

% Open markers: enthalpy-style estimate
plotByStation(CoreTime, EffectiveLatentHeat_enthalpy_MJm3, CoreStation, cols, false)

ylabel("Effective latent heat (MJ m^{-3})")
title(sprintf("(a) Lowermost %d section(s): effective latent heat", nBottomSections))

ax2 = nexttile; hold on; grid on; box on
plotByStation(CoreTime, CoreDensity, CoreStation, cols, true)
ylabel("In situ density (kg m^{-3})")
title(sprintf("(b) Lowermost %d section(s): mean sea-ice density", nBottomSections))

ax3 = nexttile; hold on; grid on; box on
plotByStation(CoreTime, CoreThickness, CoreStation, cols, true)
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

h = gobjects(5,1);

for st = 1:3
    h(st) = plot(ax1, NaN, NaN, "o", ...
        "MarkerSize", 8, ...
        "MarkerFaceColor", cols(st,:), ...
        "MarkerEdgeColor", cols(st,:), ...
        "LineStyle", "none");
end

h(4) = plot(ax1, NaN, NaN, "o", ...
    "MarkerSize", 8, ...
    "MarkerFaceColor", "k", ...
    "MarkerEdgeColor", "k", ...
    "LineStyle", "none");

h(5) = plot(ax1, NaN, NaN, "o", ...
    "MarkerSize", 8, ...
    "MarkerFaceColor", "none", ...
    "MarkerEdgeColor", "k", ...
    "LineWidth", 1.4, ...
    "LineStyle", "none");

lgd = legend(h, ...
    ["Station 1","Station 2","Station 3", ...
     "Latent only","Enthalpy"], ...
    "Location","northoutside", ...
    "NumColumns",5);
lgd.Layout.Tile = "north";

set([ax1 ax2 ax3], "FontSize", 10)
set(gcf, "Units", "inches", "Position", [4 4 10 8])

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

function plotByStation(x, y, st, cols, filledMarkers)
    for s = 1:3
        idx = st == s & ~isnan(y);

        if filledMarkers
            scatter(x(idx), y(idx), 70, "o", ...
                "MarkerFaceColor", cols(s,:), ...
                "MarkerEdgeColor", cols(s,:));
        else
            scatter(x(idx), y(idx), 70, "o", ...
                "MarkerFaceColor", "none", ...
                "MarkerEdgeColor", cols(s,:), ...
                "LineWidth", 1.4);
        end
    end
end