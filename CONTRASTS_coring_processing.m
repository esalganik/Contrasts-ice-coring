%% PROCESS ONLY: load imported -> process -> format/reorder -> rename -> round -> export
% Requires: Coring_data_imported.mat (contains T_all)
% Output:   Coring_data_processed.mat + Export/Coring_data_export.xlsx

clear; close all; clc

% --- locate project folder and Export folder ---
scriptPath = which("process_all_cores_only.m");
if isempty(scriptPath)
    scriptDir = pwd;
else
    scriptDir = fileparts(scriptPath);
end

exportFolder = fullfile(scriptDir, "Export");
if ~isfolder(exportFolder), mkdir(exportFolder); end

INFILE_IMPORTED   = fullfile(scriptDir, "Coring_data_imported.mat");
OUTFILE_PROCESSED = fullfile(scriptDir, "Coring_data_processed.mat");

if ~isfile(INFILE_IMPORTED)
    error("Imported MAT not found: %s", INFILE_IMPORTED);
end

% ========================= MAP =========================
MAP = struct();

% Common/meta
MAP.EventLabel     = "EventLabel";
MAP.GPS_Time       = "GPS_Time";
MAP.GPS_Lat        = "GPS_Lat";
MAP.GPS_Lon        = "GPS_Lon";
MAP.Station        = "Station";
MAP.StationNumber  = "StationNumber";
MAP.StationVisit   = "StationVisit";
MAP.IceAge         = "IceAge";
MAP.MeltPond       = "MeltPond";

% Core counters from import
MAP.CoreID_RHO     = "CoreID_RHO";
MAP.CoreID_T       = "CoreID_T";
MAP.CoreID_SALO18  = "CoreID_SALO18";

% RHO-specific
MAP.Salinity_raw = "salinity";
MAP.Salinity_used= "Salinity_used";
MAP.Tlab         = "Tlab";
MAP.T_ice        = "Temperature_interp";
MAP.Rho_si       = "rho_si";

MAP.Vb_export    = "vb_rho_export";
MAP.Vg_pr        = "vg_pr";
MAP.Vg           = "vg";

% ========================= 1) Load imported =========================
load(INFILE_IMPORTED, "T_all");
fprintf("Loaded imported data: %s\n", INFILE_IMPORTED);

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

% ========================= 2) Ensure required processed columns =========================
needCols = ["rho_si","rho_lab_kgm3", MAP.Salinity_used, "Temperature_interp", ...
            "vb_rho_export","vg_pr","vg", ...
            "Salinity_est_from_C"];  % internal/debug
for c = needCols
    cn = string(c);
    if ~ismember(cn, string(T_all_proc.rho.Properties.VariableNames))
        T_all_proc.rho.(cn) = NaN(height(T_all_proc.rho),1);
    end
end

% Salinity_used from raw salinity
if ismember(MAP.Salinity_raw, T_all_proc.rho.Properties.VariableNames)
    T_all_proc.rho.(MAP.Salinity_used) = toNum(T_all_proc.rho.(MAP.Salinity_raw));
else
    % fallback: old assumption col7
    T_all_proc.rho.(MAP.Salinity_used) = toNum(T_all_proc.rho{:,7});
end

% Global salinity fix (replace measured 0 by SP estimate from conductivity + temp)
SP_est_all = NaN(height(T_all_proc.rho),1);
hasCond = ismember("cond_mScm",    T_all_proc.rho.Properties.VariableNames);
hasTemp = ismember("tempC_SALO18", T_all_proc.rho.Properties.VariableNames);

if hasCond && hasTemp
    C_mScm = toNum(T_all_proc.rho.cond_mScm);
    t_C    = toNum(T_all_proc.rho.tempC_SALO18);

    okSP = ~isnan(C_mScm) & ~isnan(t_C);
    if any(okSP)
        try
            % SP_est_all(okSP) = gsw_SP_from_C(C_mScm(okSP), t_C(okSP), zeros(nnz(okSP),1));
            SP_est_all(okSP) = gsw_SP_from_C(C_mScm(okSP), 25, zeros(nnz(okSP),1));
        catch ME
            warning('gsw_SP_from_C failed (global salinity fix): %s', char(ME.message));
        end
    end

    Sused = T_all_proc.rho.(MAP.Salinity_used);
    idxRep = (Sused == 0) & ~isnan(SP_est_all);
    Sused(idxRep) = SP_est_all(idxRep);
    T_all_proc.rho.(MAP.Salinity_used) = Sused;
else
    warning('Global salinity fix skipped: need cond_mScm and tempC_SALO18 in T_all_proc.rho');
end

% Keep estimate
T_all_proc.rho.Salinity_est_from_C = SP_est_all;

% Lab density for ALL rows
if ismember("Density_gcm3", T_all_proc.rho.Properties.VariableNames)
    T_all_proc.rho.rho_lab_kgm3 = toNum(T_all_proc.rho.Density_gcm3) * 1000;
else
    % fallback: 3rd col (old behavior)
    T_all_proc.rho.rho_lab_kgm3 = toNum(T_all_proc.rho{:,3}) * 1000;
end

% ========================= 3) Match T and RHO cores by folder + descriptor + number =========================
Tmatch = table(strings(0,1), strings(0,1), strings(0,1), strings(0,1), ...
               'VariableNames', {'Station','Core','T_File','RHO_File'});

stCol = MAP.Station;
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

    if isempty(Tfiles), continue, end
    if isempty(RHOfiles)
        warning("No RHO files found in folder Station=%s / Core=%s", st, co);
        continue
    end

    [Tdesc, Tnum] = arrayfun(@parseCoreDescriptorAndNumber, Tfiles,   'UniformOutput', false);
    [Rdesc, Rnum] = arrayfun(@parseCoreDescriptorAndNumber, RHOfiles, 'UniformOutput', false);

    Tdesc = string(Tdesc);  Tnum = cell2mat(Tnum);
    Rdesc = string(Rdesc);  Rnum = cell2mat(Rnum);

    groups = unique(Tdesc, 'stable');
    for g = 1:numel(groups)
        grp = groups(g);

        iT = find(Tdesc == grp);
        iR = find(Rdesc == grp);

        if isempty(iR)
            warning("No RHO group match in folder Station=%s / Core=%s for group '%s'", st, co, grp);
            continue
        end

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

% ========================= 4) Compute rho_si + vb/vg and write back =========================
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

    depth_rho = mean(toNum(T_rho{:,1:2}), 2);

    depth_T = toNum(T_T{:,1});
    temp    = min(-0.1, toNum(T_T{:,2}));
    ok = ~isnan(depth_T) & ~isnan(temp);
    depth_T = depth_T(ok);
    temp    = temp(ok);

    if numel(depth_T) < 2
        warning('Too few temperature points — skipping');
        continue
    end

    depth_T_rescaled = depth_T * (max(depth_rho) / max(depth_T));
    T_interp = interp1(depth_T_rescaled, temp, depth_rho, 'linear', 'extrap');

    % --- Salinity used ---
    if ismember(MAP.Salinity_raw, T_rho.Properties.VariableNames)
        Srho = toNum(T_rho.(MAP.Salinity_raw));
    else
        SrhoVar = guessVar(T_rho, ["salinity","S","SALO"], "");
        if SrhoVar ~= ""
            Srho = toNum(T_rho.(SrhoVar));
        else
            Srho = NaN(height(T_rho),1);
        end
    end

    % SALO18 fallback if salinity missing in RHO
    if all(isnan(Srho))
        SALOcore = [];
        if isfield(T_all,'SALO18') && ~isempty(T_all.SALO18) && ismember(MAP.Station, T_all.SALO18.Properties.VariableNames)
            SALOcore = T_all.SALO18(T_all.SALO18.(MAP.Station) == Tmatch.Station(i), :);
        end
        if ~isempty(SALOcore)
            depth_SALO = mean(toNum(SALOcore{:,1:2}), 2);
            if ismember("Salinity", SALOcore.Properties.VariableNames)
                sal_SALO = toNum(SALOcore.Salinity);
            else
                sal_SALO = toNum(SALOcore{:,3});
            end
            Srho = interp1(depth_SALO, sal_SALO, depth_rho, 'linear', 'extrap');
        else
            warning('No SALO18 core found — using NaN for salinity');
        end
    end

    % Estimate salinity from conductivity + sample temp
    SP_est = NaN(height(T_rho),1);
    if ismember("cond_mScm", T_rho.Properties.VariableNames) && ismember("tempC_SALO18", T_rho.Properties.VariableNames)
        C_mScm = toNum(T_rho.cond_mScm);
        t_C    = toNum(T_rho.tempC_SALO18);
        okSP = ~isnan(C_mScm) & ~isnan(t_C);
        if any(okSP)
            try
                % SP_est(okSP) = gsw_SP_from_C(C_mScm(okSP), t_C(okSP), zeros(nnz(okSP),1));
                SP_est(okSP) = gsw_SP_from_C(C_mScm(okSP), 25, zeros(nnz(okSP),1));
            catch ME
                warning("gsw_SP_from_C failed for %s: %s", Tmatch.RHO_File(i), ME.message);
            end
        end
    end

    Srho_used = Srho;
    idxRep = (Srho_used == 0) & ~isnan(SP_est);
    Srho_used(idxRep) = SP_est(idxRep);

    % --- Lab temperature ---
    if ismember(MAP.Tlab, T_rho.Properties.VariableNames)
        T_lab = toNum(T_rho.(MAP.Tlab));
    else
        TlabVar = guessVar(T_rho, ["Tlab","lab"], "");
        if TlabVar ~= ""
            T_lab = toNum(T_rho.(TlabVar));
        else
            warning('No lab temperature column found — skipping');
            continue
        end
    end

    % Density (g/cm3)
    if ismember("Density_gcm3", T_rho.Properties.VariableNames)
        rho_meas = toNum(T_rho.Density_gcm3);
    else
        rho_meas = toNum(T_rho{:,3});
    end
    rho_lab_kgm3 = rho_meas * 1000;
    rho = rho_lab_kgm3;

    % ===================== CALCULATIONS =====================
    F1_pr_rho = -4.732 - 22.45*T_lab - 0.6397*T_lab.^2 - 0.01074*T_lab.^3;
    F2_pr_rho = 8.903e-2 - 1.763e-2*T_lab - 5.33e-4*T_lab.^2 - 8.801e-6*T_lab.^3;
    vb_pr_rho = rho .* Srho_used ./ F1_pr_rho;

    rhoi_pr   = 917 - 0.1403*T_lab;
    vg_pr     = max(0, 1 - rho .* (F1_pr_rho - rhoi_pr .* Srho_used/1000 .* F2_pr_rho) ./ (rhoi_pr .* F1_pr_rho));
    F3_pr     = rhoi_pr .* Srho_used/1000 ./ (F1_pr_rho - rhoi_pr .* Srho_used/1000 .* F2_pr_rho);

    T_insitu = T_interp;
    rhoi_rho = 917 - 0.1403*T_insitu;
    F1_rho = -4.732 - 22.45*T_insitu - 0.6397*T_insitu.^2 - 0.01074*T_insitu.^3;
    F2_rho = 8.903e-2 - 1.763e-2*T_insitu - 5.33e-4*T_insitu.^2 - 8.801e-6*T_insitu.^3;

    idx_LM = T_insitu > -2;
    F1_rho(idx_LM) = -4.1221e-2 - 18.407*T_insitu(idx_LM) + 0.58402*T_insitu(idx_LM).^2 + 0.21454*T_insitu(idx_LM).^3;
    F2_rho(idx_LM) = 9.0312e-2 - 0.016111*T_insitu(idx_LM) + 1.2291e-4*T_insitu(idx_LM).^2 + 1.3603e-4*T_insitu(idx_LM).^3;

    F3_rho = rhoi_rho .* Srho_used/1000 ./ (F1_rho - rhoi_rho .* Srho_used/1000 .* F2_rho);

    vb_rho_raw = vb_pr_rho .* F1_pr_rho ./ F1_rho / 1000;

    vb_rho_export = vb_rho_raw;
    vb_rho_export(vb_rho_export > 0.4 | vb_rho_export < 0) = NaN;

    vb_rho = vb_rho_raw;
    vb_rho(vb_rho > 0.4 | vb_rho < 0) = 0.4;

    R = (F3_pr.*F1_pr_rho) ./ (F3_rho.*F1_rho);
    idx0 = (Srho_used == 0);
    R(idx0) = rhoi_pr(idx0) ./ rhoi_rho(idx0);
    vg = max(0, 1 - (1 - vg_pr) .* (rhoi_rho./rhoi_pr) .* R);

    rho_si = (1 - vg) .* rhoi_rho .* F1_rho ./ (F1_rho - rhoi_rho .* Srho_used/1000 .* F2_rho);
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

    T_all_proc.rho.(MAP.Rho_si)(idxAll)        = rho_si;
    T_all_proc.rho.rho_lab_kgm3(idxAll)        = rho_lab_kgm3;

    Sprev = T_all_proc.rho.(MAP.Salinity_used)(idxAll);
    maskKeepPrev = (Srho_used == 0) & ~isnan(Sprev);
    Srho_used(maskKeepPrev) = Sprev(maskKeepPrev);
    T_all_proc.rho.(MAP.Salinity_used)(idxAll) = Srho_used;

    T_all_proc.rho.(MAP.T_ice)(idxAll)         = T_interp;
    T_all_proc.rho.(MAP.Vb_export)(idxAll)     = vb_rho_export;
    T_all_proc.rho.(MAP.Vg_pr)(idxAll)         = vg_pr;
    T_all_proc.rho.(MAP.Vg)(idxAll)            = vg;

    T_all_proc.rho.Salinity_est_from_C(idxAll) = SP_est;
end

% ========================= 5) FORMAT / REORDER / DROP columns =========================
rho = T_all_proc.rho;

if ismember(MAP.EventLabel, rho.Properties.VariableNames)
    rho.(MAP.EventLabel) = string(rho.(MAP.EventLabel));
end

rho = addTimeBest(rho, MAP.GPS_Time);
rho = addMeltPond01(rho, MAP.MeltPond);

% --- keep Comments column identical to previous export (blank if missing) ---
if ~ismember("Comments", rho.Properties.VariableNames)
    rho.Comments = repmat("", height(rho), 1);  % blank text column
else
    rho.Comments = string(rho.Comments);
end

% ---- FIX: force Depth1/Depth2 (avoid REFERENCE/zero vertical) ----
d1 = "Depth1";
d2 = "Depth2";
if ~ismember(d1, rho.Properties.VariableNames), d1 = rho.Properties.VariableNames{1}; end
if ~ismember(d2, rho.Properties.VariableNames), d2 = rho.Properties.VariableNames{2}; end

% Final salinity column
if ismember(MAP.Salinity_used, rho.Properties.VariableNames)
    rho.Salinity_final = rho.(MAP.Salinity_used);
elseif ismember(MAP.Salinity_raw, rho.Properties.VariableNames)
    rho.Salinity_final = rho.(MAP.Salinity_raw);
else
    gsal = guessVar(rho, ["salinity","S"], "");
    rho.Salinity_final = NaN(height(rho),1);
    if gsal ~= "", rho.Salinity_final = rho.(gsal); end
end
salFinal = "Salinity_final";

iceTh = guessVar(rho, ["thick","thickness","IceThickness"], "");
iceDr = guessVar(rho, ["draft","IceDraft"], "");
coLen = guessVar(rho, ["length","corelength","CoreLength"], "");

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
    "Comments"
];
T_all_proc.rho_out = keepAndOrder(rho, preferredRHO);

% --- T out ---
TT = T_all_proc.T;
TT = addTimeBest(TT, MAP.GPS_Time);
TT = addMeltPond01(TT, MAP.MeltPond);

tDepth = guessVar(TT, ["depth"], TT.Properties.VariableNames{1});
tTemp  = guessVar(TT, ["temp","temperature"], TT.Properties.VariableNames{2});

iceTh_T = guessVar(TT, ["thick","thickness","IceThickness"], "");
iceDr_T = guessVar(TT, ["draft","IceDraft"], "");
coLen_T = guessVar(TT, ["length","corelength","CoreLength"], "");

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
];
T_all_proc.T_out = keepAndOrder(TT, preferredT);

% --- SALO18 out ---
if isfield(T_all_proc,'SALO18') && ~isempty(T_all_proc.SALO18)
    S = T_all_proc.SALO18;
    S = addTimeBest(S, MAP.GPS_Time);
    S = addMeltPond01(S, MAP.MeltPond);

    sD1  = guessVar(S, ["depth1","Depth1"], S.Properties.VariableNames{1});
    sD2  = guessVar(S, ["depth2","Depth2"], S.Properties.VariableNames{2});
    sSal = guessVar(S, ["salinity","Salinity"], "");

    iceTh_S = guessVar(S, ["thick","thickness","IceThickness"], "");
    iceDr_S = guessVar(S, ["draft","IceDraft"], "");
    coLen_S = guessVar(S, ["length","corelength","CoreLength"], "");

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
    ];
    T_all_proc.SALO18_out = keepAndOrder(S, preferredS);
else
    warning('No SALO18 table found in T_all_proc');
end

% ========================= 6) Rename headers (PANGAEA convention) =========================
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
    "depth center",        "Depth, ice/snow"
    "IceThickness",        "Sea ice thickness"
    "Draft",               "Sea ice draft"

    "rho_lab_kgm3",        "Density, ice, technical"
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
    "Depth1",              "Depth, ice/snow, top/minimum"
    "Depth2",              "Depth, ice/snow, bottom/maximum"
    "Salinity",            "Sea ice salinity"
};

T_all_proc.rho_out = applyRenameMap(T_all_proc.rho_out, renameMap);
T_all_proc.T_out   = applyRenameMap(T_all_proc.T_out,   renameMap);
if isfield(T_all_proc,'SALO18_out')
    T_all_proc.SALO18_out = applyRenameMap(T_all_proc.SALO18_out, renameMap);
end

% ========================= 7) Rounding rules =========================
varsToRound_RHO = [
    "Density, ice, technical"
    "Density, ice"
    "Volume, brine"
    "Temperature, ice/snow"
];

for k = 1:numel(varsToRound_RHO)
    vn = char(varsToRound_RHO(k));
    if ismember(vn, T_all_proc.rho_out.Properties.VariableNames)
        x = T_all_proc.rho_out.(vn);
        if isnumeric(x), T_all_proc.rho_out.(vn) = round(x, 2); end
    end
end

gasVars = ["Volume, gas, technical","Volume, gas"];
for k = 1:numel(gasVars)
    vn = char(gasVars(k));
    if ismember(vn, T_all_proc.rho_out.Properties.VariableNames)
        x = T_all_proc.rho_out.(vn);
        if isnumeric(x), T_all_proc.rho_out.(vn) = round(x, 3); end
    end
end

salName = "Sea ice salinity";
if ismember(salName, T_all_proc.rho_out.Properties.VariableNames) && isnumeric(T_all_proc.rho_out.(salName))
    T_all_proc.rho_out.(salName) = round(T_all_proc.rho_out.(salName), 2);
end
if isfield(T_all_proc,'SALO18_out') && ismember(salName, T_all_proc.SALO18_out.Properties.VariableNames) ...
        && isnumeric(T_all_proc.SALO18_out.(salName))
    T_all_proc.SALO18_out.(salName) = round(T_all_proc.SALO18_out.(salName), 2);
end

% ========================= 8) Save + Export =========================
save(OUTFILE_PROCESSED, 'T_all_proc','Tmatch','rho_si_bulk','T_bulk');
fprintf('Saved processed + formatted data to %s\n', OUTFILE_PROCESSED);

outXlsx = fullfile(exportFolder, "Coring_data_export.xlsx");
writetable(T_all_proc.rho_out,    outXlsx, "Sheet", "RHO",    "WriteMode", "overwritesheet");
writetable(T_all_proc.T_out,      outXlsx, "Sheet", "T",      "WriteMode", "overwritesheet");
if isfield(T_all_proc,'SALO18_out')
    writetable(T_all_proc.SALO18_out, outXlsx, "Sheet", "SALO18", "WriteMode", "overwritesheet");
end
fprintf("Exported to %s (sheets: RHO, T, SALO18)\n", outXlsx);

%% ========================= FUNCTIONS =========================

function T_all = ensureNewCols(T_all, MAP)
    if isfield(T_all,'rho') && ~isempty(T_all.rho)
        if ~ismember(MAP.StationNumber, T_all.rho.Properties.VariableNames), T_all.rho.(MAP.StationNumber) = NaN(height(T_all.rho),1); end
        if ~ismember(MAP.StationVisit,  T_all.rho.Properties.VariableNames), T_all.rho.(MAP.StationVisit)  = repmat("", height(T_all.rho),1); end
        T_all.rho.(MAP.StationVisit) = string(T_all.rho.(MAP.StationVisit));
        if ~ismember(MAP.CoreID_RHO, T_all.rho.Properties.VariableNames),     T_all.rho.(MAP.CoreID_RHO)     = NaN(height(T_all.rho),1); end

        if ~ismember("cond_mScm", T_all.rho.Properties.VariableNames),         T_all.rho.cond_mScm = NaN(height(T_all.rho),1); end
        if ~ismember("tempC_SALO18", T_all.rho.Properties.VariableNames),      T_all.rho.tempC_SALO18 = NaN(height(T_all.rho),1); end
        if ~ismember(MAP.Salinity_used, T_all.rho.Properties.VariableNames),   T_all.rho.(MAP.Salinity_used) = NaN(height(T_all.rho),1); end
        if ~ismember("Salinity_est_from_C", T_all.rho.Properties.VariableNames), T_all.rho.Salinity_est_from_C = NaN(height(T_all.rho),1); end
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
    f = string(fname);
    f = regexprep(f, "\.xlsx$", "", "ignorecase");
    f = regexprep(f, "-(T|RHO(-TXT)?)$", "", "ignorecase");

    tokNum = regexp(f, "SI_corer_\d+cm-(\d+)", "tokens", "once");
    if isempty(tokNum), num = NaN; else, num = str2double(tokNum{1}); end

    tokDesc = regexp(f, "SI_corer_\d+cm-\d+-(.+)$", "tokens", "once");
    if isempty(tokDesc), desc = ""; else, desc = string(tokDesc{1}); end
end

function x = toNum(x)
    if iscell(x) || isstring(x)
        x = str2double(string(x));
    end
    x = double(x);
end