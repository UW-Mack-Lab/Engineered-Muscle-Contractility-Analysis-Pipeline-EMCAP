# -*- coding: utf-8 -*-
"""
Created on Mon May 19 2025

@author: pb92
@version: 1.4.2 - dtype-safe metric column initialization

CHANGELOG:
v1.4.2 - FIXED: pandas FutureWarning for incompatible dtype assignment in kinetics columns.
         - Replaced heuristic column init that set any '*Peak*' column to int(0).
         - Added explicit flag-column list so metric columns initialize as float (np.nan).
         - Preserves all computational logic and outputs; only dtype initialization behavior changed.
v1.4.1 - PERFORMANCE: Parallel per-file processing + fast no-style Excel export.
v1.4.0 - Functionalized script with run_twitch_analysis(...) entrypoint.
"""

import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, simpledialog
from openpyxl import Workbook
from tqdm import tqdm
import uuid
from datetime import datetime
import sys
from concurrent.futures import as_completed

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from file_pool import FilePool, raise_if_crashed, reserve_output_paths, tag_suffix  # noqa: E402
from filenames import protocol_from_filename  # noqa: E402

# ==========================================
#        CONFIGURATION SECTION
# ==========================================

# Recording Settings
MAX_RECORDING_TIME = 160.0
NUM_STIMS = 3

# Stimulation Protocol Settings (Standard Twitch)
INITIAL_DELAY = 17.0
STIM_DURATION = 0.2
REST_DURATION = 45.0

# Signal Processing Settings
SMOOTHING_WINDOW = 3
BASELINE_WINDOW = 1.0
PEAK_SEARCH_WINDOW = 0.3

# Twitch Detection Thresholds
MINIMUM_TWITCH_THRESHOLD_UN = 5.0
FAST_PEAK_WARNING_MS = 80

# ==========================================
#        END CONFIGURATION SECTION
# ==========================================

T0_PAIRS = []
DEFAULT_MAX_WORKERS = max(1, os.cpu_count() or 1)


def select_files(title, file_types):
    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(title=title, filetypes=file_types)
    return files


def select_directory(title):
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title=title)
    return directory


def _timestamp_suffix():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def make_timestamped_path_if_exists(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    candidate = f"{base}_{_timestamp_suffix()}{ext}"
    counter = 1
    while os.path.exists(candidate):
        candidate = f"{base}_{_timestamp_suffix()}_{counter}{ext}"
        counter += 1
    return candidate


def _load_continuous_waveforms_from_file(input_file):
    try:
        if input_file.lower().endswith('.csv'):
            raw_df = pd.read_csv(input_file)
            if 'Time (s)' in raw_df.columns:
                raw_df.rename(columns={'Time (s)': 'Time (seconds)'}, inplace=True)

            continuous_waveforms_df = pd.DataFrame()
            if 'Time (seconds)' in raw_df.columns:
                continuous_waveforms_df['Time (seconds)'] = raw_df['Time (seconds)']

            for col in raw_df.columns:
                if col in ['Time (seconds)', 'Unnamed: 0']:
                    continue
                new_col_name = f"{col} - Active Twitch Force (μN)"
                continuous_waveforms_df[new_col_name] = raw_df[col]
        else:
            xls = pd.ExcelFile(input_file)
            if 'continuous-waveforms' not in xls.sheet_names:
                return None
            continuous_waveforms_df = pd.read_excel(xls, 'continuous-waveforms')

        if 'Time (seconds)' in continuous_waveforms_df.columns:
            continuous_waveforms_df['Time (seconds)'] = (continuous_waveforms_df['Time (seconds)'] * 100).round() / 100

        return continuous_waveforms_df
    except Exception as e:
        print(f"Error reading input file {input_file}: {e}")
        return None


def _process_single_file(input_file, plate_map, max_recording_time):
    continuous_waveforms_df = _load_continuous_waveforms_from_file(input_file)
    if continuous_waveforms_df is None:
        return None

    filename = os.path.basename(input_file)
    file_info = parse_filename(filename)
    if not file_info:
        return None

    day, protocol, plate_name = file_info

    if protocol and '.' in protocol:
        protocol = protocol.split('.')[0]

    continuous_df_kinetics = process_continuous_waveforms(continuous_waveforms_df, plate_map, plate_name, max_recording_time)
    if continuous_df_kinetics.empty:
        return None

    return day, protocol, continuous_df_kinetics


def df_to_excel_sheet(wb, sheet_name, df, start_cell='A1'):
    del start_cell
    df_formatted = df.copy()
    for col in df_formatted.columns:
        if 'Time' in col and col != 'Time (seconds)' and df_formatted[col].dtype in ['float64', 'float32']:
            if any(kinetic_term in col for kinetic_term in ['R10 Time', 'R50 Time', 'R80 Time', 'R90 Time', 'TT50P Time', 'TTP90 Time']):
                df_formatted[col] = df_formatted[col].round(3)
            else:
                df_formatted[col] = df_formatted[col].round(2)
        elif 'Force' in col or 'Threshold' in col:
            df_formatted[col] = df_formatted[col].round(2)
        elif col == 'Relative Time (ms)':
            df_formatted[col] = df_formatted[col].round(0)

    safe_sheet_name = sheet_name.replace('/', '_').replace('\\', '_')[:31]
    ws = wb.create_sheet(title=safe_sheet_name)
    ws.append(list(df_formatted.columns))
    for row in df_formatted.itertuples(index=False, name=None):
        ws.append(row)


def parse_filename(filename):
    parts = os.path.basename(filename).split('_')
    day = next((part for part in parts if part.lower().startswith('d') and part[1:].isdigit()), None)

    if day:
        day_idx = parts.index(day)
        plate_name = parts[day_idx + 1] if len(parts) > day_idx + 1 else "UnknownPlate"
        protocol = protocol_from_filename(filename) or "UnknownProtocol"
        return day, protocol, plate_name
    return None


def smooth_data(df, column, window_size=3):
    try:
        df[f'{column}_Smoothed'] = df.groupby('Well')[column].rolling(window=window_size, center=True, min_periods=1).mean().reset_index(0, drop=True)
        return df
    except Exception as e:
        print(f"Error in smooth_data function: {e}")
        df[f'{column}_Smoothed'] = df[column]
        return df


def calculate_stimulation_times(num_stims, max_time):
    global T0_PAIRS

    print("\n=== Twitch Protocol Configuration ===")
    print(f"- Initial Delay: {INITIAL_DELAY}s")
    print(f"- Stim Duration: {STIM_DURATION}s")
    print(f"- Rest Duration: {REST_DURATION}s")
    print(f"- Baseline Window: {BASELINE_WINDOW}s")
    print(f"- Peak Search Window: {PEAK_SEARCH_WINDOW}s")
    print(f"- Target Stims: {num_stims}")

    T0_PAIRS = []
    for i in range(num_stims):
        t0 = INITIAL_DELAY + i * (STIM_DURATION + REST_DURATION)
        t0 = round(t0 * 100) / 100
        if t0 > max_time:
            print(f"Skipping stim {i+1} at time {t0} (beyond max time {max_time})")
            break
        T0_PAIRS.append(t0)
    print(f"Calculated {len(T0_PAIRS)} stimulation times.")
    return T0_PAIRS


def find_peaks_and_relaxation_times(df, t0_pairs, min_twitch_threshold, fast_peak_warning_s):
    df = df.copy()
    cols_to_add = [
        'Stim Peak', 'True Peak Force (μN)', 'BK corrected force (μN)', 'Normalized Force (% of T1)', 'Fast Peak Warning',
        'R10 Reached', 'R10 Time', 'R10 Threshold Required (μN)',
        'R50 Reached', 'R50 Time', 'R50 Threshold Required (μN)',
        'R80 Reached', 'R80 Time', 'R80 Threshold Required (μN)',
        'R90 Reached', 'R90 Time', 'R90 Threshold Required (μN)',
        'TT50P Reached', 'TT50P Time', 'TT50P Threshold Required (μN)',
        'TTP90 Reached', 'TTP90 Time', 'TTP90 Threshold Required (μN)',
        'Baseline Force (μN)', 'Baseline Drift From Start (μN)', 'Baseline Drift Rate (μN/s)', 'Stim Number'
    ]

    # Explicit flag columns: keep as 0/1. All metric columns initialize as float (np.nan)
    # to avoid FutureWarning when writing float values into previously-int columns.
    flag_cols = {
        'Stim Peak', 'Fast Peak Warning',
        'R10 Reached', 'R50 Reached', 'R80 Reached', 'R90 Reached',
        'TT50P Reached', 'TTP90 Reached'
    }
    for col in cols_to_add:
        if col in flag_cols:
            df[col] = 0
        else:
            df[col] = np.nan

    for well in df['Well'].unique():
        well_data = df[df['Well'] == well].copy()
        first_baseline = None
        first_peak_time_abs = None
        first_valid_bk_force = None
        stim_count = 0

        for t0 in t0_pairs:
            baseline_window_df = well_data[(well_data['Time (seconds)'] >= t0 - BASELINE_WINDOW) &
                                           (well_data['Time (seconds)'] < t0)]
            if baseline_window_df.empty:
                continue

            baseline = baseline_window_df['Active Twitch Force (μN)_Smoothed'].mean()

            peak_search_df = well_data[(well_data['Time (seconds)'] >= t0) &
                                       (well_data['Time (seconds)'] <= t0 + PEAK_SEARCH_WINDOW)]
            if peak_search_df.empty:
                continue

            stim_peak_index = peak_search_df['Active Twitch Force (μN)_Smoothed'].idxmax()
            stim_peak_time_abs = well_data.loc[stim_peak_index, 'Time (seconds)']

            df.loc[stim_peak_index, 'Stim Peak'] = 1
            stim_count += 1
            df.loc[stim_peak_index, 'Stim Number'] = stim_count

            true_peak_force = well_data.loc[stim_peak_index, 'Active Twitch Force (μN)_Smoothed']
            df.loc[stim_peak_index, 'True Peak Force (μN)'] = true_peak_force
            df.loc[stim_peak_index, 'Baseline Force (μN)'] = baseline

            if first_baseline is None:
                first_baseline = baseline
                first_peak_time_abs = stim_peak_time_abs

            drift = baseline - first_baseline
            df.loc[stim_peak_index, 'Baseline Drift From Start (μN)'] = drift

            if first_peak_time_abs is not None and stim_peak_time_abs > first_peak_time_abs:
                time_diff_for_drift = stim_peak_time_abs - first_peak_time_abs
                if time_diff_for_drift > 0:
                    drift_rate = drift / time_diff_for_drift
                    df.loc[stim_peak_index, 'Baseline Drift Rate (μN/s)'] = drift_rate

            bk_corrected_force = true_peak_force - baseline

            time_to_peak_from_t0 = stim_peak_time_abs - t0
            if time_to_peak_from_t0 < fast_peak_warning_s:
                df.loc[stim_peak_index, 'Fast Peak Warning'] = 1

            if bk_corrected_force < min_twitch_threshold:
                df.loc[stim_peak_index, 'BK corrected force (μN)'] = np.nan
                if first_valid_bk_force is None and stim_count == 1:
                    first_valid_bk_force = np.nan
                if pd.notna(first_valid_bk_force) and first_valid_bk_force > 0:
                    normalized_force = (bk_corrected_force / first_valid_bk_force) * 100
                    df.loc[stim_peak_index, 'Normalized Force (% of T1)'] = normalized_force
                elif first_valid_bk_force is None and stim_count > 1:
                    pass
                else:
                    df.loc[stim_peak_index, 'Normalized Force (% of T1)'] = np.nan
                continue

            df.loc[stim_peak_index, 'BK corrected force (μN)'] = bk_corrected_force

            if first_valid_bk_force is None:
                first_valid_bk_force = bk_corrected_force

            if pd.notna(first_valid_bk_force) and first_valid_bk_force > 0:
                normalized_force = (bk_corrected_force / first_valid_bk_force) * 100
                df.loc[stim_peak_index, 'Normalized Force (% of T1)'] = normalized_force
            else:
                df.loc[stim_peak_index, 'Normalized Force (% of T1)'] = np.nan

            r10_threshold = (bk_corrected_force * 0.9) + baseline
            r50_threshold = (bk_corrected_force * 0.5) + baseline
            r80_threshold = (bk_corrected_force * 0.2) + baseline
            r90_threshold = (bk_corrected_force * 0.1) + baseline

            df.loc[stim_peak_index, 'R10 Threshold Required (μN)'] = r10_threshold
            df.loc[stim_peak_index, 'R50 Threshold Required (μN)'] = r50_threshold
            df.loc[stim_peak_index, 'R80 Threshold Required (μN)'] = r80_threshold
            df.loc[stim_peak_index, 'R90 Threshold Required (μN)'] = r90_threshold

            relaxation_window_df = well_data[well_data['Time (seconds)'] > stim_peak_time_abs].copy()

            for r_name, r_threshold in [('R10', r10_threshold), ('R50', r50_threshold),
                                        ('R80', r80_threshold), ('R90', r90_threshold)]:
                consecutive_count = 0
                r_start_idx = None
                for idx_loop, row_loop in relaxation_window_df.iterrows():
                    if row_loop['Active Twitch Force (μN)_Smoothed'] <= (r_threshold * 1.02):
                        if consecutive_count == 0:
                            r_start_idx = idx_loop
                        consecutive_count += 1
                        if consecutive_count == 3:
                            time_diff_relax = well_data.loc[r_start_idx, 'Time (seconds)'] - stim_peak_time_abs
                            time_diff_relax = round(time_diff_relax * 1000) / 1000
                            df.loc[stim_peak_index, f'{r_name} Time'] = time_diff_relax
                            df.loc[stim_peak_index, f'{r_name} Reached'] = 1
                            break
                    else:
                        consecutive_count = max(0, consecutive_count - 1)

            post_peak_data = well_data[well_data['Time (seconds)'] > stim_peak_time_abs].copy()

            if not post_peak_data.empty:
                min_force_idx = post_peak_data['Active Twitch Force (μN)_Smoothed'].idxmin()
                min_force = post_peak_data.loc[min_force_idx, 'Active Twitch Force (μN)_Smoothed']

                tt50p_threshold = min_force + 0.5 * (true_peak_force - min_force)
                ttp90_threshold = min_force + 0.1 * (true_peak_force - min_force)

                df.loc[stim_peak_index, 'TT50P Threshold Required (μN)'] = tt50p_threshold
                df.loc[stim_peak_index, 'TTP90 Threshold Required (μN)'] = ttp90_threshold

                for rel_name, rel_threshold in [('TT50P', tt50p_threshold), ('TTP90', ttp90_threshold)]:
                    consecutive_count_rel = 0
                    rel_start_idx = None

                    for idx_loop, row_loop in post_peak_data.iterrows():
                        if row_loop['Active Twitch Force (μN)_Smoothed'] <= (rel_threshold * 1.02):
                            if consecutive_count_rel == 0:
                                rel_start_idx = idx_loop
                            consecutive_count_rel += 1
                            if consecutive_count_rel == 3:
                                time_diff_rel = well_data.loc[rel_start_idx, 'Time (seconds)'] - stim_peak_time_abs
                                time_diff_rel = round(time_diff_rel * 1000) / 1000
                                df.loc[stim_peak_index, f'{rel_name} Time'] = time_diff_rel
                                df.loc[stim_peak_index, f'{rel_name} Reached'] = 1
                                break
                        else:
                            consecutive_count_rel = max(0, consecutive_count_rel - 1)

    return df


def process_continuous_waveforms(continuous_waveforms_df, plate_map, plate_name, max_recording_time,
                                 min_twitch_threshold=MINIMUM_TWITCH_THRESHOLD_UN,
                                 fast_peak_warning_ms=FAST_PEAK_WARNING_MS):
    if continuous_waveforms_df is None or continuous_waveforms_df.empty:
        return pd.DataFrame()

    if 'Time (seconds)' not in continuous_waveforms_df.columns:
        print("Warning: 'Time (seconds)' not found in continuous-waveforms. Skipping file.")
        return pd.DataFrame()

    waveform_df = continuous_waveforms_df[continuous_waveforms_df['Time (seconds)'] <= max_recording_time].copy()

    if waveform_df.empty:
        return pd.DataFrame()

    force_cols = [c for c in waveform_df.columns if 'Active Twitch Force' in c and c != 'Time (seconds)']
    if not force_cols:
        return pd.DataFrame()

    long_df = waveform_df.melt(
        id_vars=['Time (seconds)'],
        value_vars=force_cols,
        var_name='Well_Force_Col',
        value_name='Active Twitch Force (μN)'
    )

    long_df['Well'] = long_df['Well_Force_Col'].str.extract(r'([A-H]\d{1,2})', expand=False)
    long_df = long_df.dropna(subset=['Well']).copy()

    condition_map = {}
    for row_idx in plate_map.index:
        for col_idx in plate_map.columns:
            condition_map[f"{row_idx}{col_idx}"] = plate_map.loc[row_idx, col_idx]

    long_df['Condition'] = long_df['Well'].map(condition_map).fillna('Unknown')
    long_df['Plate Name'] = plate_name

    long_df = smooth_data(long_df, 'Active Twitch Force (μN)', window_size=SMOOTHING_WINDOW)

    fast_peak_warning_s = fast_peak_warning_ms / 1000.0
    metrics_df = find_peaks_and_relaxation_times(
        long_df,
        t0_pairs=T0_PAIRS,
        min_twitch_threshold=min_twitch_threshold,
        fast_peak_warning_s=fast_peak_warning_s,
    )

    return metrics_df


def calculate_summary_stats(group_df):
    summary_data = []

    peak_data = group_df[group_df['Stim Peak'] == 1].copy()
    if peak_data.empty:
        return pd.DataFrame()

    grouped = peak_data.groupby(['Condition', 'Stim Number'], dropna=False)

    for (condition, stim_num), g in grouped:
        if pd.isna(stim_num):
            continue

        row = {
            'Condition': condition,
            'Stim Number': int(stim_num),
            'Count': len(g),

            'True Peak Force Mean (μN)': g['True Peak Force (μN)'].mean(),
            'True Peak Force StdDev (μN)': g['True Peak Force (μN)'].std(),

            'BK corrected force Mean (μN)': g['BK corrected force (μN)'].mean(),
            'BK corrected force StdDev (μN)': g['BK corrected force (μN)'].std(),

            'Normalized Force Mean (% of T1)': g['Normalized Force (% of T1)'].mean(),
            'Normalized Force StdDev (% of T1)': g['Normalized Force (% of T1)'].std(),

            'R10 Time Mean (s)': g['R10 Time'].mean(),
            'R10 Time StdDev (s)': g['R10 Time'].std(),

            'R50 Time Mean (s)': g['R50 Time'].mean(),
            'R50 Time StdDev (s)': g['R50 Time'].std(),

            'R80 Time Mean (s)': g['R80 Time'].mean(),
            'R80 Time StdDev (s)': g['R80 Time'].std(),

            'R90 Time Mean (s)': g['R90 Time'].mean(),
            'R90 Time StdDev (s)': g['R90 Time'].std(),

            'TT50P Time Mean (s)': g['TT50P Time'].mean(),
            'TT50P Time StdDev (s)': g['TT50P Time'].std(),

            'TTP90 Time Mean (s)': g['TTP90 Time'].mean(),
            'TTP90 Time StdDev (s)': g['TTP90 Time'].std(),

            'Baseline Force Mean (μN)': g['Baseline Force (μN)'].mean(),
            'Baseline Force StdDev (μN)': g['Baseline Force (μN)'].std(),
            'Baseline Drift Mean (μN)': g['Baseline Drift From Start (μN)'].mean(),
            'Baseline Drift StdDev (μN)': g['Baseline Drift From Start (μN)'].std(),
        }
        summary_data.append(row)

    summary_df = pd.DataFrame(summary_data)
    if not summary_df.empty:
        summary_df = summary_df.sort_values(by=['Condition', 'Stim Number'])
    return summary_df


def calculate_well_averages(combined_kinetics_df):
    peak_data = combined_kinetics_df[combined_kinetics_df['Stim Peak'] == 1].copy()
    if peak_data.empty:
        return pd.DataFrame()

    agg_dict = {
        'True Peak Force (μN)': ['mean', 'std'],
        'BK corrected force (μN)': ['mean', 'std'],
        'Normalized Force (% of T1)': ['mean', 'std'],
        'R10 Time': ['mean', 'std'],
        'R50 Time': ['mean', 'std'],
        'R80 Time': ['mean', 'std'],
        'R90 Time': ['mean', 'std'],
        'TT50P Time': ['mean', 'std'],
        'TTP90 Time': ['mean', 'std'],
        'Baseline Force (μN)': ['mean', 'std'],
        'Baseline Drift From Start (μN)': ['mean', 'std'],
    }

    grouped = peak_data.groupby(['Plate Name', 'Well', 'Condition']).agg(agg_dict)
    grouped.columns = [
        f"{metric} {'Mean' if stat == 'mean' else 'StdDev'}"
        for metric, stat in grouped.columns
    ]
    grouped = grouped.reset_index()
    return grouped


def pick_representative_traces(combined_df):
    reps = []
    peak_df = combined_df[combined_df['Stim Peak'] == 1].copy()
    if peak_df.empty:
        return pd.DataFrame()

    per_well = peak_df.groupby(['Condition', 'Plate Name', 'Well'], as_index=False)['BK corrected force (μN)'].mean()

    for condition, cdf in per_well.groupby('Condition'):
        cdf = cdf.dropna(subset=['BK corrected force (μN)'])
        if cdf.empty:
            continue
        target = cdf['BK corrected force (μN)'].median()
        cdf['distance'] = (cdf['BK corrected force (μN)'] - target).abs()
        best = cdf.sort_values('distance').iloc[0]

        rep_trace = combined_df[
            (combined_df['Condition'] == condition)
            & (combined_df['Plate Name'] == best['Plate Name'])
            & (combined_df['Well'] == best['Well'])
        ].copy()

        rep_trace['Representative Condition'] = condition
        reps.append(rep_trace)

    if reps:
        return pd.concat(reps, ignore_index=True)
    return pd.DataFrame()


def run_twitch_analysis(input_files, plate_map_path, output_dir, max_workers=None, file_tag=""):
    """file_tag (e.g. a plate name) is appended to every output file name."""
    if not input_files:
        raise ValueError("No input files provided.")
    if not plate_map_path or not os.path.isfile(plate_map_path):
        raise FileNotFoundError(f"Plate map not found: {plate_map_path}")
    if not output_dir:
        raise ValueError("No output directory provided.")

    os.makedirs(output_dir, exist_ok=True)

    print("\n=== Twitch Analysis Settings (from CONFIGURATION) ===")
    print(f"MAX_RECORDING_TIME: {MAX_RECORDING_TIME}")
    print(f"NUM_STIMS: {NUM_STIMS}")
    print(f"MINIMUM_TWITCH_THRESHOLD_UN: {MINIMUM_TWITCH_THRESHOLD_UN}")
    print(f"FAST_PEAK_WARNING_MS: {FAST_PEAK_WARNING_MS}")

    plate_map = pd.read_csv(plate_map_path, index_col=0)
    plate_map.index = plate_map.index.astype(str)
    plate_map.columns = plate_map.columns.astype(str)

    calculate_stimulation_times(NUM_STIMS, MAX_RECORDING_TIME)

    all_kinetics_data = {}

    effective_workers = max_workers if max_workers is not None else DEFAULT_MAX_WORKERS
    effective_workers = max(1, min(int(effective_workers), len(input_files)))

    print(f"\nProcessing {len(input_files)} file(s) with max_workers={effective_workers}...")

    with FilePool(__file__, {"T0_PAIRS": T0_PAIRS}, effective_workers) as pool:
        futures = {
            pool.submit("_process_single_file", input_file, plate_map, MAX_RECORDING_TIME): input_file
            for input_file in input_files
        }

        for future in tqdm(as_completed(futures), total=len(futures), desc="Files"):
            input_file = futures[future]
            try:
                result = future.result()
            except Exception as e:
                raise_if_crashed(e)
                print(f"Error processing {os.path.basename(input_file)}: {e}")
                continue

            if result is None:
                continue

            day, protocol, continuous_df_kinetics = result

            if day not in all_kinetics_data:
                all_kinetics_data[day] = {}
            if protocol not in all_kinetics_data[day]:
                all_kinetics_data[day][protocol] = []

            all_kinetics_data[day][protocol].append(continuous_df_kinetics)

    if not all_kinetics_data:
        print("No valid data processed.")
        return None

    groups = [
        (day, protocol, kinetics_list)
        for day, protocols in all_kinetics_data.items()
        for protocol, kinetics_list in protocols.items()
        if kinetics_list
    ]
    # Names are picked up front so parallel writers can't overwrite each other.
    output_paths = reserve_output_paths([
        os.path.join(output_dir, f"{day}_{protocol}_Twitch_Analyzed{tag_suffix(file_tag)}.xlsx") for day, protocol, _ in groups
    ])
    with FilePool(__file__, {"T0_PAIRS": T0_PAIRS}, min(effective_workers, len(groups)) or 1) as pool:
        group_futures = [
            pool.submit("_write_group_output", day, protocol, kinetics_list, output_dir, path)
            for (day, protocol, kinetics_list), path in zip(groups, output_paths)
        ]
        for future in group_futures:
            future.result()

    print("\nTwitch analysis complete.")
    return True


def _write_group_output(day, protocol, kinetics_list, output_dir, output_path):
    """Build and save the workbook for one day/protocol."""
    combined_kinetics_df = pd.concat(kinetics_list, ignore_index=True)

    summary_df = calculate_summary_stats(combined_kinetics_df)
    well_average_df = calculate_well_averages(combined_kinetics_df)
    representative_df = pick_representative_traces(combined_kinetics_df)

    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    desired_cols = [
        'Time (seconds)', 'Well', 'Condition', 'Plate Name', 'Active Twitch Force (μN)',
        'Active Twitch Force (μN)_Smoothed', 'Stim Peak', 'Stim Number',
        'True Peak Force (μN)', 'BK corrected force (μN)', 'Normalized Force (% of T1)',
        'Fast Peak Warning',
        'R10 Time', 'R10 Threshold Required (μN)', 'R10 Reached',
        'R50 Time', 'R50 Threshold Required (μN)', 'R50 Reached',
        'R80 Time', 'R80 Threshold Required (μN)', 'R80 Reached',
        'R90 Time', 'R90 Threshold Required (μN)', 'R90 Reached',
        'TT50P Time', 'TT50P Threshold Required (μN)', 'TT50P Reached',
        'TTP90 Time', 'TTP90 Threshold Required (μN)', 'TTP90 Reached',
        'Baseline Force (μN)', 'Baseline Drift From Start (μN)', 'Baseline Drift Rate (μN/s)'
    ]

    existing_cols = [c for c in desired_cols if c in combined_kinetics_df.columns]
    kinetics_for_export = combined_kinetics_df[existing_cols].copy()

    df_to_excel_sheet(wb, f"{day}_{protocol}_kinetics", kinetics_for_export)

    if not summary_df.empty:
        df_to_excel_sheet(wb, f"{day}_{protocol}_summary", summary_df)

    if not well_average_df.empty:
        df_to_excel_sheet(wb, f"{day}_{protocol}_well_avg", well_average_df)

    if not representative_df.empty:
        rep_cols = [
            'Time (seconds)', 'Well', 'Condition', 'Representative Condition', 'Plate Name',
            'Active Twitch Force (μN)', 'Active Twitch Force (μN)_Smoothed',
            'Stim Peak', 'Stim Number', 'BK corrected force (μN)'
        ]
        rep_cols = [c for c in rep_cols if c in representative_df.columns]
        df_to_excel_sheet(wb, f"{day}_{protocol}_representative", representative_df[rep_cols])

    wb.save(output_path)
    print(f"Saved output: {output_path}")


def main():
    print("Select one or more Twitch data files (.xlsx or .csv)")
    input_files = select_files(
        "Select Twitch Data Files",
        [("Data files", "*.xlsx *.csv"), ("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )

    if not input_files:
        print("No files selected.")
        return

    print("Select plate map CSV file")
    plate_map_file = select_files("Select Plate Map CSV", [("CSV files", "*.csv")])
    if not plate_map_file:
        print("No plate map file selected.")
        return
    plate_map_path = plate_map_file[0]

    output_dir = select_directory("Select Output Folder")
    if not output_dir:
        print("No output folder selected.")
        return

    max_workers = None
    try:
        root = tk.Tk()
        root.withdraw()
        user_val = simpledialog.askinteger(
            "Parallel Workers",
            f"Enter max file workers (1-{os.cpu_count() or 8}). Leave blank for default ({DEFAULT_MAX_WORKERS}).",
            minvalue=1,
            maxvalue=max(1, os.cpu_count() or 8),
        )
        if user_val is not None:
            max_workers = int(user_val)
    except Exception:
        max_workers = None

    run_twitch_analysis(
        input_files=input_files,
        plate_map_path=plate_map_path,
        output_dir=output_dir,
        max_workers=max_workers,
    )


if __name__ == '__main__':
    main()
