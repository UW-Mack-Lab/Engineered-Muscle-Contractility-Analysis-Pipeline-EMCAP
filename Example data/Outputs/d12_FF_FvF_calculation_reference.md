# Force vs Frequency (FvF) Analysis - Calculation Reference (d12_FF)

This document explains the calculation methods used in the FvF waveform analysis script, with example calculations from Well D4.

## Example Calculation for Well D4 (20uM)
### Stimulation: #1 (1Hz)

### FvF Protocol Parameters
- T0 (stimulation start): 21.00 seconds
- T1 (stimulation end): 23.00 seconds
- Stimulation frequency: 1Hz
- Detected Peak Time: 23.00 seconds
- Kinetics Peak Time: 22.94 seconds (Δ=-0.06s)

### Force Measurements
- Baseline Force: 5.901050002042516 uN
- Kinetics Peak Force: 7.820155587978661 uN
- Max Tetanic Force (raw): 82.2954680188559 uN
- Max Tetanic-BK Force: 76.39441801681339 uN
- Kinetics Peak-BK Force: 1.919105585936145 uN

### Threshold Calculations

### Time Measurements

## Force vs Frequency (FvF) Protocol

### Stimulation Schedule:
The FvF protocol consists of 14 stimulations with increasing frequencies:

  Stim  1:   1Hz
  Stim  2:   2Hz
  Stim  3:   3Hz
  Stim  4:   5Hz
  Stim  5:  10Hz
  Stim  6:  15Hz
  Stim  7:  20Hz
  Stim  8:  30Hz
  Stim  9:  40Hz
  Stim 10:  60Hz
  Stim 11:  80Hz
  Stim 12: 100Hz

### Timing Parameters:
- Initial delay: 21.0 seconds
- Each stimulation duration: 2.0 seconds
- Inter-stimulation delay: 16.0 seconds
- Baseline window: 1.0 seconds before T0

## General Calculation Methods

1. **Baseline Force**: Average force over baseline window before each T0
2. **Kinetics Peak Force**: Maximum force value in ±100ms window around T1 (for kinetics calculations)
3. **Max Tetanic Force**: Maximum force in T0 to T1+200ms window (captures late peaks after stimulation)
4. **Max Tetanic-BK**: Max Tetanic Force MINUS baseline (BK corrected for F-F analysis)
5. **Kinetics Peak-BK**: Kinetics Peak Force - Baseline (used for relaxation thresholds)
6. **FF Normalized (%)**: Each stimulation normalized as % of maximum Max Tetanic-BK within that well across all Hz
7. **Thresholds** (based on kinetics peak):
    - R10: Baseline + (Kinetics Peak-BK × 0.9)
    - R50: Baseline + (Kinetics Peak-BK × 0.5)
    - R80: Baseline + (Kinetics Peak-BK × 0.2)
    - R90: Baseline + (Kinetics Peak-BK × 0.1)
8. **Relaxation Times**: Time from T1 until force drops below threshold
9. **TT50P**: Time from T0 until force rises above 50% of peak force
10. **TTP90**: Time from T0 until force rises above 90% of peak force
11. **TTMFP (s)**: Time from T0 to the absolute maximum force within a 2.5s window.
12. **Time Over 90% (%)**: Percentage of stimulation time (T0-T1) spent above 90% of the max force found for TTMFP. Measures fatigue resistance.

## Key Changes in v1.0.1

### Max Tetanic Force Search Window Extended:
- **Old**: T0 to T1 search window (missed late peaks after stimulation)
- **New**: T0 to T1+200ms search window (captures late peaks that occur after stimulation ends)
- **Benefit**: Ensures Max Tetanic Force >= Kinetics Peak Force (logical consistency)

### Dual Max Tetanic Force Columns:
- **Max Tetanic Force (uN)**: Raw maximum force in T0 to T1+200ms window
- **Max tetanic-BK (uN)**: Same as above but baseline corrected (Max Tetanic - Baseline)
- **Column Order**: Kinetics Peak → Max Tetanic (raw) → Baseline → Kinetics Peak-BK → Max tetanic-BK

### Force-Frequency Analysis Benefits:
- **Extended search window** captures all peak types including late peaks
- **Eliminates baseline drift effects** across different Hz stimulations
- **More accurate normalization** for force-frequency curves
- **Better comparison** between conditions with different baseline forces
- **Logical consistency**: Max Tetanic always >= Kinetics Peak
- **Consistent naming**: All BK corrected columns follow same pattern

## Enhanced Detection Methods (FvF v1.0)

### Relaxation Kinetics Detection
- **Consecutive Points**: 3 consecutive points below threshold
- **Tolerance**: 102% tolerance (allows 2% above threshold for noise)
- **Gradual Decay**: Count decrements by 1 instead of reset to 0
- **Fallback**: Single-point detection if consecutive method fails

### Contraction Kinetics Detection
- **Consecutive Points**: 3 consecutive points above threshold
- **Tolerance**: 98% tolerance (allows 2% below threshold for noise)
- **Gradual Decay**: Count decrements by 1 instead of reset to 0
- **Fallback**: Single-point detection if consecutive method fails

## Data Structure for Force-Frequency Analysis

### Key Columns for Analysis:
- **Stim Number**: Sequential stimulation number (1-14)
- **Hz**: Stimulation frequency (separate filterable column)
- **Max Tetanic Force (uN)**: Raw maximum force in T0 to T1+200ms window
- **Max tetanic-BK (uN)**: BK corrected maximum tetanic force (for absolute force analysis)
- **FF Normalized (%)**: Force as % of maximum within each well (ideal for F-F curves)
- **All kinetics metrics**: R10/50/80/90 Time, TT50P Time, TTP90 Time

### Recommended Analysis Approach:
1. **Force-Frequency Curves**: Plot FF Normalized (%) vs Hz for clean F-F relationships
2. **Absolute Force Curves**: Plot Max tetanic-BK vs Hz for baseline-corrected absolute values
3. **Raw Force Curves**: Plot Max Tetanic Force vs Hz for raw absolute values
4. **Kinetics vs Frequency**: Plot relaxation times vs Hz
5. **Condition Comparisons**: Group by Condition and Hz for statistics
6. **Peak Detection Window**: ±100ms around T1 (kinetics), T0 to T1+200ms (max tetanic)

## Processing and Normalization

1. **Data Smoothing**: A centered rolling mean with window size 3 is applied to raw force
2. **Normalization**: The smoothed force is normalized to a 0-1 scale for each well
3. **Peak Detection**: All metrics are calculated from the smoothed force data
4. **Representative Traces**: Include both raw and smoothed force averages across wells

## Relaxation Time Interpretation

- **R10 Time**: Very early relaxation - typically 0.01-0.03s for healthy tissue
- **R50 Time**: Half-relaxation time - typically 0.07-0.15s
- **R80 Time**: Late relaxation - typically 0.20-0.40s
- **R90 Time**: Near-complete relaxation - typically 0.40-0.70s

## Contraction Time Interpretation

- **TT50P**: Time to reach 50% of peak force - characterizes early phase contraction
- **TTP90**: Time to reach 90% of peak force - characterizes late phase contraction

## Force-Frequency Relationship

### Expected Patterns:
- **Peak Force**: Generally increases with frequency up to optimal Hz, then plateaus/decreases
- **Kinetics**: May show frequency-dependent changes in contraction and relaxation rates
- **Individual Variation**: Different conditions may show distinct F-F relationships

### Analysis Tips:
- **Hz Column**: Use for filtering and grouping in analysis software
- **Multiple Replicates**: Each well provides one data point per frequency
- **Statistical Analysis**: Consider both within-well and between-well variation
- **BK Correction**: All force values now baseline corrected for better comparisons
- **Consistent Naming**: Kinetics Peak-BK and Max tetanic-BK follow same pattern

## Note on Timing

All time values are reported with centisecond precision (2 decimal places).
This matches the data acquisition rate of 100Hz (0.01s intervals).
Force values and normalized trace values are stored with full precision.

## Representative Traces

The Representative Traces sheet includes:

1. **Normalized Trace**: Smoothed force normalized to 0-1 scale for shape comparison
2. **Raw Force**: Absolute force values in uN
3. **Smoothed Force**: Smoothed force values used for all kinetics calculations
4. **Standard Deviations**: Variability measures for all three data types

Note: Individual twitch normalization can be derived from Representative Traces if needed.