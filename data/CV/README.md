# CV — cyclic voltammetry

One sub-folder per sample, named like the sample:

```
data/CV/S12/CV_AlCl3_S12_[-1.2, 0.2]_50mVs_20261105.xlsx
data/CV/S12/CV_AlCl3_S12_[-1.2, 0.2]_20mVs_20261105.xlsx
data/CV/S13/CV_AlCl3_S13_[-1.2, 0.2]_50mVs_20261105.xlsx
```

**File format:** a Metrohm Autolab NOVA export to Excel (`.xlsx`). The first sheet
must have the columns `Potential applied (V)`, `Time (s)`, `WE(1).Current (A)`,
`Scan` and `WE(1).Potential (V)`.

**The file name** carries what the file itself does not say:

| Part | Example | Needed |
|---|---|---|
| sample name | `S12` | yes — must match the *Sample* (or *Alias*) column of `samples.xlsx` |
| potential window | `[-1.2, 0.2]` (or `[-1,2 , 0,2]` with decimal commas) | recommended — the lower and upper limit in V, in square brackets |
| scan rate | `50mVpers`, `50mV per s`, `50 mV/s` | recommended — otherwise it is measured from the data |
| electrolyte | `AlCl3` before the sample name | optional |
| date | `20261105` | optional |

**Long cycling runs** saved in several files (e.g. 2 × 50 cycles) must say which
round each file is: `1st`, `2nd`, … or `50c`. Such files are joined into one
continuous cycle count by menu option *[2]*:

```
data/CV/S12/CV_AlCl3_S12 1st 50c [-1,2 , 0,2] 50mVpers 20261105.xlsx
data/CV/S12/CV_AlCl3_S12 2nd 50c [-1,2 , 0,2] 50mVpers 20261105.xlsx
```

**Examples.** `EX1/` and `EX2/` hold synthetic example files (not measurements).
Delete them when you add your own samples.
