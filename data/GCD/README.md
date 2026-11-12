# GCD — galvanostatic charge–discharge

One sub-folder per sample. Inside it, one file per current density for the
charge, and one for the discharge:

```
data/GCD/S12/GCD_S12_3SC_AlCl3_0.2mAcm2_20261105.xlsx
data/GCD/S12/GCD_S12_3SC_AlCl3_0.2mA Discharge cm2_20261105.xlsx
data/GCD/S12/GCD_S12_3SC_AlCl3_0.4mAcm2_20261105.xlsx
data/GCD/S12/GCD_S12_3SC_AlCl3_0.4mA Discharge cm2_20261105.xlsx
```

**File format:** a Metrohm Autolab NOVA export to Excel (`.xlsx`) with the columns
`Corrected time (s)` (or `Time (s)`), `WE(1).Potential (V)` and `WE(1).Current (A)`.

**The file name:**

| Part | Example | Needed |
|---|---|---|
| `GCD_` then the sample name | `GCD_S12_` | yes — the sample name must come right after `GCD_` |
| current density | `0.2mAcm2`, `0.2mA_cm2`, `0.2mA cm2` | yes — in mA/cm² |
| the word `Discharge` | `0.2mA Discharge cm2` | yes, in every discharge file |
| electrolyte | `AlCl3` | optional |
| date | `20261105` | optional |

The areal capacitance is calculated from the linear part of the discharge only.

**Examples.** `EX1/` and `EX2/` hold synthetic example files (not measurements).
Delete them when you add your own samples.
