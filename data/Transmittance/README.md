# Transmittance

One sub-folder per sample, with one or more spectra:

```
data/Transmittance/S12/S12_3SC_20261105.Sample.asc
```

**File format:** the ASCII export of a PerkinElmer UV/Vis spectrometer (UV WinLab),
with `#GR` and `#DATA` sections, wavelength in nm and the signal in %T.

**The file name** must start with the sample name. A number of layers written as
`3SC` and a date `20261105` are read too, if present.

The value at 550 nm is collected for every spectrum into a `.dat` text file.

**Examples.** `EX1/` and `EX2/` hold synthetic example files (not measurements).
Delete them when you add your own samples.
