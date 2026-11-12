# Thickness — stylus profilometry

One sub-folder per sample, with one or more scans across a scratched step:

```
data/Thickness/S12/S12_3SC_20261105.dat
```

**File format:** the text export of the profilometer — a few header lines such as
`"X Units:","MICRON"` and `"Z Units:","ANGSTROM"`, then one `x,z` pair per line.

**The file name** must contain the sample name.

**Accuracy.** The thickness is read automatically from the flat part of the
profile right after the step edge. This is **not always accurate**: a strongly
tilted sample or a rough film can mislead it. Always check each sample's chart in
the results. If the red window is in the wrong place, the menu lets you set it by
hand (`SAMPLE=X0:X1`, positions in µm).

**Examples.** `EX1/` and `EX2/` hold synthetic example files (not measurements).
Delete them when you add your own samples.
