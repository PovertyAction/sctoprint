# sctoprint
 Convert SurveyCTO excel version to a printable document

# Install
```stata
 net install sctoprint, all replace from(https://raw.githubusercontent.com/PovertyAction/sctoprint/master)

```

# Example
```stata
 sctoprint using "X:\Projects 2020\01_instruments\03_xls\Phase one_v1.xlsx", title("Household Questionnaire") save(X:\Projects 2020\01_instruments\02_print\Phase one_v1") pdf replace clear

 sctoprint using "X:\Projects 2020\01_instruments\03_xls\Phase one_v1.xlsx", title("Household Questionnaire") save(X:\Projects 2020\01_instruments\02_print\Phase one_v1") word replace clear

```