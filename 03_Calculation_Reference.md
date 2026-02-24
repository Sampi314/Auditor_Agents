# 03 Calculation Reference

## Key Calculations

| Sheet | Row | Plain English | Readable Formula (RFL) | A1 Reference |
|---|---|---|---|---|
| Assum_Capex | 6 | Land | `=B6-C6` | `D6` |
| Assum_Capex | 7 | Buildings | `=B7-C7` | `D7` |
| Assum_Capex | 8 | Brewing Equipment | `=B8-C8` | `D8` |
| Assum_Capex | 9 | Packaging Lines | `=B9-C9` | `D9` |
| Assum_Capex | 10 | Kegs & Returnable Packaging | `=B10-C10` | `D10` |
| Calc_Volume | 7 | Seasonality Index | `=Assum_Volume!$B$46` | `B7` |
| Calc_Volume | 9 | CUMULATIVE GROWTH INDEX | `=B9` | `C9` |
| Calc_Volume | 12 | SkyBrew | `=Assum_Volume!$C$5*Assum_Volume!$C$8/12*B7*B9` | `B12` |
| Calc_Volume | 13 | AllDark | `=Assum_Volume!$C$5*Assum_Volume!$C$9/12*B7*B9` | `B13` |
| Calc_Volume | 14 | Total | `=B12+B13` | `B14` |
| Calc_Revenue | 5 | Price Index (cumulative) | `=B5` | `C5` |
| Calc_Revenue | 8 | SkyBrew On-Trade | `=(Calc_Volume!B12*Assum_Volume!$C$16)*(Assum_Pricing!$C$6*Assum_Volume!$C$22+Assum_Pricing!$C$7*Assum_Volume!$C$23+Assum_Pricing!$C$8*Assum_Volume!$C$21)*B5/1000` | `B8` |
| Calc_Revenue | 9 | SkyBrew Off-Trade | `=(Calc_Volume!B12*Assum_Volume!$C$17)*(Assum_Pricing!$D$6*Assum_Volume!$C$26+Assum_Pricing!$D$7*Assum_Volume!$C$27)*B5/1000` | `B9` |
| Calc_Revenue | 10 | AllDark On-Trade | `=(Calc_Volume!B13*Assum_Volume!$C$16)*(Assum_Pricing!$C$9*Assum_Volume!$C$22+Assum_Pricing!$C$10*Assum_Volume!$C$23+Assum_Pricing!$C$11*Assum_Volume!$C$21)*B5/1000` | `B10` |
| Calc_Revenue | 11 | AllDark Off-Trade | `=(Calc_Volume!B13*Assum_Volume!$C$17)*(Assum_Pricing!$D$9*Assum_Volume!$C$26+Assum_Pricing!$D$10*Assum_Volume!$C$27)*B5/1000` | `B11` |