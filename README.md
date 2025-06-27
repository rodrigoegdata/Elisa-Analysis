# ELISA Calculation Using VBA

This project was created to analyze protein concentration after measurement by enzyme-linked immunosorbent assay (ELISA) using an Agilent BioTek Plate Reader.<br>

# Repository Content

1. [**ElisaResults.xlsm**](https://github.com/rodrigoegdata/Elisa-Analysis/blob/main/ElisaResults.xlsm) - This document has the raw data direct from the ELISA reader and the result after running the script following these steps:<br>  
   I) The raw absorbance values for each well in the ELISA plate are exported to a .xls file and analyzed in Microsoft Excel (Raw data).  
   II) The blank was averaged excluding outliers based on Tukey's fences method (Blank selection and averaging), and subtracted from all values in the plate (Blank subtraction).  
   III) The standard curve was constructed and a scatter graph was plotted based on the concentration of the standard vs. absorbance measured in a 8-point 2-fold dilution curve (Standard curve).  
   IV) The values of the slope and intercept of the standard curve were calculated using the linest function (Slope and Intercept) and the curve equation was applied to all values within the absorbance range of the standard curve to obtain the concentration of each sample dilution in ng/mL (Equation application).  
   V) Values were multiplied by the dilution factor of each line in the serial dilution (2-fold dilution, 1 to 128) (Adjust to serial dilution).  
   VI) Next, the average of the more consistent values was calculated excluding outliers (in this case values higher or lower than 0.5x the interquartile range were excluded) for each sample (Dilution selection and averaging).  
   VII) Recovered values were then multiplied by the initial dilution factor (standard = 1000, samples = 75) to obtain the sample concentration in ng/mL (Adjust to sample dilution).  
   VIII) Values were then divided by 1000 to obtain the concentration in µg/mL (Final concentration).<br>

2. [**ElisaCalc.bas**](https://github.com/rodrigoegdata/Elisa-Analysis/blob/main/ElisaCalc.bas) - File containing VBA code use to analyze the data.<br>

3. [**README.md**](https://github.com/rodrigoegdata/Elisa-Analysis/blob/main/README.md) - Explains the files in this repository

### Abbreviations

Q1 - first quartile <br>
Q3 - third quartile <br>
IQR - interquartile range <br>
lowBd - lower bound <br>
upBd - upper bound <br>
Averagelbub - average excluding outliers
