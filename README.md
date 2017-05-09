# ESCI5980FinalProject
Program is intended to serve as a data extraction tool/calculator to be used with specific excel files generated in Trace Metal Isotope Lab, UMN, to calculate age from U-Th MC-ICP-MS data

Program includes all code necessary for calculating and exporting U-Th ages from MC-ICP-MS data files.

The project is packaged as a Unix Executable that only works on MacOS at the moment, however, the attached python code could be ran on any system using Anaconda
To run the executable, first download the file ageCalculation and the excel files onto your computer, and check to see if it's already converted to a Unix executable type file. To make sure ageCalculation is converted to Unix executable, open terminal and type in command line the following code: 
```
chmod +x 
```
leave a space after +x but do not press enter yet, now drag the ageCalculation file into your terminal window and the path of ageCalculation should appear after +x. For example, if the ageCalculation is in your Downloads folder it could look like: 
```
chmod +x /Users/yourname/Downloads/ageCalculation
```
Press enter now and the file should be successfully converted. The icon of ageCalculation should appear as a exec type file. 
Double-click on the ageCalculation icon and the GUI should appear.

 

To see how ageCalculation works, please download the excel files.
While ageCalculation is running, a GUI with the following layout should appear 
![alt text](ESCI5980FinalProject/Screen Shot 2017-05-09 at 15.31.53.png)

Enter the following info in the order it is presented here and click submit, then upload the corresponding files and finally click the calculate and export age button.       

The ageCalculation.py program is composed of several different classes: Application, isofilter, Ucalculation, Thcalculation, backgraound_values, chemblank_values.
The Application class is where the main GUI is written, and its core calculation is done in method
```
Age_Calculation()
```
the isofilter, Ucalculation, Thcalculation, background_values and chemblank_values classes are instantiated inside the Age_Calculation() method.
Please see below for the details of each class.

## class isofilter() 

Requires openpyxl and numpy.

Requires input parameters filename, columnletter, and filternumber inputs. Calculates mean, standard deviation, and total counts of filtered and unfiltered data.

 -def getMean(): calculates mean of unfiltered data
 -def getStanddev(): calculates standard deviation of unfiltered data
 -def getCounts(): calculates counts of unfiltered data
 -def Filtered_mean(): filters data depending on criteria and calculates resulting mean
 -def Filtered_err(): filters data dpending on criteria and calculates resulting 2s error
 -def Filtered_counts(): filters data dpending on criteria and calculates resulting counts

## class chem_blank()

Requires input parameters filename, columnletter, and isotope analyzed. Calculates mean, counts, and relative 2s error for chemistry blanks, for use with Age Calculation.

  -def calc(): calculates and returns list of mean, counts, and relative 2s error for specified isotope

## class Ucalculation()

Requires numpy, and pandas
Requires input parameters spike used, abundance sensitivity, U filename. Calculates ratios and cps values from Uranium run needed for use in Th and Age Calculation functions.

  -def U_normalization_forTh(): calculates and returns list of measured 236/233 ratio and error, normalized 235/233 ratio and error, and corrected 236/233 ratio and error, for use in Th function.
  -def U_normalized_forAge(): calculates and returns list of normalized 235/233 ratio and error, 234/235 normalized and corrected ratio and error, unfiltered cycles of 233 and filtered cycles of 234/235, and unfiltered mean of 233, for use in Age Calculation function.

## class Thcalculation()


Requires input parameters spike used, abundance sensitivity, Th filename, and U_normalized_forTh() output.

Calculates ratios and cps values from Th run needed for use in Age Calculation function.

  -def Th_normalization_forAge(): calculates and returns a list of corrected and normalized 230/229 ratio and error, corrected and normalized 232/229 ratio and error, and unfiltered mean and cycles of 229, for use in Age Calculation function.

## class background_values()


Requires U wash file, Th wash file. Calculates wash values for use in Age Calculation function.

  -def U_wash(): calculates and returns list of 233, 234, and 235 wash values in cps for use in Age Calculation function.
  -def Th_wash(): calculates and returns 230 wash value in cpm for use in Age Calculation function.

## class chemblank_values()



Requires input parameters spike used, chem spike weight, wash and run files for U, wash and run files for Th. Calculates chem blank values for use in Age Calculation function.

  -def blank_calculate(): calculates and returns a list of 238 chem blank value and error in pmol, 232 chem blank value and error in pmol, and 230 chem blank value and error in fmol, for use in Age Calculation function. 
