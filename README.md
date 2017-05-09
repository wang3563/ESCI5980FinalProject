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
Enter the following info in the order it is presented here and click submit, then upload the corresponding files and finally click the calculate and export age button.       

The ageCalculation.py program is composed of several different classes: Application, isofilter, Ucalculation, Thcalculation, backgraound_values, chemblank_values.
The Application class is where the main GUI is written, and its core calculation is done in method
```
Age_Calculation()
```
the isofilter, Ucalculation, Thcalculation, background_values and chemblank_values classes are instantiated inside the Age_Calculation() method.
Please see below for the details of each class.
