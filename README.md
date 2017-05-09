# ESCI5980FinalProject
Program is intended to serve as a data extraction tool/calculator to be used with specific excel files generated in Trace Metal Isotope Lab UMNcalculating age from U-Th MC-ICP-MS data

Program includes all code necessary for calculating and exporting U-Th ages from MC-ICP-MS data files.

The project is packaged as a Unix Excutable that only works on MacOS so far,however, the attached python code could be ran on any system using Anaconda
To run the excutable, first download the file ageCalculation and the excel files onto your computer, and check to see if it's already converted to a unix excutable type file. To make sure ageCalculation is converted to Unix excutable, open terminal and type in command line the following code: 
```
chmod +x 
```
leave a space after +x but do not press enter yet, now drag the ageCalculation file into your terminal window and the path of ageCalculation should appear after +x. For example, if the ageCalculation is in your Downloads folder it could look like: 
```
chmod +x /Users/yourname/Downloads/ageCalculation
```
Press enter now and the file should be successfully converted.The icon of ageCalculation should appear as a exec type file. 
Double-click on the ageCalculation icon and the GUI should appear .

**Demo**  

To see how ageCalculation works, please download the excel files.
While ageCalculation is running, a GUI with the following layout should appear 
Enter the following info in the order it is presented here and click submit, then upload the corresponding files and finally click the calculate and export age button.       
