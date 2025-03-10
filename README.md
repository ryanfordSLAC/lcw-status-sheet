![logo](./logos/SLAC-lab-hires.png)
# Introduction

The Test_LCW.py file contains code to generate a daily Excel version of the LCW status sheet.

Modify the program prior to moving to production on a SLAC VM:  esh-rp-survey01 (C:\LCW\Prod_LCW.py):
* Uncomment final line in program to turn on the email feature
* After "############End Table 3", uncomment production (filename = "C:\\LCW\\LCW.xlsx" #prod) and comment test (filename = "C:\\Users\\ryanford\\OneDrive - SLAC National Accelerator Laboratory\\2025\\LCW.xlsx" #test).  
* Save Prod_LCW.py to V drive (V:\ESH\OHP\Field Instruments), and copy to C:\LCW\Prod_LCW.py on the VM.


# SLAC National Accelerator Laboratory
The SLAC National Accelerator Laboratory is operated by Stanford University for the US Departement of Energy.  
[DOE/Stanford Contract](https://legal.slac.stanford.edu/sites/default/files/Conformed%20Prime%20Contract%20DE-AC02-76SF00515%20as%20of%202022.10.01.pdf)

# Lisence

We are beginning with the BSD-2 lisence but this is an open discussion between code authors, SLAC management, and DOE program managers along the funding line for the project.  




