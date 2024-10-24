
To use the automation scripts "automate_AALMv3.0_yearxtox", it would be best to use the version of the AALMv3.0 included in this repository, as it has the warning messages pre-removed from the workbook code which prevents automation. For best results, download this entire repository to a folder on your c:\ drive where you have read and write privileges. Unzip the AALM.zip folder which contains the model and exectuable fortran files. If necessary, enable Macros in for this folder (add to trusted locations) via the Options tab in Excel. 

If you use the AALMv3.0 version from this repository there will be no need to follow the directions for changing the AALM VBA code.

To use these scripts, follow the directions to change folder and file paths where instructed (usually a comment to the right of the line where personal paths are necessary). The inputs required to automate are read at the top of the script, and they are in the form of a .txt file with each column representing a parameter which will be changed in every iteration (row) in the AALM. 

If you intend to use these scripts or have any questions, contact maria.cardelino@emory.edu 


______________________

SEE BELOW for the README.txt file associated with the online repository of the AALMv.30 directly provided by the EPA at https://github.com/USEPA/AALM 
_______________________

The AALMv3-0.zip contains all files necessary to run the All Ages Lead Model (AALM) version 3.0. The AALM v3.0 64-bit executable and the user interface have been most extensively tested on computers having Windows 11 Enterprise operating system with Microsoft 365 Excel. The AALM v3.0 32-bit executable has only been tested to determine its limited functionality as discussed in the users guide. 

To begin to use the model, unzip the zip file into a folder on your c:\ drive where you have read/write privileges. There are no other installation procedures. Generally, computer users will have read/write access within “c:\users\{your username}\”. Due to security safeguards, the AALM model will typically not run from your PC’s Documents folder or any network drive. 

After unzipping files, the AALM is started by opening AALMv3-0_mmddyy.xlsm, where v3-0 indicates version 3.0 and mmddyy should be 030124 or more recent. This is the AALM’s user interface. The first time this file is opened, there may be a red banner under the Excel menu bar reading, “SECURITY RISK  Microsoft has blocked macros from running because the source of this file is untrusted.” To resolve this issue, exit Excel. From the File Explorer, right click on AALMv3-0_mmddyy.xlsm and select properties. At the bottom of the file Properties, notice the Security warning. Click on the box by Unblock, then press <OK>. The next time the file is opened, there may be a yellow banner under the Excel menu bar reading, “SECURITY WARNING  Some active content has been disable.” To resolve this issue, press the <Enable Content> button to the right of the warning. The model is now ready to run simulations. You can save different Excel files with a filename of your preference for each of your different simulations. 

To familiarize yourself with the model it is recommended that you import and run example simulations. There are eight examples with detailed descriptions of these examples found within the "AALM Example Scenarios" appendix of the Users Guide for the AALM v3.0. The User Guide also contains detailed descriptions of how to modify settings within the AALM user interface as well as troubleshooting. Details related AALM exposure and biokinetic parameters, model optimization, and model evaluation are found within the Technical Support Document for the AALM v3.0. 

Model support is available from brown.james@epa.gov and PbHelp@epa.gov. 
