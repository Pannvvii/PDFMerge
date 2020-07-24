[SETUP]

PDFMerge Was created to concatenate and automatically print Matthews Casket PDF Documentation in order of casket type 

PDFMerge relies on txt files for storing data. 
If they are missing or deleted, be sure that you have 6 files:

- currentList.txt		(can be empty)
- date&default.txt  		(saved with 4 lines of anything)
- OutputLocation    		(saved with 1 line of anything)
- ProgLocation.txt  		(saved with 1 line consisting of the directory where the program is saved)
- Finish_Database_Raw.txt	(saved with the casket database)
- Finish_Database_Organized.txt	(saved with the categories of casket surrounded by brackets: [MUMMY])

*Updating the raw database can be done at any time, and will only require a restart of the program to take effect.



[USE]

Select the directories and files to be used in the menu, and input the date 
(this is used to name the output file, and can be anything)

Program will Use your default windows Printer.

****When running the program, the files sent to the system printer will be output. 
****You will then be prompted.
****Only press enter once windows is done handling the PDFs or the program will crash.




Do not touch VersionStatics&Testing folder. This is for upgrades and development purposes only.
Modifiers for pyinstaller creation:

C:\Users\ckann\AppData\Local\Programs\Python\Python38>pyinstaller --hidden-import=pywintypes --additional-hooks-dir=C:\Users\ckann\Downloads\cryptohook-0.15 PDFMerge.py




!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!********NOTICE********!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Set the ProgLocation.txt with the Full path to the PDFMerge.exe File or this WILL NOT WORK