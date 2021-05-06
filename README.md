# wind-app

The program to calculate wind power in specific regions with online extraction of the speed (via pyown API) and user-entered parameters (turbine's radius and efficiency). 
PyQt5 was used for GUI.  


--SELECT THE METHOD TO RUN THE PROGRAM--

The folder 'WindApplication' contains WindPowerApp.py and WindPowerApp.exe files. 

Choose WindPowerApp.exe, if you want to open the program without installation of any dependencies. 

If you want to access the program through the WindPowerApp.py file, please, refer to the Dependencies section below. 

--DEPENDENCIES--

Prerequisite: Python 3.0 or later

Install the libraries thtough the command window:

--pip install PyQt5 
--pip install sys
--pip install pyowm
--pip install openpyxl
--pip install datetime

--FOLDER CONTENT--
history.xlsx - Database of all process
ProjectReport.pdf - The theoretical explanation
about.py - window for the program description
forall.py - window for the calculation of power in all regions
WindPowerApp.exe - Application execution 
WindPowerApp.py - Python file of the program 
