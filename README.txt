Last Updated: 6/01/17

Summary:
The "Populate" script reads in two CSV files and automatically populates the Tivoli form online.
This is done by reading each CSV line by line, and putting each line into a list.  The program finds the 
data based on the title of the row, then calls the associated function to populate the appropriate text
area.  When a new list is read, a new chrome window is opened to be populated.

Included Files:
cgi-bin - folder contains front-end file called Populate.py; this folder and name are necessary to start the python server locally
Populate.py - source code
chromedriver.exe - the executable to open the chrome window
mail VBScript Script File - VBScript file to send notification email to user
README.txt - this file

How to Run:
1. In the command line make sure you "cd" in the correct folder.  You should be outside of the cgi-bin folder
2. To start the local python server type the following in the command line
	python -m CGIHTTPServer
3. Open your browser and go to http://localhost:8000/cgi-bin/Populate.py
4. Choose your two CSV files to be read, then click "Submit"