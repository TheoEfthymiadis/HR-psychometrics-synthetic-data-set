
Required libraries to run the Python scripts:
Package	Version	Latest Version
Faker	5.0.2	5.0.2
Pillow	8.0.1	8.0.1
cycler	0.10.0	0.10.0
et-xmlfile	1.0.1	
jdcal	1.4.1	1.4.1
kiwisolver	1.3.1	1.3.1
numpy	1.19.3	1.19.4
openpyxl	3.0.5	3.0.5
pandas	1.1.5	1.2.0
pip	19.0.3	20.3.3
pyparsing	2.4.7	2.4.7
python-dateutil	2.8.1	2.8.1
pytz	2020.4	2020.5
radar	0.3	0.3
setuptools	40.8.0	51.1.0.post20201221
six	1.15.0	1.15.0
text-unidecode	1.3	1.3
xlrd	2.0.1	2.0.1
xlwt	1.3.0	1.3.0

# Running the implementation
The implementation of everything mentioned above was executed through the development of 3 Python scripts. The scripts were developed in a virtual environment using the Pycharm software. 
In order to reproduce the results, it is suggested that the user creates a similar virtual environment and installs all necessary python libraries that are listed in the ‘README.txt’ file. 
Pay extra attention to the version of the ‘numpy’ library, because the latest version is prone to a number of bugs and should not be preferred.
After setting up the virtual environment and installing the libraries, the scripts should be executed in a specific order:
•	Run the ‘Personal_Profile.py’ script by using the command line. It will produce an excel file in the current folder containing the personal information and psychometric profile 
of all employees. The file is named ‘folder_path/employees.xlsx’. 
•	Run the ‘Departments.py’ script by using the command line. This will read the excel file that was created in the previous step, assign the employees to random departments, 
estimate their evaluation metrics and append the information back to the ‘employees.xlsx’ file. This version of the file is also provided in the deliverable.
•	Run the ‘Noise_Insertion.py’ script by using the command line. This will read the ‘employees.xlsx’ file, insert noise to the data and provide an output file named ‘noisy_employees.xlsx’. 
This file is also provided in the deliverable.
A large number of seed functions were used to control the random data generation from various python libraries. The results should be completely reproducible.
