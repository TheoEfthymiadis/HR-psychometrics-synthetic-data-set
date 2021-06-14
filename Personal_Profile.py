import numpy as np
import pandas as pd
from faker import Faker
import random
import datetime
import uuid
import radar
import sys
from openpyxl import load_workbook

folder_path = sys.path[0]
output_path = folder_path + '\\employees.xlsx'   # This is the file name where the output will be saved

# There is a number of different seed functions that should be specified to produce a controlled and consistent output
fake = Faker()
Faker.seed(1)
seed = 7
random.seed(10)
np.random.seed(seed=5)
# ------------------------------------------ PROBLEM PARAMETERS -------------------------------------------------------#

# The employees are split in 5 classes based on their age. The probability of each class to occur is imbalanced
employee_age_range = {'1': [18, 23], '2': [24, 29], '3': [30, 44], '4': [45, 54], '5': [55, 65]}
employee_age_dist = [0.05, 0.75, 0.15, 0.04, 0.01]

# The academic background describes the probability of each class to hold a: High School, Bachelor, MSc, PhD degree
# the 2nd class has two different distributions to reflect the variance in academic background of that age interval
academic_background = {'1': ['High School', 0], '2': ['Bachelor', 3], '3': ['MSc', 5], '4': ['PhD', 9]}
academic_background_dist = {'1': [0.7, 0.3, 0, 0], '2': [[0.2, 0.6, 0.2, 0], [0.1, 0.45, 0.4, 0.05]],
                            '3': [0.1, 0.45, 0.4, 0.05], '4': [0.1, 0.45, 0.4, 0.05], '5': [0.1, 0.45, 0.4, 0.05]}

basic_income = 10000   # The basic income equals the lowest possible salary an employee can receive
e = 0.1  # Maximum percentage of annual salary raise

number_employees = 1000   # The total number of employees working in our company

# The different departments of the company and the corresponding job titles
departments = {
      'Sales': ['Director of Sales', 'Sales Manager', 'Area Sales Manager', 'Sales Executive',
                'Sales Representative', 'Brand Ambassador', 'Sales Associate'],
      'Product': ['Production Manager', 'Production Technician', 'Product Integration Assistant',
                  'Product Communications Planner', 'Product Brand Associate', 'Product Implementation Manager',
                  'Product Creative Analyst'],
      'Finance': ['Pricing Analyst', 'Financial Analyst', 'Credit Risk Analyst', 'Portfolio Analyst',
                  'Investment Manager', 'Credit Risk Manager', 'Finance Manager'],
      'HR': ['Recruiting Manager', 'Recruiting Assistant', 'Talent Consultant', 'Benefits Counselor',
             'Retention Specialist', 'Workforce Analyst', 'HR Coordinator'],
      'Legal': ['Resolution Specialist', 'Legal Analyst', 'Legal Research Analyst', 'Manager Legal', 'Defense Attorney',
                'Patent Attorney', 'Attorney General'],
      'Strategy': ['CIO', 'CEO', 'Strategy Director', 'Strategic Planner', 'Business Strategy Manager',
                   'Strategy Analyst', 'Business Planner'],
      'Technology': ['Information Security Manager', 'IT Support', 'IT Director', 'Software Engineer',
                     'Database Administrator', 'Network Engineer', 'Software Engineering Manager']
}
# ---------------------------------------- PROBLEM PARAMETERS END -----------------------------------------------------#

all_employees = []   # An empty list to store all employee data in order to store them in a pandas dataframe later
all_psychometrics = []   # An empty list to store all psychometric data to store them in a pandas dataframe later

# Through the following for loop, we will create personal, professional and psychometric data for all employees
for i in range(number_employees):
    employee = {}       # Dictionary to store all personal data for the specific employee. Will be appended to the list
    psychometrics = {}  # Dictionary to store all psychometrics for the specific employee. Will be appended to the list

    rnd = random.Random()  # We have to control the seed in every iteration to return consistent UUIDs
    rnd.seed(i)
    employee_id = str(uuid.UUID(int=rnd.getrandbits(128), version=4))   # Employee ID
    employee['ID'] = employee_id
    psychometrics['ID'] = employee_id
    employee['First Name'] = fake.first_name()   # First Name
    employee['Last Name'] = fake.last_name()     # Last Name

    gender = random.random()                     # Gender
    if gender < 0.5:
        employee['Gender'] = 'M'
    else:
        employee['Gender'] = 'F'

    marital = random.random()                     # Marital Status
    if marital < 0.5:
        employee['Marital Status'] = 'Married'
    else:
        employee['Marital Status'] = 'Single'

    age_interval = np.random.choice(list(employee_age_range.keys()), p=employee_age_dist) # Generate random age interval
    age = int(np.random.uniform(employee_age_range[str(age_interval)][0],
                                employee_age_range[str(age_interval)][1]))  # Generate exact age

    old_or_new = np.random.uniform(0, 1)  # Randomly decide if employee still works for the company or not
    if old_or_new < 0.8:   # This means the employee has left the company. 80% of our records refer to old employees
        works_here = False
        time_left = int(np.random.uniform(0, 20))   # Random time interval that the employee has left the company
        age_real = age + time_left    # Adjust age for the old employees
        employee['Time Left'] = time_left
    else:
        age_real = age
        works_here = True
        employee['Time Left'] = 0

    employee['Works Here'] = works_here     # Define if the employee still works for the company
    birth_year = 2020 - age_real
    birthday = radar.random_datetime(start=datetime.date(year=birth_year, month=1, day=1),
                                     stop=datetime.date(year=birth_year, month=12, day=31)).strftime("%Y-%m-%d")
    employee['Birthday'] = birthday         # Employee's birthday

    number_children = round(max(0, np.random.normal(1.5*min(1, age_real/35), 0.5)))
    employee['Children'] = number_children

    # In order to calculate academic background, we need to distinguish the age interval [24,29] from the rest
    if ((age_interval == '2') & (age < 27)):
        random_background = np.random.choice(list(academic_background.keys()),
                                             p=academic_background_dist[age_interval][0])
    elif ((age_interval == '2') & (age >= 27)):
        random_background = np.random.choice(list(academic_background.keys()),
                                             p=academic_background_dist[age_interval][1])
    else:
        random_background = np.random.choice(list(academic_background.keys()),
                                             p=academic_background_dist[age_interval])

    employee['Academic Background'] = academic_background[random_background][0]  # Academic Background of the Employee
    study_period = academic_background[random_background][1]    # Years spent studying

    # The year that the employee was hired will be calculated here
    if works_here == True:
        year_hire = int(np.random.uniform((birth_year + 18 + study_period), 2020))
    else:
        year_hire = int(np.random.uniform((birth_year + 18 + study_period), 2020 - time_left))

    date_hire = radar.random_datetime(start=datetime.date(year=year_hire, month=1, day=1),
                                      stop=datetime.date(year=year_hire, month=12, day=31)).strftime("%Y-%m-%d")
    employee['Date Hired'] = date_hire  # The exact date that the employee was hired

    if works_here == True:
        work_exp = 0.8*(year_hire - (birth_year + 18 + study_period)) + (2020 - year_hire)
    else:
        work_exp = 0.8 * (year_hire - (birth_year + 18 + study_period)) + (2020 - time_left - year_hire)

    employee['Working Experience'] = round(max(0, work_exp)) # The working experience of the employee
    # Number of Previous Employers
    if max(0, 0.8*(year_hire - (birth_year + 18 + study_period))) == 0:
        employee['Number of prev. Employers'] = 0
    else:
        employee['Number of prev. Employers'] = int(0.8*(year_hire - (birth_year + 18 + study_period))/3) + 1

    # The salary of the employee at the date of hire
    salary_hired = np.random.uniform(basic_income,
                                      max(basic_income, basic_income +
                                        0.8*(year_hire - (birth_year + 18 + study_period)) * 2000 + study_period * 500))
    if works_here == True:
        salary = int(np.random.uniform(salary_hired, salary_hired*(1+e)**max(1, (2020 - year_hire))))
    else:
        salary = int(np.random.uniform(salary_hired, salary_hired * (1 + e)**max(1, (2020 - time_left - year_hire))))
    employee['Salary'] = salary    # Calculation of the employee salary

    random_department = int(np.random.uniform(0, len(list(departments.keys()))))   # Random index to estimate department
    employee['Department'] = list(departments.keys())[random_department]   # Random calculation of employee department
    employee['Job Title'] = departments[employee['Department']][int(np.random.uniform(
        0, len(departments[employee['Department']])))]

    # ----- Psychometric Data ------ #
    # The only constraint here is that each BIG 5 factor lies in-between its 2 facets

    # Concientiousness falls between Orderliness and Industriousness
    psychometrics['Orderliness'] = int(np.random.uniform(0, 100))
    psychometrics['Industriousness'] = int(np.random.uniform(0, 100))
    psychometrics['Concientiousness'] = int(np.random.uniform(
        min(psychometrics['Industriousness'], psychometrics['Orderliness']),
        max(psychometrics['Industriousness'], psychometrics['Orderliness'])))

    # Neuroticism falls between Withdrawal and Volatility
    psychometrics['Withdrawal'] = int(np.random.uniform(0, 100))
    psychometrics['Volatility'] = int(np.random.uniform(0, 100))
    psychometrics['Neuroticism'] = int(np.random.uniform(
        min(psychometrics['Withdrawal'], psychometrics['Volatility']),
        max(psychometrics['Withdrawal'], psychometrics['Volatility'])))

    # Extraversion falls between Enthusiasm and Assertiveness
    psychometrics['Enthusiasm'] = int(np.random.uniform(0, 100))
    psychometrics['Assertiveness'] = int(np.random.uniform(0, 100))
    psychometrics['Extraversion'] = int(np.random.uniform(
        min(psychometrics['Enthusiasm'], psychometrics['Assertiveness']),
        max(psychometrics['Enthusiasm'], psychometrics['Assertiveness'])))

    # Openness to Experience falls between Intellect and Openness
    psychometrics['Intellect'] = int(np.random.uniform(0, 100))
    psychometrics['Openness'] = int(np.random.uniform(0, 100))
    psychometrics['Openness to Experience'] = int(np.random.uniform(
        min(psychometrics['Intellect'], psychometrics['Openness']),
        max(psychometrics['Intellect'], psychometrics['Openness'])))

    # Agreeableness falls between Compassion and Politeness
    psychometrics['Compassion'] = int(np.random.uniform(0, 100))
    psychometrics['Politeness'] = int(np.random.uniform(0, 100))
    psychometrics['Agreeableness'] = int(np.random.uniform(
        min(psychometrics['Compassion'], psychometrics['Politeness']),
        max(psychometrics['Compassion'], psychometrics['Politeness'])))

    all_employees.append(employee.copy())
    all_psychometrics.append(psychometrics.copy())

# Saving our results in 2 different pandas data frames
employee_df = pd.DataFrame(all_employees)           # pandas data frame to store professional data
psychometrics_df = pd.DataFrame(all_psychometrics)  # pandas data frame to store psychometric indicators

# In this last part, we want to also add the recruiter that hired the employee. The recruiter has to meet 2 conditions:
# 1) He/She must be working in HR, 2) He/She must be working for the company before the employee

employee_df['Recruiter'] = ''
employee_df['Recruiter ID'] = ''

for i in range(number_employees):
    # Identifying all possible recruiters
    possible_recruiters = employee_df[(
        (employee_df['Department'] == 'HR')&
        (pd.to_datetime(employee_df['Date Hired']) < datetime.datetime.strptime(employee_df.at[i, 'Date Hired'], '%Y-%m-%d' ))
    )].reset_index()[['ID', 'First Name', 'Last Name']]

    # In case there is at least one possible recruiter (due to random generation), we randomly select one of them
    if len(possible_recruiters) == 0:
        employee_df.at[i, 'Recruiter'] = 'NULL'
        employee_df.at[i, 'Recruiter ID'] = 'NULL'
    else:
        recruiter = int(np.random.uniform(0, len(possible_recruiters)))
        employee_df.at[i, 'Recruiter'] = possible_recruiters.at[recruiter, 'First Name'] + ' ' + \
                                                         possible_recruiters.at[recruiter, 'Last Name']
        employee_df.at[i, 'Recruiter ID'] = possible_recruiters.at[recruiter, 'ID']

# Write the results to an excel file inside the same folder

with pd.ExcelWriter(output_path) as writer:
    employee_df.to_excel(writer, sheet_name='Professional_Profile', index=False)
    psychometrics_df.to_excel(writer, sheet_name='Psychometric_Indicators', index=False)
writer.save()
writer.close()
