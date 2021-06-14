import numpy as np
import pandas as pd
from faker import Faker
import random
import datetime
import sys
import xlwt
#import xlrd
import openpyxl

folder_path = sys.path[0]
input_path = folder_path + '\\employees.xlsx'  # Name of the input file

# There is a number of different seed functions that should be specified to produce a controlled and consistent output
fake = Faker()
Faker.seed(1)
seed = 7
random.seed(10)
np.random.seed(seed=5)

employees_df = pd.read_excel(input_path, sheet_name='Professional_Profile', engine='openpyxl')
evaluation_performance = {'1': 'Low', '2': 'Medium', '3': 'High'}  # Dictionary that will be used for evaluation


# ----------------------- Working with the HR department -------------------------------------------------------------#
# We only extract the useful information for our department to execute calculations faster
department_df = employees_df[employees_df['Department'] == 'HR'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]

all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # HR specific evaluation metrics

        # Calculating all employees hired by the specific employee
        hired_employees_df = employees_df[
            (((employees_df['Recruiter ID'] == department_df.at[i, 'ID']) &
                (pd.to_datetime(employees_df['Date Hired'], format='%Y-%m-%d') <=
                 datetime.datetime.strptime(str(calendar_year), '%Y'))))].reset_index()[['ID', 'Date Hired', 'Time Left']]
        hired_employees_df['Time in Company'] = 0

        # Calculating the exact time that each of the recruited employees worked for the company
        for j in hired_employees_df.index:
            hired_employees_df.at[j, 'Time in Company'] = 2020 - hired_employees_df.at[j, 'Time Left'] - \
                                                              int(hired_employees_df.at[j, 'Date Hired'][0:4])

        evaluation['Total Time of hired employees(years)'] = hired_employees_df['Time in Company'].sum()  # Total employee time
        evaluation['Average Recruitment Time(months)'] = float("{:.2f}".format(np.random.uniform(1, 12)))  # Average recruitment time

        active_recruits = hired_employees_df[hired_employees_df['Time Left'] == 0]['Time Left'].count()  #How many recruits are still working in the company
        evaluation['Employees Fired'] = int(0.2*(len(hired_employees_df) - active_recruits))  # 20% of the recruits that left are considered fired

        all_evaluations.append(evaluation.copy())

hr_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    hr_df.to_excel(writer,  index=False, sheet_name='HR')
writer.save()
writer.close()

# ------------------------------------------- HR FINISHED --------------------------------------------------------------


# ----------------------- Working with the Sales department ------------------------------------------------------------
# We only extract the useful information for our department to execute calculations faster
department_df = []
department_df = employees_df[employees_df['Department'] == 'Sales'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]
all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # Sales specific evaluation metrics
        evaluation['Total Sales'] = int(np.random.uniform(1000, 100000))
        evaluation['Clients Asking'] = int(np.random.uniform(0, 5))
        all_evaluations.append(evaluation.copy())

sales_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    sales_df.to_excel(writer,  index=False, sheet_name='Sales')
writer.save()
writer.close()
# ------------------------------------------- Sales FINISHED -----------------------------------------------------------

# ----------------------- Working with the Product department ---------------------------------------------------------#
# We only extract the useful information for our department to execute calculations faster
department_df = []
department_df = employees_df[employees_df['Department'] == 'Product'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]
all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # Product specific evaluation metrics
        evaluation['Total Defects'] = int(np.random.uniform(10, 50))
        evaluation['Number of Complaining Customers'] = int(np.random.uniform(0, 20))
        all_evaluations.append(evaluation.copy())

product_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    product_df.to_excel(writer,  index=False, sheet_name='Product')
writer.save()
writer.close()
# ------------------------------------------- Product FINISHED ---------------------------------------------------------

# ----------------------- Working with the Finance department ---------------------------------------------------------#
# We only extract the useful information for our department to execute calculations faster
department_df = []
department_df = employees_df[employees_df['Department'] == 'Finance'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]
all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # Finance specific evaluation metrics
        evaluation['Non - Servicing Obligactions'] = int(np.random.uniform(0, 10000))
        all_evaluations.append(evaluation.copy())

finance_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    finance_df.to_excel(writer,  index=False, sheet_name='Finance')
writer.save()
writer.close()
# ------------------------------------------- Finance FINISHED ---------------------------------------------------------

# ----------------------- Working with the Legal department ---------------------------------------------------------#
# We only extract the useful information for our department to execute calculations faster
department_df = []
department_df = employees_df[employees_df['Department'] == 'Legal'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]
all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # Legal specific evaluation metrics
        evaluation['Successful Lawsuits'] = int(np.random.uniform(0, 3))
        evaluation['Disputes amicably resolved'] = int(np.random.uniform(0, 6))
        all_evaluations.append(evaluation.copy())

legal_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    legal_df.to_excel(writer,  index=False, sheet_name='Legal')
writer.save()
writer.close()
# ------------------------------------------- Legal FINISHED ---------------------------------------------------------

# ----------------------- Working with the Strategy department --------------------------------------------------------#
# We only extract the useful information for our department to execute calculations faster
department_df = []
department_df = employees_df[employees_df['Department'] == 'Strategy'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]
all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # Strategy specific evaluation metrics
        evaluation['Total Sales'] = int(np.random.uniform(1000, 10000))
        evaluation['Number of Teams'] = int(np.random.uniform(1, 10))
        evaluation['Number of Projects'] = int(np.random.uniform(1, 20))
        all_evaluations.append(evaluation.copy())

strategy_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    strategy_df.to_excel(writer,  index=False, sheet_name='Strategy')
writer.save()
writer.close()
# ------------------------------------------- Strategy FINISHED --------------------------------------------------------

# ----------------------- Working with the Technology department ------------------------------------------------------#
# We only extract the useful information for our department to execute calculations faster
department_df = []
department_df = employees_df[employees_df['Department'] == 'Technology'].reset_index()[['ID', 'Date Hired', 'Time Left',
                                                                        'Salary', 'Working Experience', 'Recruiter ID']]
all_evaluations = []   # Empty list to append the annual evaluations of the department employees
for i in range(len(department_df)):
    evaluation = {}
    evaluation['ID'] = department_df.at[i, 'ID']
    time_in_company = 2020 - department_df.at[i, 'Time Left'] - int(department_df.at[i, 'Date Hired'][0:4])

    for year in range(min(5, time_in_company)):
        calendar_year = 2020 - department_df.at[i, 'Time Left'] - year
        evaluation['Year'] = calendar_year   # Calendar year of the specific evaluation record
        evaluation['Loyalty'] = calendar_year - int(department_df.at[i, 'Date Hired'][0:4])  # Employee Loyalty
        evaluation['Number of Promotions'] = int(evaluation['Loyalty']/4)   # Number of promotions of the employee
        evaluation['Bonus'] = int(np.random.uniform(0, 30)/100*int(department_df.at[i, 'Salary']))   # Annual Bonus
        evaluation['Overtime'] = int(np.random.uniform(0, 20) / 100 * 1816)  # Annual working hours are 1816
        evaluation['Chargeability'] = int(np.random.uniform(0, 100))

        percentile = np.random.uniform(0, 100)   # Randomly estimate the percentile of the employee within the department
        if percentile < 15:
            evaluation['Department Percentile'] = 'Bottom 15%'
            evaluation['Performance'] = 'Low'
        elif percentile > 85:
            evaluation['Department Percentile'] = 'Top 15%'
            evaluation['Performance'] = 'High'
        else:
            evaluation['Department Percentile'] = 'Mid 70%'
            evaluation['Performance'] = evaluation_performance[str(int(np.random.uniform(1, 3)))]

        # Technology specific evaluation metrics
        evaluation['Problematic Code Commits'] = int(np.random.uniform(0, 20))
        all_evaluations.append(evaluation.copy())

technology_df = pd.DataFrame(all_evaluations)

with pd.ExcelWriter(input_path, engine='openpyxl', mode='a') as writer:
    technology_df.to_excel(writer,  index=False, sheet_name='Technology')
writer.save()
writer.close()
# ------------------------------------------- Strategy FINISHED --------------------------------------------------------
