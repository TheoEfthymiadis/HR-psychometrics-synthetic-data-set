import numpy as np
import pandas as pd
from faker import Faker
import random
import sys
import xlwt
#import xlrd
import openpyxl
import string

folder_path = sys.path[0]
input_path = folder_path + '\\employees.xlsx'  # Name of the input file
output_path = folder_path + '\\noisy_employees.xlsx'  # Name of the output file
# There is a number of different seed functions that should be specified to produce a controlled and consistent output
fake = Faker()
Faker.seed(1321)
seed = 7
random.seed(10)
np.random.seed(seed=5)

nan_percentage = 0.02           # The percentage of missing values to be inserted into our data set
typographic_percentage = 0.05   # The percentage of typographic errors to be inserted into our data set
confusion_percentage = 0.02     # The percentage of typographic errors due to confusion
drop_percentage = 0.02          # The percentage of a record to be completely forgotten


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

# This function generates a random string of length = length
def get_random_string(length):
    letters = string.ascii_lowercase
    result_str = ''.join(random.choice(letters) for i in range(length))
    return result_str


# This function randomly inserts wrong characters and digits in a Data Frame to model typographic errors
def typographic_error(df, chars, digits, probability):
    for row in df.index:
        for column in chars:    # This loop inserts wrong characters in existing string columns
            error_chance = np.random.uniform(0, 1)
            if error_chance < probability:  # This simulates the random change for a typographic error to occur
                wrong_letter = int(np.random.uniform(0, len(df.at[row, column])))  # The string character to be replaced
                wrong_character = get_random_string(1)      # The random character that will replace the old character
                string_list = list(df.at[row, column])
                string_list[wrong_letter] = wrong_character
                df.at[row, column] = "".join(string_list)    # Update the old string with the new, wrong one

        for column in digits:    # This loop inserts wrong digits in existing numeric columns
            error_chance = np.random.uniform(0, 1)
            if error_chance < probability:  # This simulates the random change for a typographic error to occur
                wrong_digit = int(np.random.uniform(0, len(str(df.at[row, column]))))  # The digit that will be replaced
                wrong_number = str(int(np.random.uniform(0, 9)))
                string_list = list(str(df.at[row, column]))
                string_list[wrong_digit] = wrong_number
                df.at[row, column] = int("".join(string_list))  # Update the old number with a wrong digit

    return df


def nan_insertion(df, probability):  # This function inserts nan values in a Data Frame to model missing values
    for row in df.index:
        for column in df.columns:
            error_chance = np.random.uniform(0, 1)
            if error_chance < probability:  # This simulates the random change for a missing value to occur
                df[column].iloc[row] = np.nan
    return df


def drop_random_records(df, probability):
    for i in range(len(df)):
        error_chance = np.random.uniform(0, 1)
        if error_chance < probability:  # This simulates the chance of a record to be completely forgotten
            df = df.drop(labels=i)
    return df

# -------------------------------- Personal Profile errors ------------------------------------------------------------


employees_df = pd.read_excel(input_path, sheet_name='Professional_Profile', engine='openpyxl')
noisy_employees = typographic_error(employees_df.copy(), ['First Name', 'Last Name', 'Job Title'],
                                    ['Time Left', 'Children', 'Number of prev. Employers',
                                     'Salary'], typographic_percentage)
noisy_employees = nan_insertion(noisy_employees.copy(), nan_percentage)

for i in noisy_employees.index:  # Confusion in employee name
    error_chance = np.random.uniform(0, 1)
    if error_chance < confusion_percentage:  # This simulates the random change for a confusion to occur
        noisy_employees.at[i, 'First Name'] = fake.first_name()
    elif error_chance > 1 - confusion_percentage:
        noisy_employees.at[i, 'Last Name'] = fake.last_name()

for i in noisy_employees.index:  # Confusion in marital status and Department
    error_chance = np.random.uniform(0, 1)
    if error_chance < confusion_percentage:  # This simulates the random change for a confusion to occur
        if noisy_employees.at[i, 'Marital Status'] == 'Married':
            noisy_employees.at[i, 'Marital Status'] = 'Single'
        else:
            noisy_employees.at[i, 'Marital Status'] = 'Married'
    elif error_chance > 1 - confusion_percentage:  # The department is wrong!
        random_department = int(
            np.random.uniform(0, len(list(departments.keys()))))  # Random index to estimate department
        noisy_employees.at[i, 'Department'] = list(departments.keys())[random_department]

with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
    noisy_employees.to_excel(writer,  index=False, sheet_name='Professional_Profile')
writer.save()
writer.close()

# -------------------------------- Personal Profile errors end ---------------------------------------------------------

# --------------------------------------------------HR errors  ---------------------------------------------------------
hr_df = pd.read_excel(input_path, sheet_name='HR', engine='openpyxl')

noisy_hr = typographic_error(hr_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                             'Overtime', 'Chargeability', 'Employees Fired'],
                                                             typographic_percentage)
noisy_hr = nan_insertion(noisy_hr.copy(), nan_percentage)
noisy_hr = drop_random_records(noisy_hr, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_hr.to_excel(writer,  index=False, sheet_name='HR')
writer.save()
writer.close()
# --------------------------------------------------HR end  ------------------------------------------------------------

# --------------------------------------------------Sales errors  ------------------------------------------------------
sales_df = pd.read_excel(input_path, sheet_name='Sales', engine='openpyxl')

noisy_sales = typographic_error(sales_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                                   'Overtime', 'Chargeability', 'Total Sales',
                                                                   'Clients Asking'], typographic_percentage)
noisy_sales = nan_insertion(noisy_sales.copy(), nan_percentage)
noisy_sales = drop_random_records(noisy_sales, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_sales.to_excel(writer,  index=False, sheet_name='Sales')
writer.save()
writer.close()
# --------------------------------------------------Sales end  ---------------------------------------------------------

# --------------------------------------------------Product errors  ----------------------------------------------------
product_df = pd.read_excel(input_path, sheet_name='Product', engine='openpyxl')

noisy_product = typographic_error(product_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                                   'Overtime', 'Chargeability', 'Total Defects',
                                                                   'Number of Complaining Customers'], typographic_percentage)
noisy_product = nan_insertion(noisy_product.copy(), nan_percentage)
noisy_product = drop_random_records(noisy_product, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_product.to_excel(writer,  index=False, sheet_name='Product')
writer.save()
writer.close()
# --------------------------------------------------Product end  -------------------------------------------------------

# --------------------------------------------------Finance errors  ----------------------------------------------------
finance_df = pd.read_excel(input_path, sheet_name='Finance', engine='openpyxl')

noisy_finance = typographic_error(finance_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                                   'Overtime', 'Chargeability',
                                                                   'Non - Servicing Obligactions'], typographic_percentage)
noisy_finance = nan_insertion(noisy_finance.copy(), nan_percentage)
noisy_finance = drop_random_records(noisy_finance, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_finance.to_excel(writer,  index=False, sheet_name='Finance')
writer.save()
writer.close()
# --------------------------------------------------Finance end  -------------------------------------------------------

# --------------------------------------------------Legal errors  ------------------------------------------------------
legal_df = pd.read_excel(input_path, sheet_name='Legal', engine='openpyxl')

noisy_legal = typographic_error(legal_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                                   'Overtime', 'Chargeability', 'Successful Lawsuits',
                                                                   'Disputes amicably resolved'], typographic_percentage)
noisy_legal = nan_insertion(noisy_legal.copy(), nan_percentage)
noisy_legal = drop_random_records(noisy_legal, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_legal.to_excel(writer,  index=False, sheet_name='Legal')
writer.save()
writer.close()
# --------------------------------------------------Legal end  ---------------------------------------------------------

# --------------------------------------------------Strategy errors  ---------------------------------------------------
strategy_df = pd.read_excel(input_path, sheet_name='Strategy', engine='openpyxl')

noisy_strategy = typographic_error(strategy_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                                   'Overtime', 'Chargeability', 'Total Sales',
                                                                   'Number of Teams', 'Number of Projects'], typographic_percentage)
noisy_strategy = nan_insertion(noisy_strategy.copy(), nan_percentage)
noisy_strategy = drop_random_records(noisy_strategy, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_strategy.to_excel(writer,  index=False, sheet_name='Strategy')
writer.save()
writer.close()
# --------------------------------------------------Strategy end  ------------------------------------------------------

# --------------------------------------------------Technology errors  -------------------------------------------------
technology_df = pd.read_excel(input_path, sheet_name='Technology', engine='openpyxl')

noisy_technology = typographic_error(technology_df.copy(), ['Performance'], ['Year', 'Loyalty', 'Number of Promotions', 'Bonus',
                                                                   'Overtime', 'Chargeability',
                                                                   'Problematic Code Commits'],  typographic_percentage)
noisy_technology = nan_insertion(noisy_technology.copy(), nan_percentage)
noisy_technology = drop_random_records(noisy_technology, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_technology.to_excel(writer,  index=False, sheet_name='Technology')
writer.save()
writer.close()
# --------------------------------------------------Technology end  ----------------------------------------------------

# --------------------------------------------------Psychometric errors  -----------------------------------------------
psychometric_df = pd.read_excel(input_path, sheet_name='Psychometric_Indicators', engine='openpyxl')


noisy_psychometric = nan_insertion(psychometric_df.copy(), nan_percentage)
noisy_psychometric = drop_random_records(noisy_psychometric, drop_percentage)

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    noisy_psychometric.to_excel(writer,  index=False, sheet_name='Psychometric_Indicators')
writer.save()
writer.close()
# --------------------------------------------------Technology end  ----------------------------------------------------

