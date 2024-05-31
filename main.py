#import modules
import openpyxl
from openpyxl import load_workbook
import lxml
from configparser import ConfigParser
import time
import pandas as pd
import os, sys

#import enrollment final file
course_schedule_df = pd.read_csv(
    'C:/Users/fmixson/PycharmProjects/SP_Div_Enrollment/course_schedule.csv')

#filter to 9 week courses
nine_week_df = course_schedule_df[course_schedule_df['Session'].isin(['9A', '9B'])]
nine_week_df = nine_week_df.reset_index()
# course_schedule_df(course_schedule_df['Session'] == '9A') | (course_schedule_df['Session'] == '9B')
print(nine_week_df)

#roll through 9 week courses and assign ge category

comp = ['ENGL 100', 'ENGL 100S']
critical_thinking = ['ENGL 103', 'ENGL 110', 'PHIL 103', 'PSYC 103', 'READ 103']
oral_com = ['COMM 100', 'COMM 120']
math = ['MATH 112', 'MATH 112S', 'MATH 114', 'PSYC 210']
arts = ['ART 109', 'ARCH 112', 'DANC 100', 'DANC 101', 'HUM 109']
humanities = ['HIST 102', 'HIST 103', 'ART 109', 'HUM 100', 'HUM 109', 'HUM 125', 'PHIL 102']
soc_behav = ['COMM 110', 'ECON 201M', 'ECON 202M', 'KIN 108', 'HIST 102', 'HIST 103', 'POL 101', 'PSYC 101',
             'PSYC 150', 'PSYC 251', 'PSYC 261', 'PSYC 265', 'SOC 101', 'SOC 110', 'SOC 210', 'WGS 108', 'WGS 115',
             'WGS 140', 'WGS 202']
phys_sci = ['ESCI 104']
bio_sci = ['PSYC 241']
ethnic = ['SOC 210']
for i in range(len(nine_week_df)):
    if nine_week_df.loc[i,'Course_x'] in comp:
        nine_week_df.loc[i,'GE'] = '1.A English Composition'
    elif nine_week_df.loc[i,'Course_x'] in critical_thinking:
        nine_week_df.loc[i,'GE'] = '1.B Critical Thinking'
    elif nine_week_df.loc[i,'Course_x'] in critical_thinking:
        nine_week_df.loc[i,'GE'] = '1.C Oral Communication'
    elif nine_week_df.loc[i,'Course_x'] in math:
        nine_week_df.loc[i,'GE'] = '2 Mathematical Concepts'
    elif nine_week_df.loc[i,'Course_x'] in arts:
        nine_week_df.loc[i,'GE'] = '3.A Arts & Humanities'
    elif nine_week_df.loc[i,'Course_x'] in humanities:
        nine_week_df.loc[i,'GE'] = '3.B Arts & Humanities'
    elif nine_week_df.loc[i,'Course_x'] in soc_behav:
        nine_week_df.loc[i,'GE'] = '4 Social & Behavioral Sciences'
    elif nine_week_df.loc[i,'Course_x'] in phys_sci:
        nine_week_df.loc[i,'GE'] = '5.A Physical Sciences'
    elif nine_week_df.loc[i,'Course_x'] in soc_behav:
        nine_week_df.loc[i,'GE'] = '5.B Biological/Life Sciences'
    elif nine_week_df.loc[i,'Course_x'] in ethnic:
        nine_week_df.loc[i,'GE'] = '7 Ethnic Studies'

for i in range(len(nine_week_df)):
    if nine_week_df.loc[i, 'Instructor'] == '0':
        nine_week_df.loc[i, 'Instructor'] = 'Staff'

nine_week_ge_df = nine_week_df.dropna()
# nine_week_ge_df = nine_week_ge_df.reset_index()

# nine_week_df.to_excel('Test.xlsx')
nine_week_ge_df = nine_week_ge_df[['GE', 'Course_x', 'Class#', 'Session', 'Start', 'End', 'Current Enrollment', 'Capacity', 'Instructor', 'Room', 'Modality']]
nine_week_ge_df.to_excel("Nine Week Courses.xlsx")
#filter by ge category
#sort by ge category
#convert string to integers in fall div enrollment program