#create a function that takes in the file path and returns the dataframes

import re
import sys
import pdfplumber
import pandas as pd
from collections import namedtuple
from datetime import datetime

def event_calculation(ins_amount, rev_of_cc, total_rev):
    return float((ins_amount * rev_of_cc) / total_rev)

def alloc_perc_calculation(ins_amount, total_rev):
    return float((ins_amount / total_rev) * 100)

#def extract_dataframes(file_path):
    #text = ""
    #with pdfplumber.open(file_path) as pp:
        #for page in pp.pages:
            #page_text = page.extract_text()
            #if page_text:  # Check if text was extracted
                #text += page_text + "\n"
#
    ## Add your existing data extraction and dataframe creation logic here
#
    #return df0, df1, df2  # Return the dataframes you want to extract
#
## Call the function with the file path
#file_path = "PPs multiple groups/PROJECT PLAN-BC210K-69-72_Chen Ni.pdf"
#df0, df1, df2 = extract_dataframes(file_path)


Vessel_line = namedtuple('Vessel_line', 'Which_Group Which_Vessel Hull_Number Class_Number VID Start_Date Completion_Date Vessel_Contract_Amount')
Installment_line = namedtuple('Installment_line', 'Installment_Type Installment_Amount Installment_Date')
Org_line = namedtuple('Org_line', 'Organization_Name Cost_Center_Code Key_Member Intercompany Estimated_Hours Nonreimbursable_Expenses Reimbursable_Expenses Cost Revenue Comments')
Event_line = namedtuple('Event_line', 'Milestone_or_Installment_Type Total_Installment_Amount Installment_Date Organization_Name Cost_Center Event_Amount Total_Revenue Allocation_Percentage')

num_groups = 0
num_vessels = 0
which_group = []
which_vessel = []

inst_type = []
temp_inst_amount = []
temp_inst_date = []
inst_amount = []
inst_date = []
vca = []

opn = []
opn_ic = []
hull_num = []
class_num = []
vid = []
start_date = []
compl_date = []

org_name = []
cost_center = []
key_member = []
ic = []
hours_estimate = []
non_reimb_exp = []
reimb_exp = []
cost = []
revenue = []
comments = []

event_amounts = []
alloc_percentages = []


file = "file.pdf"

# Get the text from all pages onto the same string
text = ""
with pdfplumber.open(file) as pp:
    for page in pp.pages:
        page_text = page.extract_text()
        if page_text:  # Check if text was extracted
            text += page_text + "\n"


if text.startswith("Date:" or "Group"):
    print("Normal Extraction")
else:
    text = ""
    with pdfplumber.open(file) as pp:
        for page in pp.pages:
            page_text = page.extract_text_simple()
            if page_text:
                text += page_text + "\n"

print(text)

if text == "":
    sys.exit("Text is empty. Exiting script.")

#sometime it wont detect the header or the last page so just focus on getting the vessel info, installments, and revenue tables
gp_vl_pattern = re.compile(r'Group (\d+) \. (\d+) Vessel\(s\)')
which_vl_pattern = re.compile(r'(Vessel #\d+)')
inst_ty_pattern_l1 = re.compile(r'(?<!\S)(Installment|Milestone|<Select>)(\s(.+))')
inst_ty_pattern_l2 = re.compile(r'(?<!\S)(14th|<Auto>)(\s(.+))')
inst_amount_pattern = re.compile(r'(?<!\S)Amount (.+)')
inst_date_pattern = re.compile(r'(?<!\S)Date (.+)')
opn_pattern = re.compile(r'OPN: (\S+)')
opn_ic_pattern = re.compile(r'OPN \(Intercompany\): (\S+)')
hull_num_pattern = re.compile(r'Hull Number: (\S+)')
class_num_pattern = re.compile(r'Class Number: (\S+)')
vid_pattern = re.compile(r'VID\*: (\S+)')
start_date_pattern = re.compile(r'Start Date\*: (\S+)')
compl_date_pattern = re.compile(r'Completion Date\*: (\S+)')
vca_pattern = re.compile(r'Start Date\*:.*\n. (.*)')
org_info_pattern = re.compile(r'(.* - .*) (\d{4,6}) (.*) (N|Y) ([\d,]+) . ([\d,]+) . ([\d,]+) . ([\d,]+) . ([\d,]+)')
grand_total_pattern = re.compile(r'Grand Total ([\d,]+) . ([\d,]+) . ([\d,]+) . ([\d,]+) . ([\d,]+)')
comments_pattern = re.compile(r'Comments: (.*)')

count = 0
for line in text.split('\n'):
    # which Group and Vessel
    match = gp_vl_pattern.search(line)
    if match:
        which_group.append("Group " + match.group(1))
    match = which_vl_pattern.search(line)
    if match:
        which_vessel.append(match.group(1))
    # installment type 1
    match = inst_ty_pattern_l1.search(line)
    if match:
        for i in match.group(2).split():
            inst_type.append(i)
    # installment type 1
    match = inst_ty_pattern_l2.search(line)
    if match:
        for i in match.group(2).split():
            inst_type.append(i)
    # OPN
    match = opn_pattern.search(line)
    if match:
        opn.append(match.group(1))
    # OPN Intercompany
    match = opn_ic_pattern.search(line)
    if match:
        opn_ic.append(match.group(1))
    # Hull Number
    match = hull_num_pattern.search(line)
    if match:
        if match.group(1) == "<Enter>":
            hull_num.append("")
        else:
            hull_num.append(match.group(1))
    # Class Number
    match = class_num_pattern.search(line)
    if match:
        if match.group(1) == "<Enter>":
            class_num.append("")
        else:
            class_num.append(match.group(1))
    # VID
    match = vid_pattern.search(line)
    if match:
        if match.group(1) == "<Enter>":
            vid.append("")
        else:
            vid.append(match.group(1))
    # Start Date
    match = start_date_pattern.search(line)
    if match:
        start_date.append(datetime.strptime(match.group(1), '%d-%b-%y').date())
    # Completion Date
    match = compl_date_pattern.search(line)
    if match:
        compl_date.append(datetime.strptime(match.group(1), '%d-%b-%y').date())
    # Installment Amount
    match = inst_amount_pattern.search(line)
    if match:
        for i in match.group(1).split():
            if i != "<Enter>":
                temp_inst_amount.append(float(i.replace(",", "")))
            else:
                temp_inst_amount.append(0)
    # Installment Date
    match = inst_date_pattern.search(line)
    if match:
        for i in match.group(1).split():
            if i != "<Enter>": 
                try:
                    temp_inst_date.append(datetime.strptime(i, '%d-%b-%y').date())
                    #print("Valid date format")
                except ValueError:
                    print("Invalid date format, fixing")
                    # Split the string at the hyphen followed by a digit
                    dates = re.findall(r'\d{1,2}-\D{3}-\d{2}', i)
                    # Convert the dates to the desired format
                    formatted_dates = [datetime.strptime(date, '%d-%b-%y').strftime('%d-%b-%y') for date in dates]
                    [temp_inst_date.append(datetime.strptime(j, '%d-%b-%y').date()) for j in formatted_dates]
            else:
                temp_inst_date.append("")
    # Org Name - Revenue
    match = org_info_pattern.search(line)
    if match:
        org_name.append(match.group(1))
        cost_center.append(match.group(2))
        key_member.append(match.group(3))
        ic.append(match.group(4))
        hours_estimate.append(float(match.group(5).replace(",", "")))
        non_reimb_exp.append(float(match.group(6).replace(",", "")))
        reimb_exp.append(float(match.group(7).replace(",", "")))
        cost.append(float(match.group(8).replace(",", "")))
        revenue.append(float(match.group(9).replace(",", "")))
    # Grand totals 
    match = grand_total_pattern.search(line)
    if match:
        count += 1
        org_name.append("TOTAL "+ str(count))
        cost_center.append("")
        key_member.append("")
        ic.append("")
        hours_estimate.append(float(match.group(1).replace(",", "")))
        non_reimb_exp.append(float(match.group(2).replace(",", "")))
        reimb_exp.append(float(match.group(3).replace(",", "")))
        cost.append(float(match.group(4).replace(",", "")))
        revenue.append(float(match.group(5).replace(",", "")))
    # Comments
    match = comments_pattern.search(line)
    if match:
        if match.group(1) == "<Enter>":
            comments.append("")
        else:
            comments.append(match.group(1))  

# Vessel Contract Amount
match = vca_pattern.findall(text)
if match:
    for i in match:
       vca.append(int(i.replace(",", "")))

# Set Number of Groups and Number of Vessel counts
num_groups = len(which_group)
num_vessels = len(which_vessel)

# Fix Installment Type format when necessary
if inst_type[0] == '<Auto>':
    inst_type = ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th", "11th", "12th", "13th", "14th", "15th", "16th", "17th", "18th", "19th", "20th", "21st", "22nd", "23rd", "24th", "25th", "26th"]

if 'Keel' in inst_type:
    inst_type = ["Contract Signing", "Steel Cutting", "Keel Laying", "Launching", "Delivery", "Other"] * num_vessels
    #need to modify, this is only for if there 8 vessels
    a = 0
    b = 6
    for i in range(num_vessels):
        inst_amount.extend(temp_inst_amount[a:b])
        inst_date.extend(temp_inst_date[a:b])
        a += 26
        b += 26

if inst_type[0] != 'Contract Signing':
    inst_amount = [x for x in temp_inst_amount if x != 0]
    inst_date = [x for x in temp_inst_date if x != ""]
    inst_type = inst_type[:int(len(inst_amount)/num_vessels)] * num_vessels

num_installments = len(inst_amount) / num_vessels

#create installments table
instlines = []
for i in range(len(inst_amount)):
    instlines.append(Installment_line(inst_type[i], inst_amount[i], inst_date[i]))
df0 = pd.DataFrame(instlines)

print(df0)
#df0.to_csv("installments.csv")
#df0.to_excel("installments.xlsx")

#create vessels table
vlines = []
j = 0
for i in range(num_vessels):
    vlines.append(Vessel_line(which_group[j], which_vessel[i], hull_num[i], class_num[i], vid[i], start_date[i], compl_date[i], vca[i]))
    if i < len(which_vessel) - 1 and which_vessel[i + 1] == "Vessel #1":
            j += 1
df1 = pd.DataFrame(vlines)

print(df1)
#df1.to_csv("vessels.csv")
#df1.to_excel("vessels.xlsx")

# Reset the index of df1 to the specified indices for when the installment type is CS6 
df1.reset_index(drop=True, inplace=True)
if inst_type[0] == "Contract Signing":
    df1.index = [i*6 for i in range(len(df0)//6)]
else:
    df1.index = [i*int(len(inst_amount)/num_vessels) for i in range(len(df0)//int(len(inst_amount)/num_vessels))]

df_merged = pd.concat([df0, df1], axis=1)
df_merged = df_merged[['Which_Group', 'Which_Vessel', 'Hull_Number', 'Class_Number', 'VID', 'Start_Date', 'Completion_Date', 'Vessel_Contract_Amount', 'Installment_Type', 'Installment_Amount', 'Installment_Date']]
print(df_merged)
#df_merged.to_csv("Vessel Info.csv")
#df_merged.to_excel("Vessel Info.xlsx")

df_filled = df_merged.fillna(method='ffill').copy()
print(df_filled)
#df_merged.to_csv("Vessel Info.csv")
#df_merged.to_excel("Vessel Info.xlsx")

#create org info table
orglines = []
j = 0
for i in range(len(org_name)):
    print("i:", i, " j:", j)
    orglines.append(Org_line(org_name[i], cost_center[i], key_member[i], ic[i], hours_estimate[i], non_reimb_exp[i], reimb_exp[i], cost[i], revenue[i], comments[j]))
    if "TOTAL" in org_name[i]:
        j += 1
    
df2 = pd.DataFrame(orglines)

print(df2)
#df2.to_csv("orginfo2.csv")
#df2.to_excel("orginfo2.xlsx")

eventlines = []

total1_index = 0
total2_index = 0
total3_index = 0
total4_index = 0
total5_index = 0
total6_index = 0

rows_before_total = pd.DataFrame()
num_events_per_vessel = num_installments

k = 0
for i in range(len(df_filled)):
    if (i + 1) % num_installments == 0:
        k += 1
    if df_filled.loc[i, "Which_Group"] == "Group 1":
        if 'TOTAL 1' in df2['Organization_Name'].values:
            total1_index = df2[df2['Organization_Name'] == 'TOTAL 1'].index[0]
            rows_before_total = df2.iloc[df2.index[0]:total1_index]
            for j in range(len(rows_before_total)):
                event_amounts.append(event_calculation(df_merged.loc[i, 'Installment_Amount'], rows_before_total.iloc[j, 8], df2.loc[total1_index, 'Revenue']))
                alloc_percentages.append(alloc_perc_calculation(df_merged.loc[i, 'Installment_Amount'], df2.loc[total1_index, 'Revenue']))
                eventlines.append(Event_line(df_merged.loc[i, 'Installment_Type'], 0, df_merged.loc[i, 'Installment_Date'], rows_before_total.iloc[j, 0], rows_before_total.iloc[j, 1], event_amounts[-1], 0, alloc_percentages[-1]))
            num_events_per_vessel = int(num_installments * len(rows_before_total))
            if i == k * num_installments - 1:
                sum_of_events_rev = float(sum(event_amounts[-num_events_per_vessel:]))
                updated_event_line2 = eventlines[-1]._replace(Total_Revenue = sum_of_events_rev)
                eventlines[-1] = updated_event_line2
    elif df_filled.loc[i, "Which_Group"] == "Group 2":
        if 'TOTAL 2' in df2['Organization_Name'].values:
            total2_index = df2[df2['Organization_Name'] == 'TOTAL 2'].index[0]
            rows_before_total = df2.iloc[total1_index + 1:total2_index]
            for j in range(len(rows_before_total)):
                event_amounts.append(event_calculation(df_merged.loc[i, 'Installment_Amount'], rows_before_total.iloc[j, 8], df2.loc[total2_index, 'Revenue']))
                alloc_percentages.append(alloc_perc_calculation(df_merged.loc[i, 'Installment_Amount'], df2.loc[total2_index, 'Revenue']))
                eventlines.append(Event_line(df_merged.loc[i, 'Installment_Type'], 0, df_merged.loc[i, 'Installment_Date'], rows_before_total.iloc[j, 0], rows_before_total.iloc[j, 1], event_amounts[-1], 0, alloc_percentages[-1]))
            num_events_per_vessel = int(num_installments * len(rows_before_total))
            if i == k * num_installments - 1:
                sum_of_events_rev = float(sum(event_amounts[-num_events_per_vessel:]))
                updated_event_line2 = eventlines[-1]._replace(Total_Revenue = sum_of_events_rev)
                eventlines[-1] = updated_event_line2
    elif df_filled.loc[i, "Which_Group"] == "Group 3":
        if 'TOTAL 3' in df2['Organization_Name'].values:
            total3_index = df2[df2['Organization_Name'] == 'TOTAL 3'].index[0]
            rows_before_total = df2.iloc[total2_index + 1:total3_index]
            for j in range(len(rows_before_total)):
                event_amounts.append(event_calculation(df_merged.loc[i, 'Installment_Amount'], rows_before_total.iloc[j, 8], df2.loc[total3_index, 'Revenue']))
                alloc_percentages.append(alloc_perc_calculation(df_merged.loc[i, 'Installment_Amount'], df2.loc[total3_index, 'Revenue']))
                eventlines.append(Event_line(df_merged.loc[i, 'Installment_Type'], 0, df_merged.loc[i, 'Installment_Date'], rows_before_total.iloc[j, 0], rows_before_total.iloc[j, 1], event_amounts[-1], 0, alloc_percentages[-1]))
            num_events_per_vessel = int(num_installments * len(rows_before_total))
            if i == k * num_installments - 1:
                sum_of_events_rev = float(sum(event_amounts[-num_events_per_vessel:]))
                updated_event_line2 = eventlines[-1]._replace(Total_Revenue = sum_of_events_rev)
                eventlines[-1] = updated_event_line2
    elif df_filled.loc[i, "Which_Group"] == "Group 4":
        if 'TOTAL 4' in df2['Organization_Name'].values:
            total4_index = df2[df2['Organization_Name'] == 'TOTAL 4'].index[0]
            rows_before_total = df2.iloc[total3_index + 1:total4_index]
            for j in range(len(rows_before_total)):
                event_amounts.append(event_calculation(df_merged.loc[i, 'Installment_Amount'], rows_before_total.iloc[j, 8], df2.loc[total4_index, 'Revenue']))
                alloc_percentages.append(alloc_perc_calculation(df_merged.loc[i, 'Installment_Amount'], df2.loc[total4_index, 'Revenue']))
                eventlines.append(Event_line(df_merged.loc[i, 'Installment_Type'], 0, df_merged.loc[i, 'Installment_Date'], rows_before_total.iloc[j, 0], rows_before_total.iloc[j, 1], event_amounts[-1], 0, alloc_percentages[-1]))
            num_events_per_vessel = int(num_installments * len(rows_before_total))
            if i == k * num_installments - 1:
                sum_of_events_rev = float(sum(event_amounts[-num_events_per_vessel:]))
                updated_event_line2 = eventlines[-1]._replace(Total_Revenue = sum_of_events_rev)
                eventlines[-1] = updated_event_line2
    elif df_filled.loc[i, "Which_Group"] == "Group 5":
        if 'TOTAL 5' in df2['Organization_Name'].values:
            total5_index = df2[df2['Organization_Name'] == 'TOTAL 5'].index[0]
            rows_before_total = df2.iloc[total4_index + 1:total5_index]
            for j in range(len(rows_before_total)):
                event_amounts.append(event_calculation(df_merged.loc[i, 'Installment_Amount'], rows_before_total.iloc[j, 8], df2.loc[total5_index, 'Revenue']))
                alloc_percentages.append(alloc_perc_calculation(df_merged.loc[i, 'Installment_Amount'], df2.loc[total5_index, 'Revenue']))
                eventlines.append(Event_line(df_merged.loc[i, 'Installment_Type'], 0, df_merged.loc[i, 'Installment_Date'], rows_before_total.iloc[j, 0], rows_before_total.iloc[j, 1], event_amounts[-1], 0, alloc_percentages[-1]))
            num_events_per_vessel = int(num_installments * len(rows_before_total))
            if i == k * num_installments - 1:
                sum_of_events_rev = float(sum(event_amounts[-num_events_per_vessel:]))
                updated_event_line2 = eventlines[-1]._replace(Total_Revenue = sum_of_events_rev)
                eventlines[-1] = updated_event_line2
    elif df_filled.loc[i, "Which_Group"] == "Group 6":   #could run into an issue if a project have > 6 groups but thats a later problem
        if 'TOTAL 6' in df2['Organization_Name'].values:
            total6_index = df2[df2['Organization_Name'] == 'TOTAL 6'].index[0]
            rows_before_total = df2.iloc[total5_index + 1:total6_index]
            for j in range(len(rows_before_total)):
                event_amounts.append(event_calculation(df_merged.loc[i, 'Installment_Amount'], rows_before_total.iloc[j, 8], df2.loc[total6_index, 'Revenue']))
                alloc_percentages.append(alloc_perc_calculation(df_merged.loc[i, 'Installment_Amount'], df2.loc[total6_index, 'Revenue']))
                eventlines.append(Event_line(df_merged.loc[i, 'Installment_Type'], 0, df_merged.loc[i, 'Installment_Date'], rows_before_total.iloc[j, 0], rows_before_total.iloc[j, 1], event_amounts[-1], 0, alloc_percentages[-1]))
            num_events_per_vessel = int(num_installments * len(rows_before_total))
            if i == k * num_installments - 1:
                sum_of_events_rev = float(sum(event_amounts[-num_events_per_vessel:]))
                updated_event_line2 = eventlines[-1]._replace(Total_Revenue = sum_of_events_rev)
                eventlines[-1] = updated_event_line2
    sum_of_events_inst = float(sum(event_amounts[-len(rows_before_total):]))
    updated_event_line = eventlines[-1]._replace(Total_Installment_Amount = sum_of_events_inst)
    eventlines[-1] = updated_event_line
    
df3 = pd.DataFrame(eventlines)
print(df3)

# Format the necessary numeric columns before sending to excel
df_merged['Vessel_Contract_Amount'] = df_merged['Vessel_Contract_Amount'].map('{:,.2f}'.format)
df_merged['Installment_Amount'] = df_merged['Installment_Amount'].map('{:,.2f}'.format)
df2['Estimated_Hours'] = df2['Estimated_Hours'].map('{:,.2f}'.format)
df2['Nonreimbursable_Expenses'] = df2['Nonreimbursable_Expenses'].map('{:,.2f}'.format)
df2['Reimbursable_Expenses'] = df2['Reimbursable_Expenses'].map('{:,.2f}'.format)
df2['Cost'] = df2['Cost'].map('{:,.2f}'.format)
df2['Revenue'] = df2['Revenue'].map('{:,.2f}'.format)
df3['Total_Installment_Amount'] = df3['Total_Installment_Amount'].map('{:,.2f}'.format)
df3['Event_Amount'] = df3['Event_Amount'].map('{:,.2f}'.format)
df3['Total_Revenue'] = df3['Total_Revenue'].map('{:,.2f}'.format)
df3['Allocation_Percentage'] = df3['Allocation_Percentage'].map('{:.2f}%'.format)

print(df_merged)
print(df2)
print(df3)
#df3['Total_Installment_Amount'] = df3.groupby('Milestone_or_Installment_Type')['Event_Amount'].transform('sum')
#df3['Total_Installment_Amount'] = df3.groupby(df3.index // 3)['Event_Amount'].transform('sum')
#print(df3.groupby(df3.index // 18)['Event_Amount'].transform('sum'))

#print(df3)

#print(df_merged)
#print(len(event_amounts))

#checks: need to make sure installment amount = sum of events for the installment types
#print([sum(event_amounts[i:i+3]) for i in range(0, len(event_amounts), 3)])
#        need to make sure total revenue amount = sum of events for each vessel
#print([sum(event_amounts[i:i+18]) for i in range(0, len(event_amounts), 18)])

file_name = "Request PP Breakdown.xlsx"

with pd.ExcelWriter(file_name) as writer:
    df_merged.to_excel(writer, sheet_name= "Vessels and Installments", index=False)
    df2.to_excel(writer, sheet_name= "Expenses and Revenue", index=False)
    df3.to_excel(writer, sheet_name= "Events", index=False)

print("Finished!")

#from here, create a folder and connect flow to it so that whenever one of these is made, it triggers the flow or something
#print(which_group)
#print(which_vessel)
#print(f"Number of Groups: {num_groups} Number of Vessels:  {num_vessels}")  
#print(inst_type)
#print(len(inst_type))
#print(inst_amount)
#print(len(inst_amount))
#print(inst_date)
#print(len(inst_date))
#print(opn)
#print(hull_num)
#print(class_num)
#print(vid)
#print(start_date)
#print(compl_date)
#print(vca)
#print(org_name)
#print(cost_center)
#print(key_member)
#print(ic)
#print(hours_estimate)
#print(non_reimb_exp)
#print(reimb_exp)
#print(cost)
#print(revenue)
#print(gt_hours_estimate, gt_revenue)
#print(comments)