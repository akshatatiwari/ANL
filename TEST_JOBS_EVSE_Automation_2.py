# %% [markdown]
# # JOBS EVSE Automation 2.0
# _____________________________________________________
# ### Argonne National Laboratory, Energy Systems & Infrastructure Analysis
# 
# Project: JOBS EVSE automation
# Descrption: Automating the JOBS EVSE excel codebase from VBA in Excel to Python 
#             for greater efficiency and ease for internal use.
# 
# Yue Ke,
# Akshata Tiwari

# %% [markdown]
# ## Excel Sheet Setup
# Reading in Excel files and setting up User-Input excel sheets that will contain employment outputs

# %%
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import math


# %%
# Download and read the excel sheet

# **** CHANGE FILE PATH ****

input_file_path = '/Users/akshatatiwari/Downloads/Input.xlsx'
inputs = pd.read_excel(input_file_path)
inputs


# %%
# Download and read the excel sheet

# **** CHANGE FILE PATH ****

jobs_evse_multipliers_file_path = '/Users/akshatatiwari/Downloads/JOBS EVSE Automation 2.0 - Multipliers.xlsx'
tier_mult = pd.read_excel(jobs_evse_multipliers_file_path, sheet_name = "TierEmp Mult")
type1_mult = pd.read_excel(jobs_evse_multipliers_file_path, sheet_name = "Type1Emp Muilt")
type2_mult = pd.read_excel(jobs_evse_multipliers_file_path, sheet_name = "Type2Emp Mult")



# %%
# Download and read the excel sheet

# **** CHANGE FILE PATH ****

electricity_rate_file_path = '/Users/akshatatiwari/Downloads/electricity_rate.xlsx'
elec_rate = pd.read_excel(electricity_rate_file_path)
elec_rate

# %% [markdown]
# ## Station Development Formulas
# Formulas for all EVSE Components that have to do with Station Development
# Calculating total values based on formulas and multipliers for each JOBS EVSE Component.

# %% [markdown]
# #### Equipment Components
# Includes Cable Cooling, Charger, Conduit and Cables, Trenching and Boring Labor, On-site Electrical Storage, Safety & Traffic Control, Load Center/Panels, Transformers, Meters, Misc. (mounting hardware, etc.)

# %%
# Global variable values:
# Employment deflator:
empdef = 0.83594203
station_equip_expenses = []

# %%
# Calculating total value for Station Development Employment
# NAME:         Cable Cooling (physical component)
# DESCRIPTION:  Air conditioning, refrigeration and warm air heating equipment manufacturing

# Function that will calculate cable cooling values for station development calc for: Producer, Wholesale, Shipping Margin
cable_cooling_cost2008_list = []
count = 0
def cable_cooling(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0
        if inputs["Charger Power"][i] <= 50:
            cost2008 = 0 
        elif inputs["Charger Power"][i] > 50:
            cost2008 = 500 * 4  
        
        if (count == 1):
            cable_cooling_cost2008_list.append(cost2008)

        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)


# List of margin values:
cable_cooling_PMList = [0.717282811, 0.282717189, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_cable_cooling = []
ptype1_cable_cooling = []
ptype2_cable_cooling = []
# Wholesale Margin Value:
wtier_cable_cooling = []
wtype1_cable_cooling = []
wtype2_cable_cooling = []
# Shipping Margin Value:
stier_cable_cooling = []
stype1_cable_cooling = []
stype2_cable_cooling = []

returnval = cable_cooling(cable_cooling_PMList[0], ptier_cable_cooling, ptype1_cable_cooling, ptype2_cable_cooling, 333415)
cable_cooling(cable_cooling_PMList[1], wtier_cable_cooling, wtype1_cable_cooling, wtype2_cable_cooling, 420000)
cable_cooling(cable_cooling_PMList[2], stier_cable_cooling, stype1_cable_cooling, stype2_cable_cooling, 484000)

# Appending cable_cooling_cost2008 to total station_equip_expenses list
station_equip_expenses.append(cable_cooling_cost2008_list)

print(ptier_cable_cooling)
print(ptype1_cable_cooling)
print(ptype2_cable_cooling)
print(wtier_cable_cooling)
print(wtype1_cable_cooling)
print(wtype2_cable_cooling)
print(stier_cable_cooling)
print(stype1_cable_cooling)
print(stype2_cable_cooling)
print(cable_cooling_cost2008_list)


# %%
# Calculating total value for Station Development Employment
# NAME:         Charger (physical component)
# DESCRIPTION:  Other industrial machinery manufacturing

# Function that will calculate charger values for station development calc for: Producer, Wholesale, Shipping Margin
charger_cost2008_list = []
count = 0
def charger(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0
        
        if inputs["Charger Power"][i] <= 6.6:
            if inputs["Number of chargers per station"][i] == 1:
                cost2008 = 200 
            else:
                cost2008 = 530
        elif 6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2:
            if inputs["Number of chargers per station"][i] == 1:
                cost2008 = 900 
            else:
                cost2008 = 4900 
        elif inputs["Charger Power"][i] == 50:
            cost2008 = (27900) 
        elif inputs["Charger Power"][i] == 150:
            cost2008 = (87800)
        elif inputs["Charger Power"][i] == 350:
            cost2008 = (140000) 
        cost2008 = cost2008 * inputs["Number of chargers per station"][i]   
        
        if (count == 1):
            charger_cost2008_list.append(cost2008)

        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)


# List of margin values:
charger_PMList = [0.857113566, 0.142886434, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_charger = []
ptype1_charger = []
ptype2_charger = []
# Wholesale Margin Value:
wtier_charger = []
wtype1_charger = []
wtype2_charger = []
# Shipping Margin Value:
stier_charger = []
stype1_charger = []
stype2_charger = []

returnval = charger(charger_PMList[0], ptier_charger, ptype1_charger, ptype2_charger, "33329A")
charger(charger_PMList[1], wtier_charger, wtype1_charger, wtype2_charger, 420000)
charger(charger_PMList[2], stier_charger, stype1_charger, stype2_charger, 484000)

# Appending charger_cost2008 to total station_equip_expenses list
station_equip_expenses.append(charger_cost2008_list)

print(ptier_charger)
print(ptype1_charger)
print(ptype2_charger)
print(wtier_charger)
print(wtype1_charger)
print(wtype2_charger)
print(stier_charger)
print(stype1_charger)
print(stype2_charger)
print(charger_cost2008_list)

# %%
# Calculating total value for Station Development Employment
# NAME:         Conduit Cables (physical component)
# DESCRIPTION:  Communication and energy wire and cable manufacturing

# Function that will calculate Conduit Cables values for station development calc for: Producer, Wholesale, Shipping Margin
conduit_cables_cost2008_list = []
count = 0
def conduit_cables(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0
        
        if inputs["Charger Power"][i] <= 6.6:
            cost2008 = 7
        elif inputs["Charger Power"][i] > 6.6 and inputs["Charger Power"][i] <= 10.9:
            cost2008 = 20
        else:
            cost2008 = 25
       
        cost2008 = cost2008 * 75  

        if (count == 1):
            conduit_cables_cost2008_list.append(cost2008) 

        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)


# List of margin values:
conduit_cables_PMList = [0.793550332, 0.206449668, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_conduit_cables = []
ptype1_conduit_cables = []
ptype2_conduit_cables = []
# Wholesale Margin Value:
wtier_conduit_cables = []
wtype1_conduit_cables = []
wtype2_conduit_cables = []
# Shipping Margin Value:
stier_conduit_cables = []
stype1_conduit_cables = []
stype2_conduit_cables = []

returnval = conduit_cables(conduit_cables_PMList[0], ptier_conduit_cables, ptype1_conduit_cables, ptype2_conduit_cables, 335920)
conduit_cables(conduit_cables_PMList[1], wtier_conduit_cables, wtype1_conduit_cables, wtype2_conduit_cables, 420000)
conduit_cables(conduit_cables_PMList[2], stier_conduit_cables, stype1_conduit_cables, stype2_conduit_cables, 484000)

# Appending conduit_cables_cost2008 to total station_equip_expenses list
station_equip_expenses.append(conduit_cables_cost2008_list)

print(ptier_conduit_cables)
print(ptype1_conduit_cables)
print(ptype2_conduit_cables)
print(wtier_conduit_cables)
print(wtype1_conduit_cables)
print(wtype2_conduit_cables)
print(stier_conduit_cables)
print(stype1_conduit_cables)
print(stype2_conduit_cables)
print(conduit_cables_cost2008_list)

# %%
# Calculating total value for Station Development Employment
# NAME:         Trenching and Boring Labor
# DESCRIPTION:  Nonresidential structures

# Function that will calculate Trenching and Boring Labor values for station development calc for: Producer, Wholesale Margin
trenching_cost2008_list = []
count = 0
def trenching(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0
        
        if inputs["Charger Power"][i] <= 6.6:
            cost2008 = 0
        else:
            cost2008 = 80
       
        cost2008 = cost2008 * 75 
        
        if (count == 1):
            trenching_cost2008_list.append(cost2008)
  
        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)

# List of margin values:
trenching_PMList = [1, 1]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_trenching = []
ptype1_trenching = []
ptype2_trenching = []
# Wholesale Margin Value:
wtier_trenching = []
wtype1_trenching = []
wtype2_trenching = []

trenching(trenching_PMList[0], ptier_trenching, ptype1_trenching, ptype2_trenching, "2332E0")
trenching(trenching_PMList[1], wtier_trenching, wtype1_trenching, wtype2_trenching, 420000)

# Appending trenching_cost2008 to total station_equip_expenses list
station_equip_expenses.append(trenching_cost2008_list)

print(ptier_trenching)
print(ptype1_trenching)
print(ptype2_trenching)
print(wtier_trenching)
print(wtype1_trenching)
print(wtype2_trenching)
print(trenching_cost2008_list)

# %%
# Calculating total value for Station Development Employment
# NAME:         On-site Electrical Storage
# DESCRIPTION:  Storage battery manufacturing

# Function that will calculate Electrical Storage values for station development calc for: Producer, Wholesale, Shipping Margin
electrical_storage_cost2008_list = []
count = 0
def electrical_storage(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0
        
        if inputs["Charger Power"][i] == 50:
            cost2008 = 83200
        elif inputs["Charger Power"][i] == 150:
            cost2008 = 332800
        elif inputs["Charger Power"][i] == 350:
            cost2008 = 388266.6667
        
        if (count == 1):
            if str(inputs["yes/no to include onsite storage costs"][i]).strip().lower() == "yes":
                electrical_storage_cost2008_list.append(cost2008) 
            else:
                electrical_storage_cost2008_list.append(0)
       
        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]
        
        if str(inputs["yes/no to include onsite storage costs"][i]).strip().lower() == "yes":
            tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
            type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
            type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        else:
            tierlist[i] = 0
            type1list[i] = 0
            type2list[i] = 0

# List of margin values:
electrical_storage_PMList = [0.709280709, 0.290719291, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_electrical_storage = []
ptype1_electrical_storage = []
ptype2_electrical_storage = []
# Wholesale Margin Value:
wtier_electrical_storage = []
wtype1_electrical_storage = []
wtype2_electrical_storage = []
# Shipping Margin Value:
stier_electrical_storage = []
stype1_electrical_storage = []
stype2_electrical_storage = []

returnval = electrical_storage(electrical_storage_PMList[0], ptier_electrical_storage, ptype1_electrical_storage, ptype2_electrical_storage, 335911)
electrical_storage(electrical_storage_PMList[1], wtier_electrical_storage, wtype1_electrical_storage, wtype2_electrical_storage, 420000)
electrical_storage(electrical_storage_PMList[2], stier_electrical_storage, stype1_electrical_storage, stype2_electrical_storage, 484000)

# Appending electrical_storage_cost2008 to total station_equip_expenses list
station_equip_expenses.append(electrical_storage_cost2008_list)

print(ptier_electrical_storage)
print(ptype1_electrical_storage)
print(ptype2_electrical_storage)
print(wtier_electrical_storage)
print(wtype1_electrical_storage)
print(wtype2_electrical_storage)
print(stier_electrical_storage)
print(stype1_electrical_storage)
print(stype2_electrical_storage)
print(electrical_storage_cost2008_list)


# %%
# Calculating total value for Station Development Employment
# NAME:         Safety & Traffic Control
# DESCRIPTION:  Transportation structures and highways and streets

# Function that will calculate Safety & Traffic Control values for station development calc for: Producer, Wholesale, Shipping Margin
safety_traffic_cost2008_list = []
count = 0
def safety_traffic(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0
        if inputs["Number of chargers per station"][i] == 3 and inputs["Charger Power"][i] <= 19.2:
            cost2008 = 1000 
        else:
            cost2008 = 3000 
       
        
        if (count == 1):
            safety_traffic_cost2008_list.append(cost2008) 

        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)

# List of margin values:
safety_traffic_PMList = [1, 0, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_safety_traffic = []
ptype1_safety_traffic = []
ptype2_safety_traffic = []
# Wholesale Margin Value:
wtier_safety_traffic = []
wtype1_safety_traffic = []
wtype2_safety_traffic = []
# Shipping Margin Value:
stier_safety_traffic = []
stype1_safety_traffic = []
stype2_safety_traffic = []

returnval = safety_traffic(safety_traffic_PMList[0], ptier_safety_traffic, ptype1_safety_traffic, ptype2_safety_traffic, "2332F0")
safety_traffic(safety_traffic_PMList[1], wtier_safety_traffic, wtype1_safety_traffic, wtype2_safety_traffic, 420000)
safety_traffic(safety_traffic_PMList[2], stier_safety_traffic, stype1_safety_traffic, stype2_safety_traffic, 484000)

# Appending safety_traffic_cost2008 to total station_equip_expenses list
station_equip_expenses.append(safety_traffic_cost2008_list)

print(ptier_safety_traffic)
print(ptype1_safety_traffic)
print(ptype2_safety_traffic)
print(wtier_safety_traffic)
print(wtype1_safety_traffic)
print(wtype2_safety_traffic)
print(stier_safety_traffic)
print(stype1_safety_traffic)
print(stype2_safety_traffic)
print(safety_traffic_cost2008_list)



# %%
# Calculating total value for Station Development Employment
# NAME:         Load Center/Panels
# DESCRIPTION:  Switchgear and switchboard apparatus manufacturing

# Function that will calculate Safety & Traffic Control values for station development calc for: Producer, Wholesale, Shipping Margin
load_center_cost2008_list = []
count = 0
def load_center(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0

        if inputs["Charger Power"][i] <= 6.6:
            cost2008 = 5 
        elif 6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2:
            cost2008 = 10 
        else:
            cost2008 = 40 

        if inputs["Number of chargers per station"][i] == 3:
            cost2008 = cost2008 * 6
        elif inputs["Number of chargers per station"][i] == 4 or inputs["Number of chargers per station"][i] == 2:
            cost2008 = cost2008 * 4
        
        
        if (count == 1):
            load_center_cost2008_list.append(cost2008) 

        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
   


# List of margin values:
load_center_PMList = [0.785214135, 0.214785865, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_load_center = []
ptype1_load_center = []
ptype2_load_center = []
# Wholesale Margin Value:
wtier_load_center = []
wtype1_load_center = []
wtype2_load_center = []
# Shipping Margin Value:
stier_load_center = []
stype1_load_center = []
stype2_load_center = []

returnval = load_center(load_center_PMList[0], ptier_load_center, ptype1_load_center, ptype2_load_center, 335313)
load_center(load_center_PMList[1], wtier_load_center, wtype1_load_center, wtype2_load_center, 420000)
load_center(load_center_PMList[2], stier_load_center, stype1_load_center, stype2_load_center, 484000)

# Appending load_center_cost2008 to total station_equip_expenses list
station_equip_expenses.append(load_center_cost2008_list)

print(ptier_load_center)
print(ptype1_load_center)
print(ptype2_load_center)
print(wtier_load_center)
print(wtype1_load_center)
print(wtype2_load_center)
print(stier_load_center)
print(stype1_load_center)
print(stype2_load_center)
print(load_center_cost2008_list)



# %%
# Calculating total value for Station Development Employment
# NAME:         Transformers
# DESCRIPTION:  Power, distribution, and specialty transformer manufacturing

# Function that will calculate Transformers values for station development calc for: Producer, Wholesale, Shipping Margin
transformers_cost2008_list = []
count = 0

def transformers(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1
    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)

        # Find tier employment per charger:
        cost2008 = 0

        if inputs["Charger Power"][i] <= 6.6:
            cost2008 = 0
        elif 6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 10.9:
            cost2008 = 11219.51132
        elif inputs["Charger Power"][i] == 50:
            cost2008 = 16446.835
        elif inputs["Charger Power"][i] == 150:
            cost2008 = 37865.26
        elif inputs["Charger Power"][i] == 350:
            cost2008 = 42918.94
        
        if (count == 1):
            if str(inputs["yes/no to include transformer costs"][i]).strip().lower() == "yes":
                transformers_cost2008_list.append(cost2008) 
            else:
                transformers_cost2008_list.append(0)


        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]

        if str(inputs["yes/no to include transformer costs"][i]).strip().lower() == "yes":
            tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
            type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
            type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        else:
            tierlist[i] = 0
            type1list[i] = 0
            type2list[i] = 0
        

# List of margin values:
transformers_PMList = [0.79003035, 0.20996965, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_transformers = []
ptype1_transformers = []
ptype2_transformers = []
# Wholesale Margin Value:
wtier_transformers = []
wtype1_transformers = []
wtype2_transformers = []
# Shipping Margin Value:
stier_transformers = []
stype1_transformers = []
stype2_transformers = []

returnval = transformers(transformers_PMList[0], ptier_transformers, ptype1_transformers, ptype2_transformers, 335311)
transformers(transformers_PMList[1], wtier_transformers, wtype1_transformers, wtype2_transformers, 420000)
transformers(transformers_PMList[2], stier_transformers, stype1_transformers, stype2_transformers, 484000)

# Appending transformers_cost2008 to total station_equip_expenses list
station_equip_expenses.append(transformers_cost2008_list)

print(ptier_transformers)
print(ptype1_transformers)
print(ptype2_transformers)
print(wtier_transformers)
print(wtype1_transformers)
print(wtype2_transformers)
print(stier_transformers)
print(stype1_transformers)
print(stype2_transformers)
print(transformers_cost2008_list)


# %%
# Calculating total value for Station Development Employment
# NAME:         Meters
# DESCRIPTION:  Electrical Meters

# Function that will calculate Meters values for station development calc for: Producer, Wholesale, Shipping Margin
meters_cost2008_list = []
count = 0
def meters(marginval, tierlist, type1list, type2list, code):
    global count
    count = count + 1

    # Finding total Tier, Type I, Type II Employment for each run:
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)
       
        # Find tier employment per charger:
        cost2008 = 0

        if inputs["Charger Power"][i] <= 6.6:
            cost2008 = 0
        else:
            cost2008 = 2000
        
        if inputs["Charger Power"][i] == 150 or inputs["Charger Power"][i] == 350:
            cost2008 = cost2008 * 4
        else:
            cost2008 = cost2008 * 6
        
        if (count == 1):
            if str(inputs["yes/no to include meters"][i]).strip().lower() == "yes":
                meters_cost2008_list.append(cost2008) 
            else:
                meters_cost2008_list.append(0) 

        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]
        
        if str(inputs["yes/no to include meters"][i]).strip().lower() == "yes":
            tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
            type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
            type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        else:
            tierlist[i] = 0
            type1list[i] = 0
            type2list[i] = 0


# List of margin values:
meters_PMList = [0.878689182, 0.121310818, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_meters = []
ptype1_meters = []
ptype2_meters = []
# Wholesale Margin Value:
wtier_meters = []
wtype1_meters = []
wtype2_meters = []
# Shipping Margin Value:
stier_meters = []
stype1_meters = []
stype2_meters = []

returnval = meters(meters_PMList[0], ptier_meters, ptype1_meters, ptype2_meters, 334515)
meters(meters_PMList[1], wtier_meters, wtype1_meters, wtype2_meters, 420000)
meters(meters_PMList[2], stier_meters, stype1_meters, stype2_meters, 484000)

# Appending meters_cost2008 to total station_equip_expenses list
station_equip_expenses.append(meters_cost2008_list)

print(ptier_meters)
print(ptype1_meters)
print(ptype2_meters)
print(wtier_meters)
print(wtype1_meters)
print(wtype2_meters)
print(stier_meters)
print(stype1_meters)
print(stype2_meters)
print(meters_cost2008_list)



# %%
# Calculating total value for Station Development Employment
# NAME:         Misc. (mounting hardware, etc.)
# DESCRIPTION:  Other concrete product manufacturing

# Function that will calculate Transformers values for station development calc for: Producer, Wholesale, Shipping Margin
misc_cost2008_list = []
count = 0
def misc(marginval, tierlist, type1list, type2list, code):
    # Finding total Tier, Type I, Type II Employment for each run:
    global count
    count = count + 1
    for i in range(len(inputs)):
        tierlist.append(0)
        type1list.append(0)
        type2list.append(0)
        
        # Find tier employment per charger:
        cost2008 = 0

        if inputs["Number of chargers per station"][i] == 1:
            cost2008 = 100
        elif inputs["[Average] sessions per month"][i] == 114:
            cost2008 = 2000 * 3
        elif inputs["Charger Power"][i] <= 50:
            cost2008 = 1000 * 3
        elif inputs["Charger Power"][i] == 150:
            cost2008 = 2000 * 4
        elif inputs["Charger Power"][i] == 350:
            cost2008 = 2000 * 2

        if (count == 1):
            misc_cost2008_list.append(cost2008) 
        
        cost2008 = cost2008 * empdef 
        row_num = tier_mult[tier_mult['Code'] == code].index[0]
        tierlist[i] = ((cost2008 * marginval * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type1list[i] = ((cost2008 * marginval * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
        type2list[i] = ((cost2008 * marginval * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)


# List of margin values:
misc_PMList = [0.886582374, 0.113417626, 0.03]

# List for tier, type I, and type II Employment for each margin value for cable cooling:
# Producer Margin Value:
ptier_misc = []
ptype1_misc = []
ptype2_misc = []
# Wholesale Margin Value:
wtier_misc = []
wtype1_misc = []
wtype2_misc = []
# Shipping Margin Value:
stier_misc = []
stype1_misc = []
stype2_misc = []

misc(misc_PMList[0], ptier_misc, ptype1_misc, ptype2_misc, 327390)
misc(misc_PMList[1], wtier_misc, wtype1_misc, wtype2_misc, 420000)
misc(misc_PMList[2], stier_misc, stype1_misc, stype2_misc, 484000)

# Appending misc_cost2008 to total station_equip_expenses list
station_equip_expenses.append(misc_cost2008_list)

print(ptier_misc)
print(ptype1_misc)
print(ptype2_misc)
print(wtier_misc)
print(wtype1_misc)
print(wtype2_misc)
print(stier_misc)
print(stype1_misc)
print(stype2_misc)

print(misc_cost2008_list)



# %% [markdown]
# #### Non-Equipment Components
# Includes: Equipment Installation, Site Prep & Construction, Electrical Infrastrastructure & Make Ready, Engineering & Design, Permitting,	Contingencies - Install, Contingencies - Site Prep & Construction, Contingencies - Electrical

# %%
print(station_equip_expenses)

# %%
# Code to calculate total cost of station equipment expenses (components calculated above):
station_equip_expenses_total = []
total_contingencies = []

station_equip_len = len(station_equip_expenses[0])

for i in range(station_equip_len):
    station_equip_expenses_total.append(0)
    for j in station_equip_expenses:
        station_equip_expenses_total[i] += (j[i])

# Creating global variables:
for i in station_equip_expenses_total:
    total_contingencies.append(i * 0.05)
print(total_contingencies)
print(station_equip_expenses_total)

# %%
# NAME:         Equipment Installation
# DESCRIPTION:  Nonresidential structures

tier_einstall = []
type1_einstall = []
type2_einstall = []

for i in range(len(inputs)):
    tier_einstall.append(0)
    type1_einstall.append(0)
    type2_einstall.append(0)
    cost2008 = (station_equip_expenses_total[i] * 0.3) + (total_contingencies[i] * 0.125)
    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_einstall[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_einstall[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_einstall[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_einstall)
print(type1_einstall)
print(type2_einstall)

# %%
# NAME:         Site Prep & Construction
# DESCRIPTION:  Nonresidential structures

tier_siteprep = []
type1_siteprep = []
type2_siteprep = []

for i in range(len(inputs)):
    tier_siteprep.append(0)
    type1_siteprep.append(0)
    type2_siteprep.append(0)

    cost2008 = (station_equip_expenses_total[i] * 0.095) + (total_contingencies[i] * 0.75) + ((station_equip_expenses[3][0])/75)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_siteprep[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_siteprep[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_siteprep[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_siteprep)
print(type1_siteprep)
print(type2_siteprep)

# %%
# NAME:         Electrical Infrastrastructure & Make Ready
# DESCRIPTION:  Power, distribution, and specialty transformer manufacturing

tier_einfra = []
type1_einfra = []
type2_einfra = []

for i in range(len(inputs)):
    tier_einfra.append(0)
    type1_einfra.append(0)
    type2_einfra.append(0)

    cost2008 = (station_equip_expenses_total[i] * 0.3) + (total_contingencies[i] * 0.125)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == 335311].index[0]
    
    tier_einfra[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_einfra[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_einfra[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_einfra)
print(type1_einfra)
print(type2_einfra)

# %%
# NAME:         Engineering & Design
# DESCRIPTION:  Architectural, engineering, and related services

tier_eng = []
type1_eng = []
type2_eng = []

for i in range(len(inputs)):
    tier_eng.append(0)
    type1_eng.append(0)
    type2_eng.append(0)

    cost2008 = (station_equip_expenses_total[i] * 0.195)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == 541300].index[0]
    
    tier_eng[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_eng[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_eng[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_eng)
print(type1_eng)
print(type2_eng)

# %%
# NAME:         Permitting
# DESCRIPTION:  Architectural, engineering, and related services

tier_permit = []
type1_permit = []
type2_permit = []

for i in range(len(inputs)):
    tier_permit.append(0)
    type1_permit.append(0)
    type2_permit.append(0)

    cost2008 = (station_equip_expenses_total[i] * 0.03)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == 541300].index[0]
    
    tier_permit[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_permit[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_permit[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_permit)
print(type1_permit)
print(type2_permit)

# %%
# NAME:         Contingencies - Install
# DESCRIPTION:  Nonresidential structures

tier_cinstal = []
type1_cinstal = []
type2_cinstal = []

for i in range(len(inputs)):
    tier_cinstal.append(0)
    type1_cinstal.append(0)
    type2_cinstal.append(0)

    cost2008 = (total_contingencies[i] * 0.125)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_cinstal[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_cinstal[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_cinstal[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_cinstal)
print(type1_cinstal)
print(type2_cinstal)

# %%
# NAME:         Contingencies - Site Prep & Construction
# DESCRIPTION:  Nonresidential structures

tier_csiteprep = []
type1_csiteprep = []
type2_csiteprep = []

for i in range(len(inputs)):
    tier_csiteprep.append(0)
    type1_csiteprep.append(0)
    type2_csiteprep.append(0)

    cost2008 = (total_contingencies[i] * 0.75)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_csiteprep[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_csiteprep[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_csiteprep[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_csiteprep)
print(type1_csiteprep)
print(type2_csiteprep)

# %%
# NAME:         Contingencies - Electrical
# DESCRIPTION:  Power, distribution, and specialty transformer manufacturing

tier_celec = []
type1_celec = []
type2_celec = []


for i in range(len(inputs)):
    tier_celec.append(0)
    type1_celec.append(0)
    type2_celec.append(0)

    cost2008 = (total_contingencies[i] * 0.125)

    cost2008 = cost2008 * empdef 
    row_num = tier_mult[tier_mult['Code'] == 335311].index[0]
    
    tier_celec[i] = ((cost2008 * (tier_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type1_celec[i] = ((cost2008 * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000)
    type2_celec[i] = ((cost2008 * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000)
        

print(tier_celec)
print(type1_celec)
print(type2_celec)

# %% [markdown]
# ## Station Development Tables
# Finding the total values for each row using station development calc components.

# %% [markdown]
# #### Station Development Table

# %% [markdown]
# ##### Performing calculations for Stations Development Employment Table

# %%
# Creating Row and Column headings that will be used in each dataframe:
#   ROW HEADINGS:    Contains RUN #
#   COLUMN HEADINGS: Contains total values that will be calculated - Tier, Type I, Type II, ...

row_headings_dev = []
for i in range(len(inputs)):
    row_headings_dev.append("RUN " + str(i + 1))

column_total_headings_dev = ["Direct Employment", "Indirect Employment", "Induced Employment", "Total Employment"]

print(row_headings_dev)
print(column_total_headings_dev)

# %%
# Calculating totals for all EVSE components:

tier_total = []
type1_total = []
type2_total = []
direct_emp_total = []
indirect_emp_total = []
induced_emp_total = []
station_dev_totals = []


for i in range(len(inputs)):
    # PART 1:
    # Equipment component totals:

    if inputs["Geography"][i] != "USA-National":
        ptier = ptier_trenching[i]
        ptype1 = ptype1_trenching[i]
        ptype2 = ptype2_trenching[i]
        wtier = wtier_trenching[i]
        wtype1 = wtype1_trenching[i]
        wtype2 = wtype2_trenching[i]
        stier = 0
        stype1 = 0
        stype2 = 0
    else:
        ptier = ptier_cable_cooling[i] + ptier_charger[i] + ptier_conduit_cables[i] + ptier_trenching[i] + ptier_electrical_storage[i] + ptier_safety_traffic[i] + ptier_load_center[i] + ptier_transformers[i] + ptier_meters[i] + ptier_misc[i]
        ptype1 = ptype1_cable_cooling[i] + ptype1_charger[i] + ptype1_conduit_cables[i] + ptype1_trenching[i] + ptype1_electrical_storage[i] + ptype1_safety_traffic[i] + ptype1_load_center[i] + ptype1_transformers[i] + ptype1_meters[i] + ptype1_misc[i]
        ptype2 = ptype2_cable_cooling[i] + ptype2_charger[i] + ptype2_conduit_cables[i] + ptype2_trenching[i] + ptype2_electrical_storage[i] + ptype2_safety_traffic[i] + ptype2_load_center[i] + ptype2_transformers[i] + ptype2_meters[i] + ptype2_misc[i]

        wtier = wtier_cable_cooling[i] + wtier_charger[i] + wtier_conduit_cables[i] + wtier_trenching[i] + wtier_electrical_storage[i] + wtier_safety_traffic[i] + wtier_load_center[i] + wtier_transformers[i] + wtier_meters[i] + wtier_misc[i]
        wtype1 = wtype1_cable_cooling[i] + wtype1_charger[i] + wtype1_conduit_cables[i] + wtype1_trenching[i] + wtype1_electrical_storage[i] + wtype1_safety_traffic[i] + wtype1_load_center[i] + wtype1_transformers[i] + wtype1_meters[i] + wtype1_misc[i]
        wtype2 = wtype2_cable_cooling[i] + wtype2_charger[i] + wtype2_conduit_cables[i] + wtype2_trenching[i] + wtype2_electrical_storage[i] + wtype2_safety_traffic[i] + wtype2_load_center[i] + wtype2_transformers[i] + wtype2_meters[i] + wtype2_misc[i]

        stier = stier_cable_cooling[i] + stier_charger[i] + stier_conduit_cables[i] + stier_electrical_storage[i] + stier_safety_traffic[i] + stier_load_center[i] + stier_transformers[i] + stier_meters[i] + stier_misc[i]
        stype1 = stype1_cable_cooling[i] + stype1_charger[i] + stype1_conduit_cables[i] + stype1_electrical_storage[i] + stype1_safety_traffic[i] + stype1_load_center[i] + stype1_transformers[i] + stype1_meters[i] + stype1_misc[i]
        stype2 = stype2_cable_cooling[i] + stype2_charger[i] + stype2_conduit_cables[i] + stype2_electrical_storage[i] + stype2_safety_traffic[i] + stype2_load_center[i] + stype2_transformers[i] + stype2_meters[i] + stype2_misc[i]
 
    # Non-Equipment component totals:
    netier = tier_einstall[i] + tier_siteprep[i] + tier_einfra[i] + tier_eng[i] + tier_permit[i] + tier_cinstal[i] + tier_csiteprep[i] + tier_celec[i]
    netype1 = type1_einstall[i] + type1_siteprep[i] + type1_einfra[i] + type1_eng[i] + type1_permit[i] + type1_cinstal[i] + type1_csiteprep[i] + type1_celec[i]
    netype2 = type2_einstall[i] + type2_siteprep[i] + type2_einfra[i] + type2_eng[i] + type2_permit[i] + type2_cinstal[i] + type2_csiteprep[i] + type2_celec[i]

    # PART 2:
    # Calculating total Tier, Type I, Type II Employment totals:
    tier_total.append((ptier + wtier + stier + netier) * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))
    type1_total.append((ptype1 + wtype1 + stype1 + netype1) * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))
    type2_total.append((ptype2 + wtype2 + stype2 + netype2) * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))
    
   
    
    # Calculating Direct, Indirect, and Induced Employment totals:
    direct_emp_total.append(tier_total[i])
    if str(inputs["yes/no to include indirect effects"][i]).strip().lower() == "yes":
        indirect_emp_total.append(type1_total[i] - tier_total[i])
    else:
        indirect_emp_total.append(0)
    induced_emp_total.append(type2_total[i] - type1_total[i])
    
    
# Putting together in total table:
for i in range(len(type1_total)):
    curr_row = []
    curr_row.append(type1_total[i])
    curr_row.append(induced_emp_total[i])
    curr_row.append(type2_total[i])
    station_dev_totals.append(curr_row)

    
print(tier_total)
print(type1_total)
print(type2_total)
print(direct_emp_total)
print(indirect_emp_total)
print(induced_emp_total)
print(station_dev_totals)


# %% [markdown]
# Creating Dataframe with total output values

# %%
# Creating Row and Column headings that will be used in each dataframe:
#   ROW HEADINGS:    Contains RUN #
#   COLUMN HEADINGS: Contains total values that will be calculated - Tier, Type I, Type II, ...

row_headings_statdev = []
for i in range(len(inputs)):
    row_headings_statdev.append("RUN " + str(i + 1))

column_total_headings_statdev = ["Supply Chain Employment", "Induced Employment", "Total Employment"]

print(row_headings_statdev)
print(column_total_headings_statdev)

# %% [markdown]
# ##### Final Table:

# %%
statdev_df = pd.DataFrame(station_dev_totals, index = row_headings_statdev, columns = column_total_headings_statdev)
statdev_df

# %% [markdown]
# #### Civil Construction Employment Table:
# Includes Equipment Installation and Site Prep & Construction and takes total values for direct, indirect, and induced employment

# %% [markdown]
# ##### Performing calculations for Civil Construction Employment

# %%
civil_const_emp = []
civil_const_direct = []
civil_const_indirect = []
civil_const_induced = []
civil_const_totals = []

for i in range(len(inputs)):
    direct = 0
    indirect = 0
    induced = 0
    # Calculating direct employment:
    direct = tier_einstall[i] + tier_siteprep[i] 
    civil_const_direct.append(direct * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

    # Calculating indirect employment:
    if str(inputs["yes/no to include indirect effects"][i]).strip().lower() == "yes":
        indirect = (type1_einstall[i] + type1_siteprep[i]) - (direct)
    else:
        indirect = 0
    civil_const_indirect.append(indirect * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

    # Calculating induced employment:
    induced = (type2_einstall[i] + type2_siteprep[i]) - (type1_einstall[i] + type1_siteprep[i]) 
    civil_const_induced.append(induced * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

    # Calculating total and appending to civil_const_emp list per row:
    civil_const_emp.append((direct + indirect + induced) * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

# Appending total values to row for dataframe
for i in range(len(civil_const_direct)):
    curr_row = []
    curr_row.append(civil_const_direct[i])
    curr_row.append(civil_const_indirect[i])
    curr_row.append(civil_const_induced[i])
    curr_row.append(civil_const_emp[i])
    civil_const_totals.append(curr_row)

# Verify output
print(civil_const_direct)
print(civil_const_indirect)
print(civil_const_induced)
print(civil_const_emp)
print(civil_const_totals)



# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
civil_const_df = pd.DataFrame(civil_const_totals, index = row_headings_dev, columns = column_total_headings_dev)
civil_const_df

# %% [markdown]
# #### Electrical Construction Employment
# Includes Electrical Infrastrastructure & Make Ready and takes total values for direct, indirect, and induced employment

# %% [markdown]
# ##### Performing calculations for Electrical Construction Employment

# %%
# NOTE: Does not include induced costs in actual tool (typo - Station Operation Calc not Station Development Calc)


elec_const_emp = []
elec_const_direct = []
elec_const_indirect = []
elec_const_induced = []
elec_const_totals = []

for i in range(len(inputs)):
    # Calculating direct employment:
    direct =  tier_einfra[i] 
    elec_const_direct.append(direct * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

    # Calculating indirect employment:
    if str(inputs["yes/no to include indirect effects"][i]).strip().lower() == "yes":
        indirect = (type1_einfra[i]) - (direct)
    else:
        indirect = 0
    elec_const_indirect.append(indirect * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

    # Calculating induced employment:
    induced = (type2_einfra[i]) - (type1_einfra[i]) 
    elec_const_induced.append(induced * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i]))

    # Calculating total and appending to civil_const_emp list per row:
    elec_const_emp.append((direct + indirect + induced)  * inputs["Number of years for analysis"][i] * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])) #### TOOK OUT INDUCED TO CHECK RESULTS, PUT BACK IN AFTER

# Appending total values to row for dataframe
for i in range(len(elec_const_direct)):
    curr_row = []
    curr_row.append(elec_const_direct[i])
    curr_row.append(elec_const_indirect[i])
    curr_row.append(elec_const_induced[i])
    curr_row.append(elec_const_emp[i])
    elec_const_totals.append(curr_row)

    
print(elec_const_direct)
print(elec_const_indirect)
print(elec_const_induced)
print(elec_const_emp)
print(elec_const_totals)

# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
elec_const_df = pd.DataFrame(elec_const_totals, index = row_headings_dev, columns = column_total_headings_dev)
elec_const_df

# %% [markdown]
# ## Station Operation Formulas and Tables
# Formulas for all EVSE Components that have to do with Station Operation
# Calculating total values based on formulas and multipliers for each JOBS EVSE Component.

# %%
# Creating Row and Column headings that will be used in each dataframe:
#   ROW HEADINGS:    Contains RUN #
#   COLUMN HEADINGS: Contains total values that will be calculated - Tier, Type I, Type II, ...

row_headings_ops = []
for i in range(len(inputs)):
    row_headings_ops.append("RUN " + str(i + 1))

column_total_headings_ops = ["Supply Chain Employment", "Induced Employment", "Total Employment"]

print(row_headings_ops)
print(column_total_headings_ops)

# %%
# Creating a dictionary that maps inputs["[Average] amount of kWh dispensed per charge session"][i] to more precise values:
kWh_dict = {6.6: 6.60000,
            14.3: 14.29804,
            25.2: 25.16471,
            52.9: 52.9, 
            31.6: 31.6, 
            17.5: 17.5}

kWh_dict[14.3]


# %% [markdown]
# #### Electriticy Sector Employment
# Includes Electrical Infrastrastructure & Make Ready and takes total values for supply chain and induced employment

# %% [markdown]
# ##### Test

# %%
ans = 172126.59 * 3.57045E-07 
sum = 0
print(ans * 2)
total = 0
for i in range(5):
    sum = ans * (i + 1)
    print(sum)
    total += sum

print(total)
print("this")
print(ans * 5)
for i in range(5):
    total = total + (ans * 5)
print(total)

# %%
ans = 152*52.9 * 12
ans * 2.9675E-07 

# %% [markdown]
# ##### Performing calculations for Electricity Sector Employment

# %%
# NOTE: CHANGE Average amount of kWh dispensed per charge session to more accurate numbers

elec_sec_emp = []
elec_type1_sum = 0
elec_type1_sum_list = []
elec_induced_sum = 0
elec_induced_sum_list = []
elec_sec_totals = []

for i in range(len(inputs)):
    # Calculating supply chain employment:
    if inputs["Charger Power"][i] < 10.9:
        elec_cost = elec_rate[inputs["Geography"][i]][0]
    else:
        elec_cost = elec_rate[inputs["Geography"][i]][1]

        
    elec_in_region = elec_cost * empdef
    row_num = tier_mult[tier_mult['Code'] == "2211A0"].index[0]

    type1_elec_in_region = ((elec_in_region * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type2_elec_in_region = ((elec_in_region * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    
    avg_kwh = 0
    if inputs["[Average] amount of kWh dispensed per charge session"][i] in kWh_dict.keys():
        avg_kwh = kWh_dict[inputs["[Average] amount of kWh dispensed per charge session"][i]]
    else:
        avg_kwh = inputs["[Average] amount of kWh dispensed per charge session"][i]

    type1_total_elec_in_region = type1_elec_in_region * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    
    if str(inputs["yes/no to include induced effects"][i]).strip().lower() == "yes":
        induced_total_elec_region = ((type2_elec_in_region - type1_elec_in_region) * inputs["[Average] sessions per month"][i] * avg_kwh * 12)
    else:
        induced_total_elec_region = 0

    elec_type1_sum = 0
    elec_induced_sum = 0
    
    for j in range((inputs["Number of years for analysis"][i])):
        elec_type1_sum += type1_total_elec_in_region * (j+1)
        elec_induced_sum += induced_total_elec_region * (j+1)

    # if ((inputs["Number of years for analysis"][i]) < 10):
    #     elec_type1_sum += (type1_total_elec_in_region * (11 - (inputs["Number of years for analysis"][i])))
    #     elec_induced_sum += (induced_total_elec_region * (11 - (inputs["Number of years for analysis"][i])))
    
    elec_sector_total = elec_type1_sum + elec_induced_sum
    elec_type1_sum_list.append(elec_type1_sum)
    elec_induced_sum_list.append(elec_induced_sum)
    elec_sec_emp.append(elec_sector_total)

for i in range(len(elec_type1_sum_list)):
    curr_row = []
    curr_row.append(elec_type1_sum_list[i])
    curr_row.append(elec_induced_sum_list[i])
    curr_row.append(elec_sec_emp[i])
    elec_sec_totals.append(curr_row)

print(elec_type1_sum_list)
print(elec_induced_sum_list)
print(elec_sec_emp)
print(elec_sec_totals)



# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
elec_sect_df = pd.DataFrame(elec_sec_totals, index = row_headings_ops, columns = column_total_headings_ops)
elec_sect_df

# %% [markdown]
# #### Retail Sector Employment
# Includes Induced Purchases and takes total values for supply chain and induced employment

# %% [markdown]
# ##### Performing calculations for Retail Sector Employment

# %%
# Creating a dictionary that maps inputs["Retail dollars per session"][i] to more precise values:
retail_dict = {0:0,
               0.40: 0.397381954,
               0.70: 0.699396599,
               0.19: 0.189036}

retail_dict[0.4]

# %%
retail_sect_emp = []
retail_type1_sum = 0
retail_type1_sum_list = []
retail_induced_sum = 0
retail_induced_sum_list = []
retail_sect_totals = []


for i in range(len(inputs)):
    avg_kwh = 0
    if inputs["[Average] amount of kWh dispensed per charge session"][i] in kWh_dict.keys():
        avg_kwh = kWh_dict[inputs["[Average] amount of kWh dispensed per charge session"][i]]
    else:
        avg_kwh = inputs["[Average] amount of kWh dispensed per charge session"][i]

    retail_rev = (inputs["Retail dollars per session"][i]) / (inputs["[Average] sessions per month"][i] * avg_kwh)

    retail_in_region = retail_rev * empdef
    row_num = tier_mult[tier_mult['Code'] == 445000].index[0]
    type1_retail_in_region = ((retail_in_region * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type2_retail_in_region = ((retail_in_region * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])

    type1_total_retail_in_region = type1_retail_in_region * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    
    if str(inputs["yes/no to include induced effects"][i]).strip().lower() == "yes":
        induced_total_retail_region = ((type2_retail_in_region - type1_retail_in_region) * inputs["[Average] sessions per month"][i] * avg_kwh * 12)
    else:
        induced_total_retail_region = 0

    retail_type1_sum = 0
    retail_induced_sum = 0
    for j in range((inputs["Number of years for analysis"][i])):
        retail_type1_sum += type1_total_retail_in_region * (j+1)
        retail_induced_sum += induced_total_retail_region * (j+1)

    retail_sector_total = retail_type1_sum + retail_induced_sum
    retail_type1_sum_list.append(retail_type1_sum)
    retail_induced_sum_list.append(retail_induced_sum)
    retail_sect_emp.append(retail_sector_total)

for i in range(len(retail_type1_sum_list)):
    curr_row = []
    curr_row.append(retail_type1_sum_list[i])
    curr_row.append(retail_induced_sum_list[i])
    curr_row.append(retail_sect_emp[i])
    retail_sect_totals.append(curr_row)

print(retail_type1_sum_list)
print(retail_induced_sum_list)
print(retail_sect_emp)
print(retail_sect_totals)


# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
retail_sect_df = pd.DataFrame(retail_sect_totals, index = row_headings_ops, columns = column_total_headings_ops)
retail_sect_df

# %% [markdown]
# #### Advertising Sector Employment
# Includes Advertisements and takes total values for supply chain and induced employment

# %% [markdown]
# ##### Performing calculations for Advertising Sector Employment

# %%
# Creating a dictionary that maps inputs["Advertising"][i] to more precise values:
ad_dict = {0:0, 
           1.55: 1.554214664,
            0.49: 0.488012926,
            0.37: 0.366009695,
            0.18: 0.183004847}

ad_dict[0.37]


# %%
ad_emp = []
ad_type1_sum = 0
ad_type1_sum_list = []
ad_induced_sum = 0
ad_induced_sum_list = []
ad_totals = []

for i in range(len(inputs)):
    avg_kwh = 0
    if inputs["[Average] amount of kWh dispensed per charge session"][i] in kWh_dict.keys():
        avg_kwh = kWh_dict[inputs["[Average] amount of kWh dispensed per charge session"][i]]
    else:
        avg_kwh = inputs["[Average] amount of kWh dispensed per charge session"][i]

    ad_rev = (inputs["Advertising"][i]) / (inputs["[Average] sessions per month"][i] * avg_kwh)

    ad_in_region = ad_rev * empdef
    row_num = tier_mult[tier_mult['Code'] == 541800].index[0]
    type1_ad_in_region = ((ad_in_region * (type1_mult[inputs["Geography"][i]][row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type2_ad_in_region = ((ad_in_region * (type2_mult[inputs["Geography"][i]][row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])

    type1_total_ad_in_region = type1_ad_in_region * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    
    if str(inputs["yes/no to include induced effects"][i]).strip().lower() == "yes":
        induced_total_ad_region = ((type2_ad_in_region - type1_ad_in_region) * inputs["[Average] sessions per month"][i] * avg_kwh * 12)
    else:
        induced_total_ad_region = 0
    
    ad_type1_sum = 0
    ad_induced_sum = 0
    for j in range((inputs["Number of years for analysis"][i])):
        ad_type1_sum += type1_total_ad_in_region * (j+1)
        ad_induced_sum += induced_total_ad_region * (j+1)

    ad_sector_total = ad_type1_sum + ad_induced_sum
    ad_type1_sum_list.append(ad_type1_sum)
    ad_induced_sum_list.append(ad_induced_sum)
    ad_emp.append(ad_sector_total)

for i in range(len(ad_type1_sum_list)):
    curr_row = []
    curr_row.append(ad_type1_sum_list[i])
    curr_row.append(ad_induced_sum_list[i])
    curr_row.append(ad_emp[i])
    ad_totals.append(curr_row)

print(ad_type1_sum_list)
print(ad_induced_sum_list)
print(ad_emp)
print(ad_totals)


# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
ad_sect_df = pd.DataFrame(ad_totals, index = row_headings_ops, columns = column_total_headings_ops)
ad_sect_df

# %% [markdown]
# #### Data and Networking Sector Employment
# Includes Data Fees & Networking Fees and takes total values for supply chain and induced employment

# %% [markdown]
# ##### Performing calculations for Data and Networking Sector Employment

# %%
data_sec_emp = []
data_type1_sum = 0
data_type1_sum_list = []
data_induced_sum = 0
data_induced_sum_list = []
data_totals = []

for i in range(len(inputs)):
    # Calculating supply chain employment:
    if inputs["Charger Power"][i] <= 10.9 and inputs["Number of chargers per station"][i] == 1:
        data_cost = 0
    elif inputs["Charger Power"][i] < 10.9 and inputs["Number of chargers per station"][i] == 3:
        data_cost = 0.07575758 
    elif inputs["Charger Power"][i] <= 10.9 and inputs["[Average] sessions per month"][i] == 63:
        data_cost = 0.03330460 
    elif inputs["Charger Power"][i] <= 10.9 and inputs["[Average] sessions per month"][i] == 46:
        data_cost = 0.04561282 
    elif inputs["Charger Power"][i] == 50 and inputs["[Average] sessions per month"][i] == 76:
        data_cost = 0.03137226  
    elif inputs["Charger Power"][i] == 50 and inputs["[Average] sessions per month"][i] == 114:
        data_cost = 0.02091484   
    elif inputs["Charger Power"][i] == 150 :
        data_cost = 0.01568613 
    elif inputs["Charger Power"][i] == 350 :
        data_cost = 0.00784306  
    
    datanetwork_in_region = data_cost * empdef * 0.5
    
    data_row_num = tier_mult[tier_mult['Code'] == 518200].index[0]
    network_row_num = tier_mult[tier_mult['Code'] == 517210].index[0]

    type1_data_in_region = ((datanetwork_in_region * (type1_mult[inputs["Geography"][i]][data_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type1_network_in_region = ((datanetwork_in_region * (type1_mult[inputs["Geography"][i]][network_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])

    type2_data_in_region = ((datanetwork_in_region * (type2_mult[inputs["Geography"][i]][data_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type2_network_in_region = ((datanetwork_in_region * (type2_mult[inputs["Geography"][i]][network_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])


    avg_kwh = 0
    if inputs["[Average] amount of kWh dispensed per charge session"][i] in kWh_dict.keys():
        avg_kwh = kWh_dict[inputs["[Average] amount of kWh dispensed per charge session"][i]]
    else:
        avg_kwh = inputs["[Average] amount of kWh dispensed per charge session"][i]

    type1_total_datanetwork_in_region = (type1_data_in_region + type1_network_in_region) * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    
    if str(inputs["yes/no to include induced effects"][i]).strip().lower() == "yes":
        induced_total_datanetwork_region = ((type2_data_in_region + type2_network_in_region) - (type1_data_in_region + type1_network_in_region)) * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    else:
        induced_total_datanetwork_region = 0
        
    
    data_type1_sum = 0
    data_induced_sum = 0
    for j in range((inputs["Number of years for analysis"][i])):
        data_type1_sum += type1_total_datanetwork_in_region * (j+1)
        data_induced_sum += induced_total_datanetwork_region * (j+1)
    
    data_sector_total = data_type1_sum + data_induced_sum
    data_sec_emp.append(data_sector_total)
    data_type1_sum_list.append(data_type1_sum)
    data_induced_sum_list.append(data_induced_sum)

for i in range(len(data_type1_sum_list)):
    curr_row = []
    curr_row.append(data_type1_sum_list[i])
    curr_row.append(data_induced_sum_list[i])
    curr_row.append(data_sec_emp[i])
    data_totals.append(curr_row)

print(data_type1_sum_list)
print(data_induced_sum_list)
print(data_sec_emp)
print(data_totals)


# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
data_sect_df = pd.DataFrame(data_totals, index = row_headings_ops, columns = column_total_headings_ops)
data_sect_df

# %% [markdown]
# #### Warranty, Maintenance, Administrative Costs, & Access Fees Sector Employment														 
# Includes Data Fees & Networking Fees and takes total values for supply chain and induced employment

# %% [markdown]
# ##### Performing calculations for Warranty, Maintenance, Administrative Costs, & Access Fees Sector Employment

# %%
warr_sec_emp = []
warr_type1_sum = 0
warr_type1_sum_list = []
warr_induced_sum = 0
warr_induced_sum_list = []
warr_totals = []

for i in range(len(inputs)):
    # Calculating supply chain employment:
    if inputs["Charger Power"][i] <= 10.9 and inputs["Number of chargers per station"][i] == 1:
        admin = 0
    elif inputs["Charger Power"][i] < 10.9 and inputs["Number of chargers per station"][i] == 3:
        admin = 0.025252525  
    elif inputs["Charger Power"][i] <= 10.9 and inputs["[Average] sessions per month"][i] == 63:
        admin = 0.011101533  
    elif inputs["Charger Power"][i] <= 10.9 and inputs["[Average] sessions per month"][i] == 46:
        admin = 0.015204274  
    elif inputs["Charger Power"][i] == 50 and inputs["[Average] sessions per month"][i] == 76:
        admin = 0.005228710   
    elif inputs["Charger Power"][i] == 50 and inputs["[Average] sessions per month"][i] == 114:
        admin = 0.003485807    
    elif inputs["Charger Power"][i] == 150 :
        admin = 0.002614355  
    elif inputs["Charger Power"][i] == 350 :
        admin = 0.001307177 
    
    if inputs["Charger Power"][i] == 6.6 and inputs["Number of chargers per station"][i] == 1:
        maint =  0.025252525 
        warr =  0.050505051 
    elif inputs["Charger Power"][i] == 6.6 and inputs["Number of chargers per station"][i] == 3: 
        maint =  0.012626263 
        warr =  0.025252525 
    elif inputs["Charger Power"][i] == 10.9 and inputs["[Average] sessions per month"][i] == 30: 
        maint = 0.011656610  
        warr =  0.023313220
    elif inputs["Charger Power"][i] == 10.9 and inputs["[Average] sessions per month"][i] == 63: 
        maint = 0.005550767  
        warr = 0.011101533 
    elif inputs["Charger Power"][i] == 10.9 and inputs["[Average] sessions per month"][i] == 46: 
        maint = 0.007602137   
        warr = 0.015204274 
    elif inputs["Charger Power"][i] == 50 and inputs["[Average] sessions per month"][i] == 76: 
        maint = 0.002614355  
        warr = 0.005228710  
    elif inputs["Charger Power"][i] == 50 and inputs["[Average] sessions per month"][i] == 114: 
        maint = 0.001742903 
        warr = 0.003485807  
    elif inputs["Charger Power"][i] == 150 :
        maint = 0.001307177  
        warr = 0.002614355  
    elif inputs["Charger Power"][i] == 350 :
        maint = 0.000653589 
        warr = 0.001307177  
     
    
    admin_in_region = admin * empdef
    maint_in_region = maint * empdef
    warr_in_region = warr * empdef

    
    admin_row_num = tier_mult[tier_mult['Code'] == 561100].index[0]
    maint_row_num = tier_mult[tier_mult['Code'] == 811300].index[0]
    warr_row_num = tier_mult[tier_mult['Code'] == "5241XX"].index[0]

    type1_admin_in_region = ((admin_in_region * (type1_mult[inputs["Geography"][i]][admin_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type1_maint_in_region = ((maint_in_region * (type1_mult[inputs["Geography"][i]][maint_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type1_warr_in_region = ((warr_in_region * (type1_mult[inputs["Geography"][i]][warr_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])

    type2_admin_in_region = ((admin_in_region * (type2_mult[inputs["Geography"][i]][admin_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type2_maint_in_region = ((maint_in_region * (type2_mult[inputs["Geography"][i]][maint_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])
    type2_warr_in_region = ((warr_in_region * (type2_mult[inputs["Geography"][i]][warr_row_num])) / 1000000) * (inputs["Number of stations"][i] / inputs["Number of years for analysis"][i])

    avg_kwh = 0
    if inputs["[Average] amount of kWh dispensed per charge session"][i] in kWh_dict.keys():
        avg_kwh = kWh_dict[inputs["[Average] amount of kWh dispensed per charge session"][i]]
    else:
        avg_kwh = inputs["[Average] amount of kWh dispensed per charge session"][i]

    type1_total_warr_in_region = (type1_admin_in_region + type1_maint_in_region + type1_warr_in_region) * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    
    if str(inputs["yes/no to include induced effects"][i]).strip().lower() == "yes":
        induced_total_warr_region = ((type2_admin_in_region + type2_maint_in_region + type2_warr_in_region) - (type1_admin_in_region + type1_maint_in_region + type1_warr_in_region)) * inputs["[Average] sessions per month"][i] * avg_kwh * 12
    else:
        induced_total_warr_region = 0
    
    
    warr_type1_sum = 0
    warr_induced_sum = 0
    for j in range((inputs["Number of years for analysis"][i])):
        warr_type1_sum += type1_total_warr_in_region * (j+1)
        warr_induced_sum += induced_total_warr_region * (j+1)
    
    warr_sector_total = warr_type1_sum + warr_induced_sum
    warr_sec_emp.append(warr_sector_total)
    warr_type1_sum_list.append(warr_type1_sum)
    warr_induced_sum_list.append(warr_induced_sum)

for i in range(len(warr_type1_sum_list)):
    curr_row = []
    curr_row.append(warr_type1_sum_list[i])
    curr_row.append(warr_induced_sum_list[i])
    curr_row.append(warr_sec_emp[i])
    warr_totals.append(curr_row)

print(warr_type1_sum_list)
print(warr_induced_sum_list)
print(warr_sec_emp)
print(warr_totals)



# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
warr_sect_df = pd.DataFrame(warr_totals, index = row_headings_ops, columns = column_total_headings_ops)
warr_sect_df

# %% [markdown]
# #### Station Operations Employment Table

# %% [markdown]
# ##### Performing calculations for Stations Operations Employment Table

# %%
# Calculating totals for all EVSE components:

stat_ops_type1 = []
stat_ops_type2 = []
stat_ops_induced = []
station_ops_totals = []

for i in range(len(inputs)):
    type1 = (elec_type1_sum_list[i] + retail_type1_sum_list[i] + ad_type1_sum_list[i] + data_type1_sum_list[i] + warr_type1_sum_list[i])
    induced = (elec_induced_sum_list[i] + retail_induced_sum_list[i] + ad_induced_sum_list[i] + data_induced_sum_list[i] + warr_induced_sum_list[i]) 
    type2 = (elec_sec_emp[i] + retail_sect_emp[i] + ad_emp[i] + data_sec_emp[i] + warr_sec_emp[i])
    stat_ops_type1.append(type1)
    stat_ops_induced.append(induced)
    stat_ops_type2.append(type2)

for i in range(len(stat_ops_type1)):
    curr_row = []
    curr_row.append(stat_ops_type1[i])
    curr_row.append(stat_ops_induced[i])
    curr_row.append(stat_ops_type2[i])
    station_ops_totals.append(curr_row)

station_ops_totals

# %% [markdown]
# ##### Creating Dataframe with total output values

# %%
statops_df = pd.DataFrame(station_ops_totals, index = row_headings_ops, columns = column_total_headings_ops)
statops_df

# %% [markdown]
# ## Writing Output to Final Sheet

# %% [markdown]
# Writing output and resulting calculations to input excel sheet

# %% [markdown]
# ##### Station Development Employment:

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    statdev_df.to_excel(writer, sheet_name = "Station Development")

# Validation - Check to see if output matches expected values

# %% [markdown]
# ##### Station Operations Employment:

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    statops_df.to_excel(writer, sheet_name = "Station Operations")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Civil Construction Employment:

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    civil_const_df.to_excel(writer, sheet_name = "Civil Construction")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Electrical Construction Employment

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    elec_const_df.to_excel(writer, sheet_name = "Electrical Construction")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Electricity Sector Employment		

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    elec_sect_df.to_excel(writer, sheet_name = "Electricity Sector")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Retail Sector Employment		

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    retail_sect_df.to_excel(writer, sheet_name = "Retail Sector")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Advertising Sector Employment	

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    ad_sect_df.to_excel(writer, sheet_name = "Advertising Sector")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Data and Networking Sector Employment	

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    data_sect_df.to_excel(writer, sheet_name = "Data and Networking Sector")

# Validation - Check to see if output matches expected values

# %% [markdown]
# #### Warranty, Maintenance, Administrative Costs, & Access Fees Sector Employment														

# %%
# Final Part: Adding DataFrame as a new excel sheet in User-Input Excel sheet:
with pd.ExcelWriter(input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  
    warr_sect_df.to_excel(writer, sheet_name = "Warranty, Maintenance, Administrative Costs, & Access Fees Sector Employment Sector")

# Validation - Check to see if output matches expected values


