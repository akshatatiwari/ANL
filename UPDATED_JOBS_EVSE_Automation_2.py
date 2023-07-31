# %% [markdown]
# ## JOBS EVSE Automation 2.0
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
# ### Excel Sheet Setup
# Reading in Excel files and setting up User-Input excel sheets that will contain employment outputs

# %% [markdown]
# 

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
# ### Station Development Calc
# Formulas for all EVSE Components that have to do with Station Development
# Calculating total values based on formulas and multipliers for each JOBS EVSE Component.

# %%
# NAME:         Cable Cooling (physical component)
# DESCRIPTION:  Air conditioning, refrigeration and warm air heating equipment manufacturing

tier_cable_cooling = []
type1_cable_cooling = []
type2_cable_cooling = []
# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Charger Power"][i] <= 50:
        cost2008 = 0
    elif inputs["Charger Power"][i] > 50:
        cost2008 = 500 
    
    row_num = tier_mult[tier_mult['Code'] == 333415].index[0]
    
    each_charger_tier = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    each_charger_type1 = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    each_charger_type2 = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

    total_val_tier = each_charger_tier * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    total_val_type1 = each_charger_type1 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    total_val_type2 = each_charger_type2 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]

    tier_cable_cooling.append(total_val_tier)
    type1_cable_cooling.append(total_val_type1)
    type2_cable_cooling.append(total_val_type2)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------


# NAME:         Charger (physical component)
# DESCRIPTION:  Other industrial machinery manufacturing

tier_charger = []
type1_charger = []
type2_charger = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Charger Power"][i] <= 6.6:
        if inputs["Number of chargers per station"][i] == 1:
            cost2008 = 200 * 0.84
        else:
            cost2008 = 530 * 0.84
    elif 6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2:
        if inputs["Number of chargers per station"][i] == 1:
            cost2008 = 900 * 0.84
        else:
            cost2008 = 4900 * 0.84
    elif inputs["Charger Power"][i] == 50:
        cost2008 = (27900) * 0.84
    elif inputs["Charger Power"][i] == 150:
        cost2008 = (87800) * 0.84
    elif inputs["Charger Power"][i] == 350:
        cost2008 = (140000) * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == "33329A"].index[0]
    
    tier_each_charger = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_each_charger = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_each_charger = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

    tier_total_val = tier_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    type1_total_val = type1_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    type2_total_val = type2_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]

    tier_charger.append(tier_total_val)
    type1_charger.append(type1_total_val)
    type2_charger.append(type2_total_val)



# %%
# NAME:         Conduit Cables (physical component)
# DESCRIPTION:  Communication and energy wire and cable manufacturing

tier_conduit_cables = []
type1_conduit_cables = []
type2_conduit_cables = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Charger Power"][i] <= 6.6:
        cost2008 = (525) * 0.84
    else:
        cost2008 = (1500) * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == 335920].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

    tier_conduit_cables.append(tier_total_val)
    type1_conduit_cables.append(type1_total_val)
    type2_conduit_cables.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------


# NAME:         Trenching and Boring Labor
# DESCRIPTION:  Nonresidential structures

tier_trenching = []
type1_trenching = []
type2_trenching = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 6000 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

    tier_trenching.append(tier_total_val)
    type1_trenching.append(type1_total_val)
    type2_trenching.append(type2_total_val)



# %%
# NAME:         On-site Electrical Storage
# DESCRIPTION:  Storage battery manufacturing

tier_electrical_storage = []
type1_electrical_storage = []
type2_electrical_storage = []
# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["yes/no to include onsite storage costs"][i] == "Yes" or inputs["yes/no to include onsite storage costs"][i] == "yes": 
        if inputs["Charger Power"][i] == 50 or inputs["Charger Power"][i] == 150 or inputs["Charger Power"][i] == 350:
            cost2008 = inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * inputs["Charger Power"][i] * 400 * 0.84
        else:
            cost2008 = 0
        
        row_num = tier_mult[tier_mult['Code'] == 335911].index[0]
        
        tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
        type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
        type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    else:
        tier_total_val = 0
        type1_total_val = 0
        type2_total_val = 0
    tier_electrical_storage.append(tier_total_val)
    type1_electrical_storage.append(type1_total_val)
    type2_electrical_storage.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Safety & Traffic Control
# DESCRIPTION:  Transportation structures and highways and streets

tier_safety_traffic = []
type1_safety_traffic = []
type2_safety_traffic = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Number of chargers per station"][i] == 3 and inputs["Charger Power"][i] <= 19.2:
        cost2008 = 1000 * 0.84
    else:
        cost2008 = 3000 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == "2332F0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of stations"][i]
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of stations"][i]
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of stations"][i]
    
    tier_safety_traffic.append(tier_total_val)
    type1_safety_traffic.append(type1_total_val)
    type2_safety_traffic.append(type2_total_val)



# %%
# NAME:         Load Center/Panels
# DESCRIPTION:  Switchgear and switchboard apparatus manufacturing

tier_load_center = []
type1_load_center = []
type2_load_center = []
# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Charger Power"][i] <= 6.6:
        cost2008 = 5 * 0.84
    elif 6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2:
        cost2008 = 10 * 0.84
    else:
        cost2008 = 40 * 0.84  
    
    row_num = tier_mult[tier_mult['Code'] == 335313].index[0]
    
    tier_each_charger = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_each_charger = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_each_charger = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

    tier_total_val = tier_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    type1_total_val = type1_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    type2_total_val = type2_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]

    tier_load_center.append(tier_total_val)
    type1_load_center.append(type1_total_val)
    type2_load_center.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Transformers
# DESCRIPTION:  Power, distribution, and specialty transformer manufacturing

tier_transformers = []
type1_transformers = []
type2_transformers = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["yes/no to include transformer costs"][i] == "Yes" or inputs["yes/no to include transformer costs"][i] == "yes": 
        if inputs["Charger Power"][i] == 50 or inputs["Charger Power"][i] == 150 or inputs["Charger Power"][i] == 350:
            kva = inputs["Number of chargers per station"][i] * inputs["Charger Power"][i] * 0.9
            cost2008 = (0.0066 * (kva*kva)) + (48.43 * kva) + (9788.5) * 0.84
        else:
            cost2008 = 0
        row_num = tier_mult[tier_mult['Code'] == 335311].index[0]
        
        tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of stations"][i] 
        type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of stations"][i] 
        type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of stations"][i] 
    else:
        tier_total_val = 0
        type1_total_val = 0
        type2_total_val = 0
    tier_transformers.append(tier_total_val)
    type1_transformers.append(type1_total_val)
    type2_transformers.append(type2_total_val)


# %%
# NAME:         Meters
# DESCRIPTION:  Electrical Meters

# Code not found in the JOBS EVSE Formulas sheet
# Add check to see Yes/No to include meters in input sheet

tier_meters = []
type1_meters = []
type2_meters = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["yes/no to include meters"][i] == "Yes" or inputs["yes/no to include meters"][i] == "yes":
        if inputs["Number of chargers per station"][i] == 1:
            cost2008 = 0
        else:
            if inputs["Charger Power"][i] < 150:
                cost2008 = 2000 * 2 * 0.84 
            else:
                cost2008 = 2000 * 0.84
        row_num = tier_mult[tier_mult['Code'] == 334515].index[0]
        
        tier_each_charger = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
        type1_each_charger = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
        type2_each_charger = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

        tier_total_val = tier_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
        type1_total_val = type1_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
        type2_total_val = type2_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]

    else:
        tier_total_val = 0
        type1_total_val = 0
        type2_total_val = 0
    
    tier_meters.append(tier_total_val)
    type1_meters.append(type1_total_val)
    type2_meters.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Misc. (mounting hardware, etc.)
# DESCRIPTION:  Other concrete product manufacturing

tier_misc = []
type1_misc = []
type2_misc = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Number of chargers per station"][i] == 1:
        cost2008 = 100 * 0.84
    elif (6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2) or (inputs["Charger Power"][i] == 50):
        cost2008 = 1000 * 0.84
    else:
        cost2008 = 2000 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == 327390].index[0]
    
    tier_each_charger = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_each_charger = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_each_charger = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008

    tier_total_val = tier_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    type1_total_val = type1_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]
    type2_total_val = type2_each_charger * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i]

    tier_misc.append(tier_total_val)
    type1_misc.append(type1_total_val)
    type2_misc.append(type2_total_val)

# %% [markdown]
# #### EVSE Components that include total cost of physical components as formula calculation.
# (Still part of Station Development)

# %%
# EVSE COMPONENTS where formula = total cost of physical components 

# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# Function that calculates total cost of physical components:
def physical_components_total(i):
    total_cost = 0
    # Cable cooling:
    if inputs["Charger Power"][i] <= 50:
        total_cost += 0
    elif inputs["Charger Power"][i] > 50:
        total_cost += 500

    # Charger:
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

    # Conduit Cables:
    if inputs["Charger Power"][i] <= 6.6:
        total_cost += (525)
    else:
        total_cost += (1500) 

    # On-site Electrical Storage
    if inputs["Charger Power"][i] == 50 or inputs["Charger Power"][i] == 150 or inputs["Charger Power"][i] == 350:
        total_cost += inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * inputs["Charger Power"][i] * 400 
    else:
        total_cost += 0

    # Safety and Traffic Control: <<
    if inputs["Number of chargers per station"][i] == 3 and inputs["Charger Power"][i] <= 19.2:
        total_cost += 1000 
    else:
        total_cost += 3000 
    
    # Load Center 
    if inputs["Charger Power"][i] <= 6.6:
        total_cost += 5 
    elif 6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2:
        total_cost += 10 
    else:
        total_cost += 40

    # Transformers
    if inputs["Charger Power"][i] == 50 or inputs["Charger Power"][i] == 150 or inputs["Charger Power"][i] == 350:
        kva = inputs["Number of chargers per station"][i] * inputs["Charger Power"][i] * 0.9
        total_cost += (0.0066 * (kva*kva)) + (48.43 * kva) + (9788.5) 
    else:
        total_cost += 0

    # Meters
    if inputs["Number of chargers per station"][i] == 1:
        total_cost += 0
    else:
        if inputs["Charger Power"][i] < 150:
            total_cost += 2000 * 2 
        else:
            total_cost += 2000 

    # Misc
    if inputs["Number of chargers per station"][i] == 1:
        total_cost += 100 
    elif (6.6 < inputs["Charger Power"][i] and inputs["Charger Power"][i] <= 19.2) or (inputs["Charger Power"][i] == 50):
        total_cost += 1000
    else:
        total_cost += 2000
    
    return total_cost

# %%
# NAME:         Equipment wholesale margin
# DESCRIPTION:  Wholesale trade

tier_wholesale = []
type1_wholesale = []
type2_wholesale = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.84
    if inputs["Geography"][i] == "USA-National":
        cost2008 = 0	
    
    row_num = tier_mult[tier_mult['Code'] == 420000].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_wholesale.append(tier_total_val)
    type1_wholesale.append(type1_total_val)
    type2_wholesale.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Equipment transportation
# DESCRIPTION:  Truck transportation

tier_etransp = []
type1_etransp = []
type2_etransp = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.84
    if inputs["Geography"][i] == "USA-National":
        cost2008 = cost2008 * 0.03
    else:
        cost2008 = cost2008 * 0.5 * 0.03
    
    row_num = tier_mult[tier_mult['Code'] == 484000].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_etransp.append(tier_total_val)
    type1_etransp.append(type1_total_val)
    type2_etransp.append(type2_total_val)


# %%
# NAME:         Equipment Installation
# DESCRIPTION:  Nonresidential structures

tier_einstall = []
type1_einstall = []
type2_einstall = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.3 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_einstall.append(tier_total_val)
    type1_einstall.append(type1_total_val)
    type2_einstall.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Site Prep & Construction
# DESCRIPTION:  Nonresidential structures

tier_siteprep = []
type1_siteprep = []
type2_siteprep = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.095 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_siteprep.append(tier_total_val)
    type1_siteprep.append(type1_total_val)
    type2_siteprep.append(type2_total_val)

# %%
# NAME:         Electrical Infrastrastructure & Make Ready
# DESCRIPTION:  Power, distribution, and specialty transformer manufacturing

tier_einfra = []
type1_einfra = []
type2_einfra = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.3 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == 335311].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_einfra.append(tier_total_val)
    type1_einfra.append(type1_total_val)
    type2_einfra.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Engineering & Design
# DESCRIPTION:  Architectural, engineering, and related services

tier_eng = []
type1_eng = []
type2_eng = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.195 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == 541300].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_eng.append(tier_total_val)
    type1_eng.append(type1_total_val)
    type2_eng.append(type2_total_val)


# %%
# NAME:         Permitting
# DESCRIPTION:  Architectural, engineering, and related services

tier_permit = []
type1_permit = []
type2_permit = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.03 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == 541300].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_permit.append(tier_total_val)
    type1_permit.append(type1_total_val)
    type2_permit.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Contingencies - Install
# DESCRIPTION:  Nonresidential structures

tier_cinstal = []
type1_cinstal = []
type2_cinstal = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 =  (physical_components_total(i) * 0.025 * 0.84 * 0.5) + (physical_components_total(i) * 0.025 * 0.84 * 0.5)
    
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_cinstal.append(tier_total_val)
    type1_cinstal.append(type1_total_val)
    type2_cinstal.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Contingencies - Site Prep & Construction
# DESCRIPTION:  Nonresidential structures

tier_csiteprep = []
type1_csiteprep = []
type2_csiteprep = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.025 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == "2332E0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_csiteprep.append(tier_total_val)
    type1_csiteprep.append(type1_total_val)
    type2_csiteprep.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Contingencies - Electrical
# DESCRIPTION:  Power, distribution, and specialty transformer manufacturing

tier_celec = []
type1_celec = []
type2_celec = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = physical_components_total(i) * 0.025 * 0.84
    
    row_num = tier_mult[tier_mult['Code'] == 335311].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008
 
    tier_celec.append(tier_total_val)
    type1_celec.append(type1_total_val)
    type2_celec.append(type2_total_val)


# %% [markdown]
# ### Station Operation Calc
# Formulas for all EVSE Components that have to do with Station Development

# %%
# NAME:         Electricity Cost to Station
# DESCRIPTION:  Electric power generation, transmission, and distribution
tier_eleccost = []
type1_eleccost = []
type2_eleccost = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = elec_rate[inputs["Geography"][i]][2] * inputs["[Average] sessions per month"][i] * inputs["[Average] amount of kWh dispensed per charge session"][i] * 12 * 0.84
    row_num = tier_mult[tier_mult['Code'] == "2211A0"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_eleccost.append(tier_total_val)
    type1_eleccost.append(type1_total_val)
    type2_eleccost.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Administrative Expense
# DESCRIPTION:  Office administrative services
tier_admincost = []
type1_admincost = []
type2_admincost = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 10 * 0.84

    row_num = tier_mult[tier_mult['Code'] == 561100].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_admincost.append(tier_total_val)
    type1_admincost.append(type1_total_val)
    type2_admincost.append(type2_total_val)


# %%
# NAME:         Maintenance Expense
# DESCRIPTION:  Commercial and industrial machinery and equipment repair and maintenance
tier_maint = []
type1_maint = []
type2_maint = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 5 * 0.84

    row_num = tier_mult[tier_mult['Code'] == 811300].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_maint.append(tier_total_val)
    type1_maint.append(type1_total_val)
    type2_maint.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Warranty Expense
# DESCRIPTION:  Insurance carriers, except direct life insurance
tier_warr = []
type1_warr = []
type2_warr = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 10 * 0.84

    row_num = tier_mult[tier_mult['Code'] == "5241XX"].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * inputs["Number of chargers per station"][i] * inputs["Number of stations"][i] * 12 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_warr.append(tier_total_val)
    type1_warr.append(type1_total_val)
    type2_warr.append(type2_total_val)


# %%
# NAME:         Data Fees (assume 50% of data and networking)
# DESCRIPTION:  Data processing, hosting, and related services
tier_data = []
type1_data = []
type2_data = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Number of chargers per station"][i] == 1:
        cost2008 = 0
    elif inputs["Charger Power"][i] <= 19.2:
        cost2008 = 30 * 0.84 * 12 * 0.5
    else:
        cost2008 = 60 * 0.84 * 12 * 0.5

    row_num = tier_mult[tier_mult['Code'] == 518200].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_data.append(tier_total_val)
    type1_data.append(type1_total_val)
    type2_data.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Networking Fees (assumed 50% of data and networking)
# DESCRIPTION:  Wireless telecommunications carriers (except satellite)
tier_networking = []
type1_networking = []
type2_networking = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 0
    if inputs["Number of chargers per station"][i] == 1:
        cost2008 = 0
    elif inputs["Charger Power"][i] <= 19.2:
        cost2008 = 30 * 0.84 * 12 * 0.5
    else:
        cost2008 = 60 * 0.84 * 12 * 0.5
        
    row_num = tier_mult[tier_mult['Code'] == 517210].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_networking.append(tier_total_val)
    type1_networking.append(type1_total_val)
    type2_networking.append(type2_total_val)

# %%
# NAME:         Advertisements
# DESCRIPTION:  Advertising Agencies
tier_ads = []
type1_ads = []
type2_ads = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = 1400 * 12 * 0.84
    row_num = tier_mult[tier_mult['Code'] == 541800].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_ads.append(tier_total_val)
    type1_ads.append(type1_total_val)
    type2_ads.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Induced Purchases
# DESCRIPTION:  Retail Sales
tier_induced = []
type1_induced = []
type2_induced = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    if inputs["yes/no to include induced effects"][i] == "Yes" or inputs["yes/no to include induced effects"][i] == "yes":
        cost2008 = inputs["Retail dollars per session"][i] * inputs["[Average] sessions per month"][i] * 12 * 0.84
        row_num = tier_mult[tier_mult['Code'] == 445000].index[0]
        
        tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
        type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
        type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    else:
        tier_total_val = 0
        type1_total_val = 0
        type2_total_val = 0
    tier_induced.append(tier_total_val)
    type1_induced.append(type1_total_val)
    type2_induced.append(type2_total_val)


# ------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------

# NAME:         Access Fees
# DESCRIPTION:  Office administrative services
tier_access = []
type1_access = []
type2_access = []

# Finding total Tier Employment for each run:
for i in range(len(inputs)):
    # Find tier employment per charger:
    cost2008 = inputs["Access fees"][i] * 0.84
    row_num = tier_mult[tier_mult['Code'] == 561100].index[0]
    
    tier_total_val = ((tier_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type1_total_val = ((type1_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
    type2_total_val = ((type2_mult[inputs["Geography"][i]][row_num]) / 1000000) * cost2008 * (inputs["Number of years for analysis"][i] - 1)
 
    tier_access.append(tier_total_val)
    type1_access.append(type1_total_val)
    type2_access.append(type2_total_val)


# %% [markdown]
# ### Calculating Total Values

# %% [markdown]
# Tier, Type I, Type II, Direct, Indirect, Induced Employment

# %% [markdown]
# #### Full List of EVSE Component Totals:

# %% [markdown]
# Calculating total values and outputs for civil construction employment.
# INCLUDES:
# 
# Civil Construction Employment	
# 
# - Cable Cooling
# - Charger
# - Conduit and Cables
# - Trenching and Boring Labor
# - Safety & Traffic Control
# - Meters
# - Misc. (mounting hardware, etc.)
# 
# 
# Electrical Construction Employment	
# 
# - On-site Electrical Storage
# - Load Center/Panels
# - Transformers
# - Electrical Infrastrastructure
# - Electricity Cost to Station
# 
# 
# Retail Sector Employment		
# - Equipment wholesale margin
# - Equipment transportation
# - Equipment Installation
# - Site Prep & Construction
# - Engineering & Design
# - Permitting
# - Contingencies - Install
# - Contingencies - Site Prep
# - Contingencies - Electrical
# - Induced Purchases
# 
# Advertising Sector Employment	
# - Advertisements
# 
# Data and Networking Sector Employment	
# - Data Fees 
# - Networking Fees
# 
# Warranty, Maintenance, Administrative Costs, & Access Fees Sector Employment														
# - Administrative Expense
# - Maintenance Expense
# - Warranty Expense
# - Access Fees
# 

# %% [markdown]
# #### EVSE Component Totals:
# Includes all JOBS EVSE Components are their total values for:
# - Tier Employment
# - Type I Employment
# - Type II Employment
# - Direct employment
# - Indirect employment
# - Induced employment

# %% [markdown]
# Formulas Used:
# - Tier employment = the corresponding1 multiplier from the tier employment table / 1,000,000 * cost2008
# - Type 1 employment = the corresponding1 multiplier from the type 1 employment table / 1,000,000 *
# cost2008
# - Type II employment = the corresponding1 multiplier from the type 2 employment table / 1,000,000 *
# cost2008
# - Direct employment = Tier Employment
# - Indirect employment = Type I employment - Tier employment
# - Induced employment = Type II employment - Type I employment

# %%
# Creating Row and Column headings that will be used in each dataframe:
#   ROW HEADINGS:    Contains RUN #
#   COLUMN HEADINGS: Contains total values that will be calculated - Tier, Type I, Type II, ...

row_headings = []
for i in range(len(inputs)):
    row_headings.append("RUN " + str(i + 1))

column_total_headings = ["Tier employment", "Type 1 Employment", "Type II Employment", 
                        "Direct Employment", "Indirect Employment", "Induced Employment"]

print(row_headings)
print(column_total_headings)

# %%
# Calculating totals for all EVSE components:

tier_total = []
type1_total = []
type2_total = []
direct_emp_total = []
indirect_emp_total = []
induced_emp_total = []
all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    tier_total_eachrow = tier_cable_cooling[i] + tier_charger[i] + tier_conduit_cables[i] + tier_trenching[i] + tier_electrical_storage[i] + tier_safety_traffic[i] + tier_load_center[i] + tier_transformers[i] + tier_meters[i] + tier_misc[i] + tier_wholesale[i] + tier_etransp[i] + tier_einstall[i] + tier_siteprep[i] + tier_einfra[i] + tier_eng[i] + tier_permit[i] + tier_cinstal[i] + tier_csiteprep[i] + tier_celec[i] + tier_eleccost[i] + tier_admincost[i] + tier_maint[i] + tier_warr[i] + tier_data[i] + tier_networking[i] + tier_ads[i] + tier_induced[i] + tier_access[i]
    type1_total_eachrow = type1_cable_cooling[i] + type1_charger[i] + type1_conduit_cables[i] + type1_trenching[i] + type1_electrical_storage[i] + type1_safety_traffic[i] + type1_load_center[i] + type1_transformers[i] + type1_meters[i] + type1_misc[i] + type1_wholesale[i] + type1_etransp[i] + type1_einstall[i] + type1_siteprep[i] + type1_einfra[i] + type1_eng[i] + type1_permit[i] + type1_cinstal[i] + type1_csiteprep[i] + type1_celec[i] + type1_eleccost[i] + type1_admincost[i] + type1_maint[i] + type1_warr[i] + type1_data[i] + type1_networking[i] + type1_ads[i] + type1_induced[i] + type1_access[i]
    type2_total_eachrow = type2_cable_cooling[i] + type2_charger[i] + type2_conduit_cables[i] + type2_trenching[i] + type2_electrical_storage[i] + type2_safety_traffic[i] + type2_load_center[i] + type2_transformers[i] + type2_meters[i] + type2_misc[i] + type2_wholesale[i] + type2_etransp[i] + type2_einstall[i] + type2_siteprep[i] + type2_einfra[i] + type2_eng[i] + type2_permit[i] + type2_cinstal[i] + type2_csiteprep[i] + type2_celec[i] + type2_eleccost[i] + type2_admincost[i] + type2_maint[i] + type2_warr[i] + type2_data[i]+ type2_networking[i] + type2_ads[i] + type2_induced[i] + type2_access[i]
    tier_total.append(tier_total_eachrow)
    type1_total.append(type1_total_eachrow)
    type2_total.append(type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(tier_total)):
    direct_emp_total.append(tier_total[i])
    indirect_emp_total.append(type1_total[i] - tier_total[i])
    induced_emp_total.append(type2_total[i] - type1_total[i])
    
for j in range(len(tier_total)):
    dummy_list = []
    dummy_list.append(tier_total[j])
    dummy_list.append(type1_total[j])
    dummy_list.append(type2_total[j])
    dummy_list.append(direct_emp_total[j])
    dummy_list.append(indirect_emp_total[j])
    dummy_list.append(induced_emp_total[j])
    all_totals.append(dummy_list)


print(all_totals)


# %%
all_totals_df = pd.DataFrame(all_totals, index = row_headings, columns = column_total_headings)
all_totals_df

# %% [markdown]
# #### Civil Construction Employment:
# 
# Includes:
# 
# - Cable Cooling
# - Charger
# - Conduit and Cables
# - Trenching and Boring Labor
# - Safety & Traffic Control
# - Meters
# - Misc. (mounting hardware, etc.)

# %%
# Calculating totals for all EVSE components:

civil_tier_total = []
civil_type1_total = []
civil_type2_total = []
civil_direct_emp_total = []
civil_indirect_emp_total = []
civil_induced_emp_total = []
civil_all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    civil_tier_total_eachrow = tier_cable_cooling[i] + tier_charger[i] + tier_conduit_cables[i] + tier_trenching[i] + tier_safety_traffic[i] + tier_meters[i] + tier_misc[i]
    civil_type1_total_eachrow = type1_cable_cooling[i] + type1_charger[i] + type1_conduit_cables[i] + type1_trenching[i] + type1_safety_traffic[i] + type1_meters[i] + type1_misc[i] 
    civil_type2_total_eachrow = type2_cable_cooling[i] + type2_charger[i] + type2_conduit_cables[i] + type2_trenching[i] + type2_safety_traffic[i] + type2_meters[i] + type2_misc[i] 
    civil_tier_total.append(civil_tier_total_eachrow)
    civil_type1_total.append(civil_type1_total_eachrow)
    civil_type2_total.append(civil_type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(civil_tier_total)):
    civil_direct_emp_total.append(civil_tier_total[i])
    civil_indirect_emp_total.append(civil_type1_total[i] - civil_tier_total[i])
    civil_induced_emp_total.append(civil_type2_total[i] - civil_type1_total[i])
    
for j in range(len(civil_tier_total)):
    civil_dummy_list = []
    civil_dummy_list.append(civil_tier_total[j])
    civil_dummy_list.append(civil_type1_total[j])
    civil_dummy_list.append(civil_type2_total[j])
    civil_dummy_list.append(civil_direct_emp_total[j])
    civil_dummy_list.append(civil_indirect_emp_total[j])
    civil_dummy_list.append(civil_induced_emp_total[j])
    civil_all_totals.append(civil_dummy_list)


print(civil_all_totals)




# %%
civil_all_totals_df = pd.DataFrame(civil_all_totals, index = row_headings, columns = column_total_headings)
civil_all_totals_df

# %% [markdown]
# #### Electrical Construction Employment	
# 
# - On-site Electrical Storage
# - Load Center/Panels
# - Transformers
# - Electrical Infrastrastructure
# - Electricity Cost to Station

# %%
# Calculating totals for all EVSE components:

elec_tier_total = []
elec_type1_total = []
elec_type2_total = []
elec_direct_emp_total = []
elec_indirect_emp_total = []
elec_induced_emp_total = []
elec_all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    elec_tier_total_eachrow = tier_electrical_storage[i] + tier_load_center[i] + tier_transformers[i] + tier_einfra[i] + tier_eleccost[i]
    elec_type1_total_eachrow = type1_electrical_storage[i] + type1_load_center[i] + type1_transformers[i] + type1_einfra[i] + type1_eleccost[i]
    elec_type2_total_eachrow = type2_electrical_storage[i] + type2_load_center[i] + type2_transformers[i] + type2_einfra[i] + type2_eleccost[i]
    elec_tier_total.append(elec_tier_total_eachrow)
    elec_type1_total.append(elec_type1_total_eachrow)
    elec_type2_total.append(elec_type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(elec_tier_total)):
    elec_direct_emp_total.append(elec_tier_total[i])
    elec_indirect_emp_total.append(elec_type1_total[i] - elec_tier_total[i])
    elec_induced_emp_total.append(elec_type2_total[i] - elec_type1_total[i])
    
for j in range(len(elec_tier_total)):
    elec_dummy_list = []
    elec_dummy_list.append(elec_tier_total[j])
    elec_dummy_list.append(elec_type1_total[j])
    elec_dummy_list.append(elec_type2_total[j])
    elec_dummy_list.append(elec_direct_emp_total[j])
    elec_dummy_list.append(elec_indirect_emp_total[j])
    elec_dummy_list.append(elec_induced_emp_total[j])
    elec_all_totals.append(elec_dummy_list)


print(elec_all_totals)


# %%
elec_all_totals_df = pd.DataFrame(elec_all_totals, index = row_headings, columns = column_total_headings)
elec_all_totals_df

# %% [markdown]
# #### Retail Sector Employment		
# - Equipment wholesale margin
# - Equipment transportation
# - Equipment Installation
# - Site Prep & Construction
# - Engineering & Design
# - Permitting
# - Contingencies - Install
# - Contingencies - Site Prep
# - Contingencies - Electrical
# - Induced Purchases

# %%
# Calculating totals for all EVSE components:

retail_tier_total = []
retail_type1_total = []
retail_type2_total = []
retail_direct_emp_total = []
retail_indirect_emp_total = []
retail_induced_emp_total = []
retail_all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    retail_tier_total_eachrow = tier_wholesale[i] + tier_etransp[i] + tier_einstall[i] + tier_siteprep[i] + tier_eng[i] + tier_permit[i] + tier_cinstal[i] + tier_csiteprep[i] + tier_celec[i] + tier_induced[i]
    retail_type1_total_eachrow = type1_wholesale[i] + type1_etransp[i] + type1_einstall[i] + type1_siteprep[i] + type1_eng[i] + type1_permit[i] + type1_cinstal[i] + type1_csiteprep[i] + type1_celec[i] + type1_induced[i]
    retail_type2_total_eachrow = type1_wholesale[i] + type2_etransp[i] + type2_einstall[i] + type2_siteprep[i] + type2_eng[i] + type2_permit[i] + type2_cinstal[i] + type2_csiteprep[i] + type2_celec[i] + type2_induced[i]
    retail_tier_total.append(retail_tier_total_eachrow)
    retail_type1_total.append(retail_type1_total_eachrow)
    retail_type2_total.append(retail_type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(retail_tier_total)):
    retail_direct_emp_total.append(retail_tier_total[i])
    retail_indirect_emp_total.append(retail_type1_total[i] - retail_tier_total[i])
    retail_induced_emp_total.append(retail_type2_total[i] - retail_type1_total[i])
    
for j in range(len(retail_tier_total)):
    retail_dummy_list = []
    retail_dummy_list.append(retail_tier_total[j])
    retail_dummy_list.append(retail_type1_total[j])
    retail_dummy_list.append(retail_type2_total[j])
    retail_dummy_list.append(retail_direct_emp_total[j])
    retail_dummy_list.append(retail_indirect_emp_total[j])
    retail_dummy_list.append(retail_induced_emp_total[j])
    retail_all_totals.append(retail_dummy_list)


print(retail_all_totals)


# %%
retail_all_totals_df = pd.DataFrame(retail_all_totals, index = row_headings, columns = column_total_headings)
retail_all_totals_df

# %% [markdown]
# #### Advertising Sector Employment	
# - Advertisements
# 

# %%
# Calculating totals for all EVSE components:

ad_tier_total = []
ad_type1_total = []
ad_type2_total = []
ad_direct_emp_total = []
ad_indirect_emp_total = []
ad_induced_emp_total = []
ad_all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    ad_tier_total_eachrow = tier_ads[i]
    ad_type1_total_eachrow = type1_ads[i]
    ad_type2_total_eachrow = type2_ads[i]
    ad_tier_total.append(ad_tier_total_eachrow)
    ad_type1_total.append(ad_type1_total_eachrow)
    ad_type2_total.append(ad_type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(ad_tier_total)):
    ad_direct_emp_total.append(ad_tier_total[i])
    ad_indirect_emp_total.append(ad_type1_total[i] - ad_tier_total[i])
    ad_induced_emp_total.append(ad_type2_total[i] - ad_type1_total[i])
    
for j in range(len(tier_total)):
    ad_dummy_list = []
    ad_dummy_list.append(ad_tier_total[j])
    ad_dummy_list.append(ad_type1_total[j])
    ad_dummy_list.append(ad_type2_total[j])
    ad_dummy_list.append(ad_direct_emp_total[j])
    ad_dummy_list.append(ad_indirect_emp_total[j])
    ad_dummy_list.append(ad_induced_emp_total[j])
    ad_all_totals.append(ad_dummy_list)


print(ad_all_totals)


# %%
ad_all_totals_df = pd.DataFrame(ad_all_totals, index = row_headings, columns = column_total_headings)
ad_all_totals_df

# %% [markdown]
# #### Data and Networking Sector Employment	
# - Data Fees 
# - Networking Fees

# %%
# Calculating totals for all EVSE components:

data_tier_total = []
data_type1_total = []
data_type2_total = []
data_direct_emp_total = []
data_indirect_emp_total = []
data_induced_emp_total = []
data_all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    data_tier_total_eachrow = tier_data[i] + tier_networking[i]
    data_type1_total_eachrow = type1_data[i] + type1_networking[i]
    data_type2_total_eachrow = type2_data[i]+ type2_networking[i]
    data_tier_total.append(data_tier_total_eachrow)
    data_type1_total.append(data_type1_total_eachrow)
    data_type2_total.append(data_type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(data_tier_total)):
    data_direct_emp_total.append(data_tier_total[i])
    data_indirect_emp_total.append(data_type1_total[i] - data_tier_total[i])
    data_induced_emp_total.append(data_type2_total[i] - data_type1_total[i])
    
for j in range(len(data_tier_total)):
    data_dummy_list = []
    data_dummy_list.append(data_tier_total[j])
    data_dummy_list.append(data_type1_total[j])
    data_dummy_list.append(data_type2_total[j])
    data_dummy_list.append(data_direct_emp_total[j])
    data_dummy_list.append(data_indirect_emp_total[j])
    data_dummy_list.append(data_induced_emp_total[j])
    data_all_totals.append(data_dummy_list)


print(data_all_totals)


# %%
data_all_totals_df = pd.DataFrame(data_all_totals, index = row_headings, columns = column_total_headings)
data_all_totals_df

# %% [markdown]
# #### Warranty, Maintenance, Administrative Costs, & Access Fees Sector Employment														
# - Administrative Expense
# - Maintenance Expense
# - Warranty Expense
# - Access Fees

# %%
# Calculating totals for all EVSE components:

warr_tier_total = []
warr_type1_total = []
warr_type2_total = []
warr_direct_emp_total = []
warr_indirect_emp_total = []
warr_induced_emp_total = []
warr_all_totals = []


for i in range(len(inputs)):
    # Tier totals:
    warr_tier_total_eachrow = tier_admincost[i] + tier_maint[i] + tier_warr[i] + tier_access[i]
    warr_type1_total_eachrow = type1_admincost[i] + type1_maint[i] + type1_warr[i] + type1_access[i]
    warr_type2_total_eachrow = type2_admincost[i] + type2_maint[i] + type2_warr[i] + type2_access[i]
    warr_tier_total.append(warr_tier_total_eachrow)
    warr_type1_total.append(warr_type1_total_eachrow)
    warr_type2_total.append(warr_type2_total_eachrow)   


# Calculating direct, indirect, and induced employment totals:
for i in range(len(warr_tier_total)):
    warr_direct_emp_total.append(warr_tier_total[i])
    warr_indirect_emp_total.append(warr_type1_total[i] - warr_tier_total[i])
    warr_induced_emp_total.append(warr_type2_total[i] - warr_type1_total[i])
    
for j in range(len(warr_tier_total)):
    warr_dummy_list = []
    warr_dummy_list.append(warr_tier_total[j])
    warr_dummy_list.append(warr_type1_total[j])
    warr_dummy_list.append(warr_type2_total[j])
    warr_dummy_list.append(warr_direct_emp_total[j])
    warr_dummy_list.append(warr_indirect_emp_total[j])
    warr_dummy_list.append(warr_induced_emp_total[j])
    warr_all_totals.append(warr_dummy_list)


print(warr_all_totals)


# %%
warr_all_totals_df = pd.DataFrame(warr_all_totals, index = row_headings, columns = column_total_headings)
warr_all_totals_df

# %% [markdown]
# ### Writing Output to Final Sheet

# %% [markdown]
# Writing output and resulting calculations to input excel sheet

# %% [markdown]
# Have removed the code for this part for now and will add it in at the end. 

# %% [markdown]
# 


