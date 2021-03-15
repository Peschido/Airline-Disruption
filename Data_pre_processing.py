from openpyxl import load_workbook 
import numpy as np 

dataset = "Datasets/A1_6088570/A01_6088570/"

#####   AIRCRAFT FILE
wb = load_workbook(dataset+"aircraft.xlsx", data_only=True) 
ws = wb.active 

aircraft_raw = np.array([[i.value for i in j] for j in ws ['A1':'A85']]) 
aircraft = np.array([["                                      " \
           for i in range(len(aircraft_raw[1][0].split(" ")) -1)] for j in range(len(aircraft_raw))])
for i in range(len(aircraft_raw)): 
    for j in range(len(aircraft_raw[i][0].split(" "))-1):
        aircraft[i][j] = aircraft_raw[i][0].split(" ")[j]
    
#####   AIRPORTS FILE
wb = load_workbook(dataset+"airports.xlsx", data_only=True) 
ws = wb.active 

airports_raw = np.array([[i.value for i in j] for j in ws ['A1':'A35']]) 
airports = np.array([["                                      " \
           for i in range(len(airports_raw[23][0].split(" ")) -1)] for j in range(len(airports_raw))])
for i in range(len(airports_raw)): 
    for j in range(len(airports_raw[i][0].split(" "))-1):
        airports[i][j] = airports_raw[i][0].split(" ")[j]
        
#####   CONFIGURATION FILE
wb = load_workbook(dataset+"config.xlsx", data_only=True) 
ws = wb.active 

config_raw = np.array([[i.value for i in j] for j in ws ['A1':'A7']]) 
config = np.array([["                                      " \
           for i in range(len(config_raw[4][0].split(" ")) -1)] for j in range(len(config_raw))])
for i in range(len(config_raw)): 
    for j in range(len(config_raw[i][0].split(" "))-1):
        config[i][j] = config_raw[i][0].split(" ")[j]
        
#####   DISTANCE FILE
wb = load_workbook(dataset+"dist.xlsx", data_only=True) 
ws = wb.active 

dist_raw = np.array([[i.value for i in j] for j in ws ['A1':'A1190']]) 
dist = np.array([["                                      " \
           for i in range(len(dist_raw[1][0].split(" ")) -1)] for j in range(len(dist_raw))])
for i in range(len(dist_raw)): 
    for j in range(len(dist_raw[i][0].split(" "))-1):
        dist[i][j] = dist_raw[i][0].split(" ")[j]
        
#####   CONFIGURATION FILE
wb = load_workbook(dataset+"config.xlsx", data_only=True) 
ws = wb.active 

config_raw = np.array([[i.value for i in j] for j in ws ['A1':'A7']]) 
config = np.array([["                                      " \
           for i in range(len(config_raw[4][0].split(" ")) -1)] for j in range(len(config_raw))])
for i in range(len(config_raw)): 
    for j in range(len(config_raw[i][0].split(" "))-1):
        config[i][j] = config_raw[i][0].split(" ")[j]
        
