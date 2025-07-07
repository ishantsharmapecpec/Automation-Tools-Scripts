# -*- coding: utf-8 -*-
"""
Created on Mon Jul 22 10:46:04 2024

@author: ISTM
"""

import pandas as pd
import xml.etree.ElementTree as ET
import os
import copy
#from xml.dom import minidom

# Constants
EXCEL_FILE = r"\\rammdhfile02\Filery Supplementary Folders\Buildings\2023\RUK2023N00283 - HS2 Curzon Station CAT III Checking\4 Delivery\Geo\11 Analyses\WP05\Stiffness matrix\Stiffness MatrixV27.xlsm"
NEW_DIRECTORY = r"\\rammdhfile02\Filery Supplementary Folders\Buildings\2023\RUK2023N00283 - HS2 Curzon Station CAT III Checking\4 Delivery\Geo\11 Analyses\WP05\Repute ANALYSIS - TO RUN SCRIPTS\SET B\Set B below GL10 25oct24_REC"

# Load Excel file
df = pd.read_excel(EXCEL_FILE, sheet_name='Repute Input', engine='openpyxl')

# Change to the new directory
os.chdir(NEW_DIRECTORY)

# Function Definitions
def duplicate_and_insert(parent, element_name, num_copies):
    elements = parent.findall(element_name)
    for element in elements:
        base_name = element.find(".//Item.name").text
        try:
            current_count = int(base_name.split()[-1])
        except ValueError:
            #print(f"Warning: The base name '{base_name}' does not end with an integer. Skipping this element.")
            continue
        duplicates = []
        for i in range(1, num_copies + 1):
            new_element = copy.deepcopy(element)
            item_name = new_element.find(".//Item.name")
            if item_name is not None:
                new_name = " ".join(base_name.split()[:-1] + [str(current_count + i)])
                item_name.text = new_name
            duplicates.append(new_element)
        for new_element in reversed(duplicates):
            parent.insert(list(parent).index(element) + 1, new_element)

def update_actions_for_combination(combination_name, new_force_name, new_moment_name, root):
    combination = root.find(f".//CombinationOfActions[Item.name='{combination_name}']")
    if combination is not None:
        for action in combination.findall("CombinationOfActions.action"):
            if action.text == "Force 1":
                action.text = new_force_name
            elif action.text == "Moment 1":
                action.text = new_moment_name
    else:
        #print(f"Combination {combination_name} not found")
        pass
def update_combination_name_in_situations(situation_name, old_name, new_name, root):
    situ = root.find(f".//Situation[Item.name='{situation_name}']")
    if situ is not None:
        for element in situ.findall("ConstructionStage.element"):
            if element.attrib.get('name') == old_name:
                element.set('name', new_name)

def update_comandstage_calculation(Item_name, old_name, new_name, root):
    com = root.find(f".//BoundaryElementAnalysis[Item.name='{Item_name}']")
    if com is not None:
        for stage in com.findall("Command.stage"):
            if stage.text == old_name:
                stage.text = new_name

def update_combination_name(old_name, new_name, root):
    combination = root.find(f".//CombinationOfActions[Item.name='{old_name}']")
    if combination is not None:
        item_name = combination.find("Item.name")
        if item_name is not None:
             item_name.text = new_name
    else:
        pass
        #print(f"Combination {old_name} not found")
        
def duplicate_specific_combination_actions(xml_string, n, predefined_names,y):
    # Parse the XML string
    #root = ET.fromstring(xml_string)
    #i=2
    # Find the specific CombinationOfActions with Item.name 'Set B LC-6'
    for combo in root.findall('.//CombinationOfActions'):
        item_name = combo.find('Item.name')
        if item_name is not None and item_name.text == y:
            actions = combo.findall('CombinationOfActions.action')
           # print(type(actions))
            o=0
            for action in actions:
                #n=1
                
                #for i in range(0,n+1):
                    # Create a new action element with a predefined name if available
                new_action = ET.Element('CombinationOfActions.action')
                    
                new_action.text = lst[o%2]
                o=o+1
                    #i+=1
                combo.append(new_action)
    
    # Convert the modified XML back to a string
    return ET.tostring(root, encoding='unicode')
    
    # Convert the modified XML back to a string
    #return ET.tostring(root, encoding='unicode')     



# Main Processing
for file_index in range(234, 236):
    xml_file = str(df.iloc[file_index, 17])
    print(f"Processing file: {xml_file}")

    # Search for the XML file in the directory
    for root_dir, dirs, files in os.walk(NEW_DIRECTORY):
        if xml_file in files:
            print(f"Found file: {xml_file}")
            break
    else:
        print("Empty cell in excel spread sheet")
        continue

    # Parse XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()
    print("XML file is processing")

    # Update <Item.name> elements
    updates = [
        ('.//Scenarios/Situations/Situation/Item.name', df.iloc[3, 23]),
        ('.//Calculations/PileCalculations/BoundaryElementAnalysis/Item.name', df.iloc[2, 23]),
        ('.//Calculations/PileCalculations/BoundaryElementAnalysis/Command.stage', df.iloc[2, 23])
    ]
    for xpath, new_text in updates:
        for element in root.findall(xpath):
            element.text = new_text

    # Duplicate and insert elements
    concentrated_actions = root.find(".//ConcentratedActions")
    if concentrated_actions:
        num_copies = df.iloc[file_index, 15] * df.iloc[file_index, 16] - 1
        duplicate_and_insert(concentrated_actions, "ForceAction", num_copies)
        duplicate_and_insert(concentrated_actions, "MomentAction", num_copies)


    #num_copies=3
    #for ac in root.findall(".//CombinationsOfActions/CombinationOfActions/CombinationOfActions.action"):
        #duplicate_and_insert(ac, "CombinationOfActions.action", num_copies)
    
    #modified_xml = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    #n = 3
    #predefined_names = [
        #'Force Action 1',
        #'Force Action 2',
        #'Force Action 3'
    #]
    
    # Call the function and print the result
    #modified_xml = duplicate_specific_combination_actions(modified_xml, n, predefined_names)
    
    modified_xml = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    #print(modified_xml)
        
    num_copies = df.iloc[file_index, 16] - 1
    for coa in root.findall(".//CombinationsOfActions"):
        duplicate_and_insert(coa, "CombinationOfActions", num_copies)
    
    
        
    situations = root.find(".//Situations")
    if situations:
        duplicate_and_insert(situations, "Situation", num_copies)

    pile_calculations = root.find(".//PileCalculations")
    if pile_calculations:
        duplicate_and_insert(pile_calculations, "BoundaryElementAnalysis", num_copies)
    
    modified_xml = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    #print(modified_xml)
    
    # Update forces and moments
    num_copies = 5
    for j in range(0, df.iloc[file_index, 15]):
        for i in range(0 + 6 * j, num_copies + 1 + 6 * j):
            # Force Action Updates
            force_action = root.find(f".//ForceAction[Item.name='Force {i + 1}']")
            if force_action:
                force_action.find("./SpaceItem.x").set('value', str(df.iloc[file_index + j, 12]))
                force_action.find("./SpaceItem.y").set('value', str(df.iloc[file_index + j, 13]))
                force_action.find("./SpaceItem.depth").set('value', str(df.iloc[file_index + j, 14]))
                force_action.find("./ForceAction.force_x").set('value', str(df.iloc[file_index + j, 23 + 6 * (i - 6 * j)]))
                force_action.find("./ForceAction.force_y").set('value', str(df.iloc[file_index + j, 24 + 6 * (i - 6 * j)]))
                force_action.find("./ForceAction.force_z").set('value', str(df.iloc[file_index + j, 25 + 6 * (i - 6 * j)]))
                force_action.find("./Item.name").text = f'Force Column {df.iloc[file_index + j, 11]} {df.iloc[7, 23 + 6 * (i - 6 * j)]}'

            # Moment Action Updates
            moment_action = root.find(f".//MomentAction[Item.name='Moment {i + 1}']")
            if moment_action:
                moment_action.find("./SpaceItem.x").set('value', str(df.iloc[file_index + j, 12]))
                moment_action.find("./SpaceItem.y").set('value', str(df.iloc[file_index + j, 13]))
                moment_action.find("./SpaceItem.depth").set('value', str(df.iloc[file_index + j, 14]))
                moment_action.find("./MomentAction.moment_x").set('value', str(df.iloc[file_index + j, 26 + 6 * (i - 6 * j)]))
                moment_action.find("./MomentAction.moment_y").set('value', str(df.iloc[file_index + j, 27 + 6 * (i - 6 * j)]))
                moment_action.find("./MomentAction.moment_z").set('value', str(df.iloc[file_index + j, 28 + 6 * (i - 6 * j)]))
                moment_action.find("./Item.name").text = f'Moment Column {df.iloc[file_index + j, 11]} {df.iloc[7, 23+ 6 * (i - 6 * j)]}'

            Force_name = force_action.find("./Item.name").text if force_action else ""
            Moment_name = moment_action.find("./Item.name").text if moment_action else ""

            # Combination of Actions Updates
            combination_action = root.find(f".//CombinationOfActions[Item.name='Combination {(i - 6 * j) + 1}']")
            if combination_action:
                #combination_action.find("./SpaceItem.x").set('value', str(df.iloc[file_index, 11]))
                #combination_action.find("./SpaceItem.y").set('value', str(df.iloc[file_index, 12]))
                combination_action.find("./SpaceItem.depth").set('value', str(df.iloc[file_index, 14]))

            update_actions_for_combination(f'Combination {(i - 6 * j) + 1}', Force_name, Moment_name, root)
            update_combination_name(f'Combination {(i - 6 * j) + 1}', str(df.iloc[7, 23 + 6 * (i - 6 * j)]), root)
            update_combination_name_in_situations(f'SET {(i - 6 * j) + 1}', 'Combination 1', str(df.iloc[7, 23 + 6 * (i - 6 * j)]), root)
            update_comandstage_calculation(f'Calculation {(i - 6 * j) + 1}', 'Calculation 1', f'SET {(i - 6 * j) + 1}', root)
            
            # Parameters
            #n = df.iloc[file_index,14]-1
            
             
            #root = ET.Element("CombinationsOfActions")
            #for l in range (1,6):
                
            if j < df.iloc[file_index, 15]-1:
                
                #for p in range(0,df.iloc[file_index,14]-1):
                lst=[f'Force Column {j+1} Set B LC-{i+1-6*j}',f'Moment Column {j+1} Set B LC-{i+1-6*j}']
                modified_xml = duplicate_specific_combination_actions(modified_xml, df.iloc[file_index,15]-1, lst,str(df.iloc[7, 23 + 6 * (i - 6 * j)]))
                    #modified_xml = duplicate_specific_combination_actions(modified_xml, 1, f'Moment Column {p+2} Set B LC-{i}',str(df.iloc[7, 17 + 6 * (o - 6 * p)]))
                #create_combination_of_actions(root, "Set B LC-2", n)
            #xml_str = prettify_xml(root)
            

    # Write the modified XML to a string
    modified_xml = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    # print(modified_xml)
    
    # Create a new Repute element with attributes
    new_root = ET.Element('Repute', 
                          version="2.5", 
                          release="11", 
                          build="21071", 
                          copyright="©2002-21 Geocentrix Ltd. All rights reserved",
                          licence="RPT-0250-f7fc-00be-6cc8-Ent",
                          xmlns_xsi="http://www.w3.org/2001/XMLSchema-instance",
                          xsi_noNamespaceSchemaLocation="https://www.geocentrix.co.uk/xml-schemas/Repute2.xsd")

    # Copy all children of the original Repute element to the new Repute element
    for child in root:
        new_root.append(child)

    # Create ElementTree with the new root
    tree = ET.ElementTree(new_root)

    # Write the modified XML content to the specified file path
    file_path = os.path.join(NEW_DIRECTORY, xml_file)
    with open(file_path, 'wb') as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        tree.write(f, encoding='utf-8')

    print(f'Modified XML content has been written to {file_path}')
            
