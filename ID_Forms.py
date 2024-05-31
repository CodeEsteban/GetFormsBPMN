import os
import pandas as pd
from xml.etree import ElementTree as ET
from openpyxl import Workbook

def extract_data_from_bpmn(file_path):
    # Parse the XML file
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Namespace to access elements correctly
    ns = {'bpmn': 'http://www.omg.org/spec/BPMN/20100524/MODEL'}

    # Extract the name of the process
    participant = root.find('.//bpmn:participant', ns)
    process_name = participant.get('name') if participant is not None else 'Unknown'

    # Find all user tasks and their details
    tasks = root.findall('.//bpmn:userTask', ns)
    data = []
    for task in tasks:
        task_name = task.get('name')
        form_ref = task.get('{http://camunda.org/schema/1.0/bpmn}formRef')
        if form_ref:
            data.append({
                'Nombre del Tramite': process_name,
                'Nombre de la Actividad': task_name,
                'ID Formulario': form_ref
            })
    return data

def process_bpmn_folder(folder_path):
    # List to hold all task data from all files
    all_data = []

    # Iterate over all .bpmn files in the given directory
    for filename in os.listdir(folder_path):
        if filename.endswith('.bpmn'):
            file_path = os.path.join(folder_path, filename)
            file_data = extract_data_from_bpmn(file_path)
            all_data.extend(file_data)
    
    # Create DataFrame
    df = pd.DataFrame(all_data)

    # Write to Excel
    df.to_excel('output.xlsx', index=False)

# Example usage
folder_path = r'C:\Documentos_Macarena\Excel_ID_Forms\BPMNMACARENIA'
process_bpmn_folder(folder_path)
