import pandas as pd
import openpyxl 
text_file_path = "add file"

def txt2xlsx(text_file_path):

    # Open .txt file
    with open(text_file_path, 'r') as file:
        input_text = file.readlines()

    # Obtaining Job Name
    job_name = None
    for line in input_text:
        if "Job name: " in line:
            job_name = line.strip().split(": ")[-1] 
            break

    # Initiating Excel Writer
    writer = pd.ExcelWriter(f"{job_name}.xlsx", engine='openpyxl')

    # Lists to sort and categorize data from text files to 
    node_data = []
    constrained_adaptivity_data = []
    initial_stress_data = []
    element_data = []

    # Process txt file for data tables
    current_section = None
    for line in input_text:
        if "*NODE" in line:
            current_section = "Node"

        elif "*CONSTRAINED_ADAPTIVITY" in line:
            current_section = "constrained_adaptivity"

        elif "*INITIAL_STRESS" in line:
            current_section = "initial_stress"

        elif"*ELEMENT" in line:
            current_section="element"

        elif current_section is not None:
            data = line.strip().split()
            if data:
                if current_section == "Node":
                    node_data.append(data)

                elif current_section == "constrained_adaptivity":
                    constrained_adaptivity_data.append(data)

                elif current_section == "initial_stress":
                    initial_stress_data.append(data)

                elif current_section == "element":
                    element_data.append(data)


    #List to DataFrame
    df_node = pd.DataFrame(node_data)
    df_constrained_adaptivity = pd.DataFrame(constrained_adaptivity_data)
    df_element = pd.DataFrame(element_data)
    df_initial_stress = pd.DataFrame(initial_stress_data)

    #Cleaning to remove the trailing text that is extra 
    df_node_cleaned = df_node.iloc[:-3]
    df_constrained_adaptivity_cleaned = df_constrained_adaptivity.iloc[:-3]
    df_element_cleaned = df_element.iloc[:-3]
    df_initial_stress_cleaned = df_initial_stress.iloc[:-1]

    # Using the Excel writer to write data in one excel file with different sheets
    df_node_cleaned.to_excel(writer, sheet_name='Node', index=False, header=False)
    df_constrained_adaptivity_cleaned.to_excel(writer, sheet_name='Constraint', index=False,header=False)
    df_element_cleaned.to_excel(writer, sheet_name='Elements', index=False,header=False)
    df_initial_stress_cleaned.to_excel(writer, sheet_name='Stress', index=False,header=False)

    # Save Excel File
    writer._save()
    print(f"Excel file created: {job_name}.xlsx")
txt2xlsx(text_file_path)
