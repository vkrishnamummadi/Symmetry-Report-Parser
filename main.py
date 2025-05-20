import pandas as pd
import csv

def parse_excel_file(excel_path):
    df = pd.read_excel(excel_path)
    print("\n[DEBUG] Columns in Excel file:", df.columns.tolist())
    print("[DEBUG] First 2 rows from Excel file:")
    print(df.head(2))
    return df

def parse_text_file(text_path):
    parsed_data = []
    current_entry = {}
    current_key = None

    # main_keys = ["Last Name", "First Name", "Card Number", "Access Codes", "Reader Groups", "Readers"]
    # data = []

    with open(text_path, 'r') as file:
        for line in file:
            line = line.rstrip('\n')
            # if not line:
            #     if current_entry:
            #         parsed_data.append(current_entry)
            #         current_entry = {}
            #     continue
            if ':' in line:
                current_key = None
                key, value = line.split(':', 1)
                key = key.strip()
                value = value.strip()
                # Handle multiple Card Numbers per person
                if key == "Card Number":
                    if "Card Number" in current_entry and current_entry["Card Number"]:
                        current_entry["Card Number"] += "; " + value
                    else:
                        current_entry["Card Number"] = value
                else:
                    current_entry[key] = value

            elif line.strip() in ["Access Codes", "Reader Groups", "Readers"]:
                current_key = line.strip()
                current_entry[current_key] = ''

            elif current_key and line.strip():
                # Always treat as string and strip
                if current_entry[current_key]:
                    current_entry[current_key] += '; ' + line.strip()
                else:
                    current_entry[current_key] = line.strip()

            elif not line.strip():
                if current_entry:
                    parsed_data.append(current_entry)
                    current_entry = {}
                current_key = None

        if current_entry:
            parsed_data.append(current_entry)
    # Debug: print first 2 parsed text entries
    print("\n[DEBUG] First 2 parsed text entries:")
    for entry in parsed_data[:2]:
        print(entry)
    return parsed_data

def compare_data(excel_df, text_data):
    matched_data = []
    matches_found = 0

    for _, excel_row in excel_df.iterrows():
        excel_last = str(excel_row['Last Name']).strip().lower()
        excel_first = str(excel_row['First Name']).strip().lower()
        parks_employee = str(excel_row['Staff Confirmations']).strip()
        found_match = False
        for text_entry in text_data:
            text_last = str(text_entry.get('Last Name', '')).strip().lower()
            text_first = str(text_entry.get('First Name', '')).strip().lower()
            # Debug: print names being compared
            # print(f"[DEBUG] Comparing EXCEL: '{excel_first} {excel_last}' <-> TEXT: '{text_first} {text_last}'")
            if (excel_last == text_last and excel_first == text_first and not parks_employee == 'Not applicable - Parks employee'):
                print("[DEBUG] MATCH FOUND!")
                print("[DEBUG] Matched text entry:", text_entry)
                merged = {**excel_row.to_dict(),
                          **{k: text_entry.get(k, '') for k in [
                              'Card Number', 'Access Codes', 'Reader Groups', 'Readers'
                              ]}
                        }
                print("[DEBUG] Merged row to be appended:", merged)
                matched_data.append(merged)
                matches_found += 1
                found_match = True
                break
        if not found_match:
            print(f"[DEBUG] No match found for EXCEL: '{excel_first} {excel_last}'")
    print(f"\n[DEBUG] Total matches found: {matches_found}")
    return matched_data

def multiline_fields(df, columns):
    # Replace ';' with '\n' for multiline display in Excel
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(';', '; \n')
    return df

def write_output(df, output_path):
    df['Name'] = df['First Name'] + ' ' + df['Last Name']
    columns = ['Name', 'Last Name', 'First Name', 'Wisard Emp. Status',
        'SAM Status', 'Department', 'Division', 'Position',
        'BG Level', 'Staff Confirmations', 'Card Number',
        'Access Codes', 'Reader Groups', 'Readers'
    ]

    for col in columns:
        if col not in df.columns:
            df[col] = ''
    
    df = multiline_fields(df, ['Card Number', 'Access Codes', 'Reader Groups', 'Readers'])

    print("\n[DEBUG] Writing the following columns to Excel:", columns)
    print("[DEBUG] First 2 rows to be written:")
    print(df[columns].head(2))
    df[columns].to_excel(output_path, index=False)
    print(f"\nData written to {output_path}")

def main():
    # excel_path = 'Comparable SO Badge Active Facilities Staff List 2025-04-22.xlsx'
    excel_path = r"C:/Users/VamshiM/Documents/Python_Scripts/SymmetryReportParser/Staff_List_04-22_25.xlsx"
    txt_path = 'cards.txt'
    output_file = 'SymmetryReport.xlsx'
    excel_df = parse_excel_file(excel_path)
    print("Columns in Excel file:", excel_df.columns.tolist())
    text_data = parse_text_file(txt_path)
    print("First 2 parsed text entries:", text_data[:2])
    matched_data = compare_data(excel_df, text_data)
    matched_df = pd.DataFrame(matched_data)
    write_output(matched_df, output_file)

if __name__ == "__main__":
    main()
