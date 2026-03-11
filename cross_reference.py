import pandas as pd
import re

officials_file = 'MA_Municipal_Officials.xlsx'
voters_file = 'MyExport_2820.csv'  # or whatever your current export file is named

try:
    officials = pd.read_excel(officials_file)
    norfolk_officials = officials[officials['County'].str.lower() == 'norfolk'].copy()
    print(f'Total officials in Norfolk County: {len(norfolk_officials)}')
except Exception as e:
    print(f'Error reading officials: {e}')
    exit(1)

try:
    repubs = pd.read_csv(voters_file, dtype=str)
    print(f'Total republicans in CSV: {len(repubs)}')
except Exception as e:
    print(f'Error reading republicans CSV: {e}')
    exit(1)

def clean_name(name):
    if pd.isna(name): return ''
    return re.sub(r'[^a-z]', '', str(name).lower())

repubs['clean_first'] = repubs['FirstName'].apply(clean_name)
repubs['clean_last'] = repubs['LastName'].apply(clean_name)
repubs['clean_city'] = repubs['PrimaryCity'].apply(clean_name)

matched = []

for idx, row in norfolk_officials.iterrows():
    name_str = str(row['Name'])
    name_parts = name_str.lower().split()
    if len(name_parts) < 2:
        continue
    
    first_name_candidate = clean_name(name_parts[0])
    last_name_candidate = clean_name(name_parts[-1])
    city_candidate = clean_name(row['Municipality'])
    
    # Match: First Name, Last Name, and City
    match_cond = (
        (repubs['clean_first'] == first_name_candidate) & 
        (repubs['clean_last'] == last_name_candidate) &
        (repubs['clean_city'] == city_candidate)
    )
    
    matches_found = repubs[match_cond]
    if len(matches_found) > 0:
        match = matches_found.iloc[0]
        row_dict = row.to_dict()
        row_dict['Voter_Address'] = match['PrimaryAddress1']
        row_dict['Voter_City'] = match['PrimaryCity']
        row_dict['Match_Type'] = 'Exact (Name & City)'
        matched.append(row_dict)
    else:
        # Match only First Name and Last Name
        match_cond2 = (
            (repubs['clean_first'] == first_name_candidate) & 
            (repubs['clean_last'] == last_name_candidate)
        )
        matches_found2 = repubs[match_cond2]
        if len(matches_found2) > 0:
            match = matches_found2.iloc[0]
            row_dict = row.to_dict()
            row_dict['Voter_Address'] = match['PrimaryAddress1']
            row_dict['Voter_City'] = match['PrimaryCity']
            row_dict['Match_Type'] = 'Name Only (City Mismatch)'
            matched.append(row_dict)

matched_df = pd.DataFrame(matched)
matched_df.to_csv('norfolk_republican_officials_detailed.csv', index=False)

if len(matched_df) > 0:
    print("\n--- MATCHED OFFICIALS ---")
    for _, r in matched_df.iterrows():
        note = '' if r['Match_Type'] == 'Exact (Name & City)' else f"(Note: Voter lives in {r['Voter_City']})"
        print(f"* {r['Name']} | {r['Municipality']} | {r['Committee']} | {note}".strip(' |'))
else:
    print("No matches found.")
