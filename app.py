import os
import re
import pandas as pd
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

OFFICIALS_FILE = r'MA_Municipal_Officials.xlsx'
VOTERS_FILE = r'Full List.csv'

def clean_name(name):
    if pd.isna(name): return ''
    return re.sub(r'[^a-z]', '', str(name).lower())

# Cache the dataframes so we don't reload 40MB+ on every request if possible
# (For a real production app we'd use a database, but this is a local tool)
data_cache = {
    'officials': None,
    'voters': None
}

def load_officials():
    if data_cache['officials'] is None:
        print("Loading officials Excel...")
        officials = pd.read_excel(OFFICIALS_FILE)
        # Drop rows with entirely null essential data
        officials = officials.dropna(subset=['Name'])
        # Add tracking of original index for sorting if needed, or just sort
        data_cache['officials'] = officials
    return data_cache['officials']

def load_voters():
    if data_cache['voters'] is None:
        print("Loading voters CSV (This might take a moment)...")
        # Load only necessary columns to save memory
        usecols = ['FirstName', 'LastName', 'MiddleName', 'SuffixName', 'PrimaryAddress1', 'PrimaryCity', 'PrimaryPhone', 'StateVoterId']
        # The file might have different columns, let's load what we need, but safely
        # We'll just read strings
        voters = pd.read_csv(VOTERS_FILE, dtype=str, usecols=lambda c: c in usecols or c.startswith('Primary'))
        
        print("Pre-cleaning voter names...")
        voters['clean_first'] = voters.get('FirstName', '').apply(clean_name)
        voters['clean_last'] = voters.get('LastName', '').apply(clean_name)
        voters['clean_city'] = voters.get('PrimaryCity', '').apply(clean_name)
        
        data_cache['voters'] = voters
    return data_cache['voters']


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/counties', methods=['GET'])
def get_counties():
    try:
        officials = load_officials()
        # Drop NaN and get unique counties, filter out miscellaneous rows, and sort them
        counties = sorted([
            str(c) for c in officials['County'].dropna().unique() 
            if str(c).strip() and "key resources" not in str(c).lower()
        ])
        return jsonify({"success": True, "counties": counties})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/api/filter', methods=['POST'])
def filter_data():
    try:
        data = request.json
        selected_counties = data.get('counties', []) # List of strings or 'All'
        only_republicans = data.get('onlyRepublicans', False)
        
        officials = load_officials()
        
        # Step 1: Filter by county
        if 'All' in selected_counties or not selected_counties:
            filtered_officials = officials.copy()
        else:
            # Case insensitive match for counties
            selected_counties_lower = [c.lower() for c in selected_counties]
            filtered_officials = officials[officials['County'].astype(str).str.lower().isin(selected_counties_lower)].copy()
        
        results = []
        
        # Step 2: Cross reference if required
        filtered_officials = filtered_officials.fillna('')
        
        if only_republicans:
            voters = load_voters()
            
            for idx, row in filtered_officials.iterrows():
                name_str = str(row['Name'])
                name_parts = name_str.lower().split()
                if len(name_parts) < 2:
                    continue
                
                first_name_candidate = clean_name(name_parts[0])
                last_name_candidate = clean_name(name_parts[-1])
                city_candidate = clean_name(row.get('Municipality', ''))
                
                # Match
                match_cond = (
                    (voters['clean_first'] == first_name_candidate) & 
                    (voters['clean_last'] == last_name_candidate) &
                    (voters['clean_city'] == city_candidate)
                )
                
                matches_found = voters[match_cond]
                if len(matches_found) > 0:
                    match = matches_found.iloc[0].fillna('')
                    
                    # Carefully format the name
                    first = str(match.get('FirstName', '')).strip()
                    mid = str(match.get('MiddleName', '')).strip()
                    last = str(match.get('LastName', '')).strip()
                    full_name = f"{first} {mid} {last}".replace('nan', '').replace('  ', ' ').strip().title()
                    
                    phone = str(match.get('PrimaryPhone', 'N/A'))
                    if not phone or phone.lower() == 'nan':
                        phone = 'N/A'
                        
                    address = str(match.get('PrimaryAddress1', 'Unknown')).title()
                    if not address or address.lower() == 'nan':
                        address = 'Unknown'
                        
                    voter_id = str(match.get('StateVoterId', ''))
                    if not voter_id or voter_id.lower() == 'nan':
                        voter_id = ''

                    # Append enriched row
                    results.append({
                        'County': str(row.get('County', '')),
                        'Municipality': str(row.get('Municipality', '')),
                        'Committee': str(row.get('Committee', '')),
                        'Name': full_name,
                        'Role': str(row.get('Role', '')),
                        'Address': address,
                        'Phone': phone,
                        'StateVoterId': voter_id
                    })
        else:
            # Just return the filtered officials
            for idx, row in filtered_officials.iterrows():
                results.append({
                    'County': str(row.get('County', '')),
                    'Municipality': str(row.get('Municipality', '')),
                    'Committee': str(row.get('Committee', '')),
                    'Name': str(row.get('Name', '')).title(),
                    'Role': str(row.get('Role', '')),
                    'Address': '',
                    'Phone': '',
                    'StateVoterId': ''
                })
                
        # Step 3: Sort the results
        # Alphabetical by County -> Municipality -> Last Name
        # We need a small helper to extract last name for sorting
        def get_last_name(r):
            parts = str(r['Name']).split()
            return parts[-1].lower() if parts else ''
            
        # First sort by Last Name, then Municipality, then County (Python's sort is stable)
        results.sort(key=get_last_name)
        results.sort(key=lambda x: str(x['Municipality']).lower())
        results.sort(key=lambda x: str(x['County']).lower())
        
        return jsonify({
            "success": True, 
            "count": len(results),
            "data": results
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)})

if __name__ == '__main__':
    # Ensure templates and static folders exist
    os.makedirs('templates', exist_ok=True)
    os.makedirs('static', exist_ok=True)
    app.run(debug=True, port=5000)
