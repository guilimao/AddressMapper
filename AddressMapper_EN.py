import pandas as pd
import requests
import hashlib
import time
from urllib.parse import urlencode
import pandas as pd
import simplekml
from tqdm import tqdm

def excel_to_kml(input_excel, output_kml):
    """
    Convert an Excel file containing latitude and longitude information to KML format
    :param input_excel: Path to the input Excel file
    :param output_kml: Path to the output KML file
    """
    try:
        # Read Excel data
        df = pd.read_excel(input_excel, engine='openpyxl')
        
        # Create KML object
        kml = simplekml.Kml()
        
        # Count valid data
        valid_count = 0
        skipped = []
        
        # Progress bar
        with tqdm(total=len(df), desc="Conversion Progress") as pbar:
            for index, row in df.iterrows():
                try:
                    # Check coordinate validity
                    lng = float(row['Longitude'])
                    lat = float(row['Latitude'])
                    
                    # Create placemark
                    point = kml.newpoint(
                        name=row['Company Name'],
                        description=f"""
                        Contact: {row['Contact Name']}
                        Address: {row['Company Address']}
                        Parsing Result: {row.get('Parsing Result', '')}
                        """,
                        coords=[(lng, lat)]
                    )
                    
                    # Set style
                    point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pal4/icon28.png'
                    valid_count += 1
                except (ValueError, KeyError):
                    skipped.append(index + 2)  # Record Excel row number (including header)
                pbar.update(1)
        
        # Save KML file
        kml.save(output_kml)
        
        print(f"\nConversion complete! Valid markers: {valid_count}")
        if skipped:
            print(f"Skipped invalid data rows (Excel row numbers): {skipped}")
    
    except Exception as e:
        print(f"Conversion failed: {str(e)}")
        
def read_excel_data(file_path):
    """Read the new three-column Excel file (handle empty contacts)"""
    columns = [
        'Company Name',
        'Company Address',
        'Contact Name'
    ]
    
    try:
        df = pd.read_excel(
            file_path,
            engine='openpyxl',
            header=None,
            names=columns,
            skiprows=0,
            keep_default_na=False  # New: do not convert empty values to NaN
        )
        # Convert empty strings to None
        df = df.replace(r'^\s*$', None, regex=True)
        return df.to_dict(orient='records')
    
    except Exception as e:
        print(f"Error reading file: {str(e)}")
        return []

def get_coordinates(company):
    """Call the Tencent Maps API stored in a text file (adapted for new address field)"""
    try:
        with open('tencent_api.txt', 'r') as file:
            API_KEY = file.readline().strip()
            API_SECRET = file.readline().strip()
    except FileNotFoundError:
        print("API key file not found. Please ensure 'tencent_api.txt' exists.")
        return None, None
    
    # Use the full address field directly
    full_address = company['Company Address']
    
    params = {
        "address": full_address,
        "key": API_KEY
    }

    try:
        sorted_params = sorted(params.items(), key=lambda x: x[0])
        param_pairs = [f"{k}={v}" for k, v in sorted_params]
        raw_param_str = "&".join(param_pairs)
        
        sign_str = f"/ws/geocoder/v1?{raw_param_str}{API_SECRET}"
        sig = hashlib.md5(sign_str.encode('utf-8')).hexdigest().lower()
        
        encoded_params = []
        for k, v in sorted_params:
            encoded_value = requests.utils.quote(v, safe='')
            encoded_params.append(f"{k}={encoded_value}")
        
        encoded_params.append(f"sig={sig}")
        url = f"https://apis.map.qq.com/ws/geocoder/v1?{'&'.join(encoded_params)}"
        
        headers = {'x-legacy-url-decode': 'no'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        data = response.json()
        
        if data.get('status') == 0:
            location = data['result']['location']
            return {
                'Longitude': location['lng'],
                'Latitude': location['lat'],
                'Parsing Result': data['result']['title']
            }
        else:
            return {
                'Longitude': '',
                'Latitude': '',
                'Parsing Result': f"Error: {data.get('message', 'Unknown error')}"
            }
    
    except Exception as e:
        print(f"Request failed: {str(e)}")
        return {
            'Longitude': '',
            'Latitude': '',
            'Parsing Result': f"Exception: {str(e)}"
        }

def write_to_excel(data, output_path):
    """Improved export function (handle empty contact display)"""
    try:
        df = pd.DataFrame(data)
        # Convert None to empty string
        df['Contact Name'] = df['Contact Name'].fillna('')
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"Data successfully written to {output_path}")
    except Exception as e:
        print(f"Write failed: {str(e)}")


if __name__ == "__main__":
    input_file = "1.xlsx"
    output_file = "Address_With_GPS.xlsx"
    batch_size = 20

    companies = read_excel_data(input_file)
    
    if not companies:
        print("No valid data read. Please check the input file.")
        exit()
    
    print(f"Processing {len(companies)} records...")
    
    try:
        for idx, company in enumerate(companies, 1):
            print(f"Processing record {idx} ({idx/len(companies):.1%})", end='\r')
            
            coordinates = get_coordinates(company)
            company.update(coordinates)
            
            if idx % batch_size == 0:
                write_to_excel(companies, output_file)
                print(f"\n✅ Successfully saved the first {idx} records")
            
            time.sleep(0.2)
        
        if len(companies) % batch_size != 0:
            write_to_excel(companies, output_file)
            print(f"\n✅ Final save complete (total {len(companies)} records)")
        

        input_excel = output_file
        output_kml = "Address_With_GPS.kml"
        excel_to_kml(input_excel, output_kml)
        print("\nProcessing complete! Result files saved. Import the KML file in Google Earth.")

    except KeyboardInterrupt:
        write_to_excel(companies, output_file)
        print(f"\n⚠️ User interrupted! Saved {len(companies)} processed records")
    except Exception as e:
        write_to_excel(companies, output_file)
        print(f"\n⚠️ Exception occurred: {str(e)}! Saved {len(companies)} processed records")
        raise