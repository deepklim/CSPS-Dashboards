import numpy as np
import pandas as pd

# Import latitude and longitude data
geo_data = pd.read_csv('D:\\Maps\\Data\\geo_data.csv', index_col='city_prov')

# Import LSR
lsr = pd.read_csv('D:\\Maps\\Data\\LSR.csv',
                  usecols=['Delivery Type', 'Offering Province', 'Offering City', 'Reg Status', 'Billing Dept Code'])

# Concatenate city and province and delete original columns
lsr['city_prov'] = lsr['Offering City'] + ', ' + lsr['Offering Province']
lsr.drop(labels=['Offering City', 'Offering Province'], axis=1, inplace=True)

# Filters
reg_filter = (lsr['Reg Status'] == 'Confirmed')
deliv_filter = (lsr['Delivery Type'] == 'Classroom') | (lsr['Delivery Type'] == 'Learning Event (In Person)')

# Apply filters and keep only relevant columns
lsr = lsr.loc[reg_filter & deliv_filter, ['Billing Dept Code', 'city_prov']]

# Combine nearby cities
# Convert geo_data to numpy array
index = geo_data.index.values
index.resize(len(index), 1)
values = geo_data.values.astype(float)
geo_data_np = np.concatenate([index, values], axis=1)

# Loop through geo_data_np and replace nearby cities in LSR
for i in range(len(geo_data_np)):
    for j in range(i+1, len(geo_data_np)):
        if abs(geo_data_np[i][1] - geo_data_np[j][1]) < 0.3 and abs(geo_data_np[i][2] - geo_data_np[j][2]) < 0.3:
            lsr['city_prov'].replace(geo_data_np[j][0], geo_data_np[i][0], inplace=True)

# Map city_size to icon color
def icon_color(city_size):
    if city_size>0 and city_size<=10:
        return 'red'
    elif city_size>10 and city_size<=50:
        return 'green'
    elif city_size>50 and city_size<=100:
        return 'blue'
    elif city_size>100 and city_size<=500:
        return 'cyan'
    elif city_size>500:
        return 'magenta'
    else:
        raise ValueError('We\'re having tremendous problems.')

def city_counts(dept_code):
    # Filter LSR for given department
    dept_lsr = lsr.loc[lsr['Billing Dept Code'] == dept_code, 'city_prov']
    
    # Get counts and map city size to icon color
    counts = dept_lsr.value_counts().map(icon_color)
    
    # Convert from series back to dataframe
    counts = pd.concat([counts], axis=1)
    
    # Add latitude and longitude
    counts['Latitude'] = counts.index.to_series().map(geo_data['Latitude'])
    counts['Longitude'] = counts.index.to_series().map(geo_data['Longitude'])
    
    return counts
