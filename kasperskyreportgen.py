import pandas as pd
import datetime
import subprocess
import sys

# Get today's date in the desired format
date = datetime.date.today().strftime('%d-%b-%y')
date2 = datetime.date.today().strftime("%-m-%-d-%Y")

# Run xls_to_xlsx.sh script to convert xls files to xlsx format
process = subprocess.Popen(["/bin/bash", "xls_to_xlsx.sh"])

# Wait for the process to complete and store the return code
return_code = process.wait()

# Check if the script executed successfully
if return_code == 0:
    print("File conversion Successful")
else:
    sys.exit("File conversion Failed")


# Define file paths
file_01 = '~/Documents/logs/kaspersky_av/all_hosts.csv'
file_02 = '~/Documents/logs/kaspersky_av/Report-EDR-Optimum_' + date2 + '.xlsx'
file_03 = '~/Documents/logs/kaspersky_av/Report-EDR_' + date2 + '.xlsx'
new_file_name = '~/Documents/logs/kaspersky_av/kaspersky_EDR_deployment_report_' + date + '_.xlsx'

# Read data into DataFrames
df_edr = pd.read_excel(file_03, sheet_name='Details')
df_optimum = pd.read_excel(file_02, sheet_name="Details")
df_all_hosts = pd.read_csv(file_01, delimiter='\t')

# Select only necessary columns from DataFrames
new_df_optimum = df_optimum.iloc[1:][["Device", "Component status", "Component name"]]
new_df_edr = df_edr[["Device", "Component status", "Component name"]]

# Concatenate DataFrames
vlook_df = pd.concat([new_df_edr, new_df_optimum], ignore_index=True)

# Select only necessary colu*mns from df_all_hosts
selected_columns = [
    "Name", "Operating system type", "Windows domain", "Network Agent is installed",
    "Network Agent is running", "Real-time protection", "Last connected to Administration Server",
    "Protection last updated", "Status", "Status description", "Information last updated",
    "DNS domain", "Total number of threats detected", "Connection IP address",
    "Network Agent version", "Application version", "Anti-virus databases last updated",
    "Description", "Endpoint Sensor status", "Operating system release ID",
    "Name of virtual or secondary Administration Server", "Parent group"
]

# Rename 'Name' column to 'Device'
new_df_all_hosts = df_all_hosts[selected_columns].rename(columns={'Name': 'Device'})

# Merge DataFrames based on 'Device'
merged_df = new_df_all_hosts.merge(vlook_df, on='Device', how='left')

# Extract and insert 'Component name' and 'Component status' columns
merged_df.insert(2, 'Component Name', merged_df.pop('Component name'))
merged_df.insert(3, 'Component Status', merged_df.pop('Component status'))

# Write the merged DataFrame to a new Excel file
merged_df.to_excel(new_file_name, index=False, sheet_name="All-Hosts_" + date, header=True)

# Append EDR details to the same Excel file
with pd.ExcelWriter(new_file_name, engine='openpyxl', mode='a') as writer:
    vlook_df.to_excel(writer, sheet_name="EDR_" + date, index=False)

# Count Section
kaspersky_total_hosts = merged_df.count()["Device"]
EDR_required_hosts = merged_df['Operating system type'].str.contains('Windows').sum()
EDR = merged_df['Component Name'].str.contains('Endpoint').sum()
EDR_not_installed = merged_df['Component Status'].str.contains('Not installed').sum()

EDR_deployed = EDR - EDR_not_installed
EDR_pending = EDR_required_hosts - EDR_deployed

# Print the counts
print(f"Kaspersky Total Hosts: {kaspersky_total_hosts}\n"
      f"EDR Required hosts: {EDR_required_hosts}\n"
      f"EDR Deployed: {EDR_deployed}\n"
      f"EDR Pending: {EDR_pending}")