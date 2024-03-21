import openpyxl
import subprocess
def ping(ip_address, output_file):
   with open(output_file, 'a') as f:
       f.write(f"Pinging {ip_address}...\n")
       ping_result = subprocess.run(['ping', '-n', '4', ip_address], capture_output=True, text=True)
       f.write(ping_result.stdout)
       f.write("\n\n")
def traceroute(ip_address, output_file):
   with open(output_file, 'a') as f:
       f.write(f"Traceroute to {ip_address}...\n")
       traceroute_result = subprocess.run(['tracert', ip_address], capture_output=True, text=True)
       f.write(traceroute_result.stdout)
       f.write("\n\n")
def main():
   output_file = "ping_traceroute_output.txt"
   # Open the Excel file
   workbook = openpyxl.load_workbook('ip_addresses.xlsx')
   sheet = workbook.active
   # Iterate through each row in the Excel file
   for row in sheet.iter_rows(values_only=True):
       ip_address = row[0]
       # Perform ping
       ping(ip_address, output_file)
       # Perform traceroute
       traceroute(ip_address, output_file)
   print("Process Completed!")
if __name__ == "__main__":
   main()