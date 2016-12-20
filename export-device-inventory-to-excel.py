from apicem import *
import sys
import xlsxwriter

device = []
try:
    resp = get(api="network-device")
    status = resp.status_code
    print("Status: ", status)
    response_json = resp.json()
    # all network-device detail is in "response"
    device = response_json["response"]
except:
    print("Something wrong, cannot get network device information")
    sys.exit()

if status != 200:
    print("Response status: %s,Something wrong !" % status)
    print(resp.text)
    sys.exit()

if device == []:  # response is empty, no network-device is discovered.
    print("No network device found !")
    sys.exit()

#Compile list of devices
device_list = []
for item in device:
    device_list.append([item["hostname"],
                        item["managementIpAddress"],
                        item["type"],
                        item['softwareVersion'],
                        item['upTime'],
                        item['macAddress'],
                        item['serialNumber'],
                        item['role'],
                        item['tagCount'],
                        item['series'],
                        item['platformId'],
                        item['reachabilityStatus']])
# Create Excel workbook and sheet 1 spreadsheet.
#Name it Test . It can be renamed later and name can be changed.

workbook = xlsxwriter.Workbook('Test1.xlsx')
worksheet = workbook.add_worksheet('Test1')

#Make bold format to be used for headers
bold = workbook.add_format({'bold': True})

# Format and Add Headers
worksheet.write('A1', 'Device Name', bold)
worksheet.write('B1', 'IP Address', bold)
worksheet.write('C1', 'Type', bold)
worksheet.write('D1', 'Software Version', bold)
worksheet.write('E1', 'Up Time', bold)
worksheet.write('F1', 'Mac Address', bold)
worksheet.write('G1', 'Serial Number', bold)
worksheet.write('H1', 'Role', bold)
worksheet.write('I1', 'Tag Count', bold)
worksheet.write('J1', 'Series', bold)
worksheet.write('K1', 'Platform ID', bold)
worksheet.write('L1', 'Reachability Status', bold)

row = 1
col = 0
#Go through and write lines to excel workbook. Will work to make this its own function so that calling it i simpler. This works for now

for hostname, ip, type, softwareVersion, upTime, macAddress, serialNumber, role, tagcount, series, platformid, reachabilitystatus in (
device_list):
    worksheet.write(row, col, hostname)
    worksheet.write(row, col + 1, ip)
    worksheet.write(row, col + 2, type)
    worksheet.write(row, col + 3, softwareVersion)
    worksheet.write(row, col + 4, upTime)
    worksheet.write(row, col + 5, macAddress)
    worksheet.write(row, col + 6, serialNumber)
    worksheet.write(row, col + 7, role)
    worksheet.write(row, col + 8, tagcount)
    worksheet.write(row, col + 9, series)
    worksheet.write(row, col + 10, platformid)
    worksheet.write(row, col + 11, reachabilitystatus)

    row += 1

workbook.close()

# print (tabulate(device_list, headers=['hostname','ip','type'],tablefmt="rst"))

