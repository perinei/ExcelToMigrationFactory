#!/usr/bin/python
import json
import xlrd
import csv
import argparse
import sys


#1
def put_machine_names_from_csvfiles_in_array():
    print("")
    with open('myMachinesFromCloudendure.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        machine_names_in_csv = []
        for row in csv_reader:
            if line_count == 0:
                # print(f'\t{row[2]}')
                line_count += 1
            else:
                # print(f'\t{row[2]}')
                machine_names_in_csv.append(row[2])
                line_count += 1
        print(f'machines in CloudEndure (CSV) = {line_count-1}.')
        print(machine_names_in_csv)
        print("-------------")
        return machine_names_in_csv


#3
def get_field_location(sheet, col_string, row_with_field_names):
    col_num_wave = -1
    # print(sheet.ncols)
    for col in range(sheet.ncols):
        if sheet.cell_value(row_with_field_names, col) == col_string:
            col_num = col
            print(f'{col_string} name found in column {col_num}')
            break
    if col_num == -1:
        print(f'{col_string} name not found.')
    
    print("----------")
    return col_num


#4
def put_server_names_from_excelfile_in_array(sheet, servers_col, waves_col, wave_name, row_with_field_names):
    machine_name_in_excel = []
    line_count = 0
    print("Total rows in the file = ", sheet.nrows)
    for nrow in range(row_with_field_names+1, sheet.nrows):
        if str(sheet.cell_value(nrow, waves_col)) == wave_name:
            machine_name_in_excel.append(str(sheet.cell_value(nrow, servers_col)))
            line_count += 1
    print (f'machines in Excel on wave {wave_name} = {line_count}')
    print (machine_name_in_excel)
    print("------")
    return (machine_name_in_excel)

#5
def compare_arrays_of_machine_names(csv, excel):
    update_machines = []
    item_count = 0
    for machine_excel in excel:
        for machine_csv in csv:
            if machine_csv == machine_excel:
                update_machines.append(machine_csv)
                item_count += 1
                break
    print(f'# of machines that match Excel file with CloudEndure = {item_count}')
    # create_csv(update_machines)
    print (update_machines)
    print("------")
    return update_machines


#6
def create_csv(machines, sheet, app_col, servers_col, first_tag_col, last_tag_col, row_with_field_names):

    export_csv = "CE-blueprint.csv"
    file = open(export_csv, 'w+')
    # first row
    file.write('wave_id,app_name,cloudendure_projectname,aws_accountid,server_name,server_os,server_os_version,server_fqdn,server_tier,server_environment,subnet_IDs,privateIPs,securitygroup_IDs,subnet_IDs_test,securitygroup_IDs_test,iamRole,instanceType,tenancy\n')
    line_count = 0
    # print(machines)
    for machine in machines:
        file.write(f'111,{get_field_value(sheet, machine, app_col, servers_col)},333,444,{machine},"","","","{format_tags(sheet, machine, first_tag_col, last_tag_col, servers_col)}","","","","","","",,,""\n')

#6.5
def get_field_value(sheet, machine, field_col, server_col):
    for row in range(sheet.nrows):
        if sheet.cell_value(row, server_col) == machine:
            row_machine = row
            break
    value = str(sheet.cell_value(row_machine, field_col))
    return value


# 7
def format_tags(sheet, machine, first_tag_col, last_tag_col, servers_col):
    # print('# of columns: ', sheet.ncols)
    
    for row in range(sheet.nrows):
        if sheet.cell_value(row, servers_col) == machine:
            row_machine = row
            # print("col_machine: " + str(col_machine))
            break
        
    string_tag = "["
    # add key value to string_tag
    for tag_num in range(first_tag_col, last_tag_col+1):
        string_tag += "{'key':'"
        string_tag += str(sheet.cell_value(1, tag_num))
        string_tag += "','value':'"
        string_tag += str(sheet.cell_value(row_machine, tag_num))
        string_tag += "'}"
        if tag_num < last_tag_col:
            string_tag += ","

    string_tag += "]"

    return string_tag 
#### excelToCE end ####




###################################################################################################



def main(args):
	print(chr(27) + "[2J")
	print("Excel To Migration Factory")
	while (args.wave is None) or (args.wave == ''):
		args.wave = input("Enter wave. Ex. R02W03:")

# This main function gets all the servers and project in an account and create a csv file to be
# used as a template input file to update the blueprints
# 
# Usage: main(args)


	###### ExcelToCE start ########

	with open('config.json') as json_file:
		data = json.load(json_file)
		for param in data['config']:
			wb_machine_names = xlrd.open_workbook(param['Excel_File_Name'])
			sheet = wb_machine_names.sheet_by_name(param['Excel_Tab_Name'])
			row_with_field_names = param['Row_With_Field_Names'] - 1
			server_col_string = param['Server_Column_Name']
			wave_col_string = param['Wave_Column_Name']
			app_name_string = param['Application_Column_Name']
			first_tag_string = param['First_Tag_Name']
			last_tag_string = param['Last_Tag_Name']

	#1
	# csv_machines = put_machine_names_from_csvfiles_in_array()
	# servers_col = get_server_location(sheet, server_col_string, row_with_field_names)
	#3
	servers_col = get_field_location(sheet, server_col_string, row_with_field_names)
	waves_col = get_field_location(sheet, wave_col_string, row_with_field_names)
	first_tag_col = get_field_location(sheet, first_tag_string, row_with_field_names)
	last_tag_col = get_field_location(sheet, last_tag_string, row_with_field_names)
	app_col = get_field_location(sheet, app_name_string, row_with_field_names)
	#4
	excel_servers = put_server_names_from_excelfile_in_array(sheet, servers_col, waves_col, args.wave, row_with_field_names)
	#5
	# machines = compare_arrays_of_machine_names(csv_machines, excel_servers)
	#6
	create_csv(excel_servers, sheet, app_col, servers_col, first_tag_col, last_tag_col, row_with_field_names)

###### excelToCE end ########

if __name__ == '__main__':

	parser = argparse.ArgumentParser()
	# parser.add_argument('-u', '--user', required=False, help='User name')
	# parser.add_argument('-p', '--password', required=False, help='Password')
	# ##### excelToCE start ######
	# parser.add_argument('-t', '--task', required=False, help='<add>, <del> or <dryrun>')
	parser.add_argument('-w', '--wave', required=False, help='example R0xWxx, Pilot or <all> for all server on spreadsheet')
	# ###### excelToCE end ########
	parser.add_argument('-o', '--outputfile', required=False, help='Output CSV file for backup before change')
	main(args = parser.parse_args())