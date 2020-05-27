#!/usr/bin/python

import argparse
import requests
import json
import sys
# ExcelToCE start #
import xlrd
import csv
import argparse
import sys
# ExcelToCe end #
from datetime import datetime
import logging

# to delete files after done"
import os

import getpass

logging.basicConfig(filename='/Users/perineia/Documents/Carrier/cloudendure/logs/CE_Update_Blueprints.log', format='%(asctime)s %(levelname)s:%(message)s', level=logging.DEBUG)
log = logging.getLogger(__name__)

TITLES = [('projectName', 'targetCloud', 'machineName', 'iamRole', 'privateIPs', 'placementGroup', 'staticIp', 'tags', 'publicIPAction',
			'disks', 'instanceType', 'securityGroupIDs', 'staticIpAction', 'subnetIDs', 'subnetsHostProject', 'privateIPAction', 'runAfterLaunch', 'tenancy')]


AZURE_KEYS = ('N/A','','N/A','N/A','N/A','N/A','N/A','','','N/A','','N/A','','N/A','N/A')
GCP_KEYS = ('N/A','','N/A','N/A','N/A','N/A','','','N/A','N/A','','','','N/A','N/A') #machine name
AWS_KEYS = ('','','','','','','','','','','','N/A','','','')

# List of keys avilable for each target cloud
AZURE_KEYS2 = ['instanceType','subnetIDs','securityGroupIDs','privateIPAction', 'privateIPs']
GCP_KEYS2 = ['instanceType','machineName','subnetIDs','privateIPAction', 'privateIPs','disks', 'subnetsHostProject']
AWS_KEYS2 = ['instanceType','subnetIDs','privateIPAction', 'privateIPs','disks','iamRole', 'placementGroup', 'staticIp', 'tags', 
			'publicIPAction', 'securityGroupIDs', 'staticIpAction', 'runAfterLaunch', 'tenancy']

HOST = 'https://console.cloudendure.com'
ENDPOINT = '/api/latest/{}'


###################################################################################################
def _dump_csv(rows, output):

# This function write the output csv file
# 
# Usage: _dump_csv(rows, output):
# 	'rows' 		the data to be written to the fils
#	'output'  	the output file
# 
# Returns: 	None
	print ('Creating csv template...')
	for row in rows:
		for value in row:
			try:
				output.write('{},'.format(value))
			except:
				print (value)
				pass
		output.write('\n')
		
		
###################################################################################################
def _get_cloud_ids(session):

# This function login into CloudEdnure API
# 
# Usage: _get_cloud_ids(session) 	
# 
# Returns: 	a dictionary with CloudId and keys.

	print ('Getting cloud ids...')
	clouds_resp = session.get(url=HOST+ENDPOINT.format('clouds'))	
	clouds = json.loads(clouds_resp.content)['items']
	
	cloud_ids = {}
	for cloud in clouds:
		if cloud['name'] == 'GCP':
			cloud_ids[cloud['id']] = ('GCP', GCP_KEYS)
		elif cloud['name'] == 'AWS':
			cloud_ids[cloud['id']] = ('AWS', AWS_KEYS)
		elif cloud['name'] == 'AZURE_ARM':
			cloud_ids[cloud['id']] = ('AZURE_ARM', AZURE_KEYS)
	
	return cloud_ids
	
###################################################################################################
def _login(args):

# This function login into CloudEdnure API
# 
# Usage: _login(args) 	
# 
# Returns: 	None
	global ENDPOINT
	
	session = requests.Session()
	session.headers.update({'Content-type': 'application/json', 'Accept': 'text/plain'})
	print ('Logging in...')
	resp = session.post(url=HOST+ENDPOINT.format('login'), 
						data=json.dumps({'username': args.user, 'password': args.password}))
	if resp.status_code != 200 and resp.status_code != 307:
		print ('Could not login!')
		sys.exit(2)
		
	# check if need to use a different API entry point
	if resp.history:
		print ('URL Redirected...')
		ENDPOINT = '/' + '/'.join(resp.url.split('/')[3:-1]) + '/{}'
		resp = session.post(url=HOST+ENDPOINT.format('login'),
						data=json.dumps({'username': args.user, 'password': args.password}))

	if session.cookies.get('XSRF-TOKEN'):
		session.headers['X-XSRF-TOKEN'] = session.cookies.get('XSRF-TOKEN')	
		
	return session

##### ExcelToCE ######
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


#2
def get_server_location(sheet, server_col_string, row_with_field_names):
    col_num_machine = -1
    # print(sheet.ncols)
    for col in range(sheet.ncols):
        if sheet.cell_value(row_with_field_names, col) == server_col_string:
            col_num_machine = col
            # print("Server name in column ", col_num_machine)
            break
    if col_num_machine == -1:
        print("Server name not found.")
        sys.exit()  
    print("----------")
    return col_num_machine

#3
def get_wave_location(sheet, wave_col_string, row_with_field_names):
    col_num_wave = -1
    # print(sheet.ncols)
    for col in range(sheet.ncols):
        if sheet.cell_value(row_with_field_names, col) == wave_col_string:
            col_num_wave = col
            # print("Wave name in column ", col_num_wave)
            break
    if col_num_wave == -1:
        print("Wave name not found.")
    
    print("----------")
    return col_num_wave


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
def create_csv(machines, sheet, task, servers_col, first_tag, last_tag, row_with_field_names):
    if task == "add":
        
        # identify where is the first col TAGS
        first_tag_col = -1
        for col in range(sheet.ncols):
            if sheet.cell_value(row_with_field_names, col) == first_tag:
                first_tag_col = col
                print(f'First Tag ({first_tag}) is in column {first_tag_col}')
                break
        if first_tag_col == -1:
            print("First tag not found")  
            sys.exit()  

        # identify where is the last col TAGS
        last_tag_col = -1
        for col in range(sheet.ncols):
            if sheet.cell_value(row_with_field_names, col) == last_tag:
                last_tag_col = col
                print(f'Last Tag ({last_tag}) is in column {last_tag_col}')
                break
        if last_tag_col == -1:
            print("Last tag not found")   
            sys.exit()  

    export_csv = "CE-blueprint.csv"
    file = open(export_csv, 'w+')
    # first row
    file.write('"projectName","targetCloud","machineName","iamRole","privateIPs","privateIPAction","placementGroup","staticIp","tags","publicIPAction","disks","instanceType","securityGroupIDs","staticIpAction","subnetIDs","subnetsHostProject","runAfterLaunch","tenancy"\n')
    line_count = 0
    # print(machines)
    for machine in machines:
        with open('myMachinesFromCloudendure.csv') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            for row in csv_reader:
                # print(f'\t{row[2]}')
                if machine == row[2]:
                    if task == 'add':
                        file.write(f'"{row[0]}","{row[1]}","{row[2]}","","[]","","","","{format_tags(sheet, machine, first_tag_col, last_tag_col, servers_col)}","","","","","","",,,""\n')
                    if task == 'del':
                        file.write(f'"{row[0]}","{row[1]}","{row[2]}","","[]","","","","[]","","","","","","",,,""\n')
                    line_count += 1   
                    break 


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

def _read_blueprints_csv(input_file):

# This function read the new blueprints config CSV file
# 
# Usage: _read_blueprints_csv(input_file)
# 	'input_file' user input for the CSV file containing new blueprint settings to apply
# 	
# 
# Returns: 	an array of blueprints read from the CSV file

	result = []
	with open(input_file, mode='r') as infile:
		reader = csv.DictReader(infile)
		for row in reader:
			result.append(row)
	return result

###################################################################################################
def _write_blueprints_csv(output_file, blueprints):

# This function write the output csv file
# 
# Usage: _write_blueprints_csv(output_file, blueprints)
# 	'output_file' user output for the CSV filename to create with the current blueprint settings 
#				 as backup before changing
#	'bluprints'  existing blueprints data read before applying the new changes
# 	
# 
# Returns: 	None

	with open(output_file, mode='wb') as outfile:
		field_names = blueprints[0].keys()
		# Move machineName to be first
		field_names.remove('machineName')
		field_names.insert(0, 'machineName')
		if 'dedicatedHostIdentifier' not in field_names:
			field_names.append('dedicatedHostIdentifier')
		writer = csv.DictWriter(outfile, fieldnames = field_names)
		writer.writeheader()
		for bp in blueprints:
			writer.writerow(bp)
			
###################################################################################################
def _login2(args):

# This function login into CloudEdnure API
# 
# Usage: _login(args) 	
# 
# Returns: 	None
	global ENDPOINT
	
	session = requests.Session()
	session.headers.update({'Content-type': 'application/json', 'Accept': 'text/plain'})
	
	log.info('Logging in...')
	resp = session.post(url=HOST+ENDPOINT.format('login'), 
						data=json.dumps({'username': args.user, 'password': args.password}))
	if resp.status_code != 200 and resp.status_code != 307:
		log.error('Could not login!')
		print('Could not login!')
		sys.exit(2)
		
	# check if need to use a different API entry point
	if resp.history:
		log.info('URL Redirected...')
		ENDPOINT = '/' + '/'.join(resp.url.split('/')[3:-1]) + '/{}'
		resp = session.post(url=HOST+ENDPOINT.format('login'),
						data=json.dumps({'username': args.user, 'password': args.password}))

	if session.cookies.get('XSRF-TOKEN'):
		session.headers['X-XSRF-TOKEN'] = session.cookies.get('XSRF-TOKEN')	
		
	return session

###################################################################################################
def _get_cloud_ids2(session):

# This function login into CloudEdnure API
# 
# Usage: _get_cloud_ids(session) 	
# 
# Returns: 	a dictionary with CloudId and keys.

	log.info('Getting cloud ids...')
	clouds_resp = session.get(url=HOST+ENDPOINT.format('clouds'))	
	clouds = json.loads(clouds_resp.content)['items']
	
	cloud_ids = {}
	for cloud in clouds:
		if cloud['name'] == 'GCP':
			cloud_ids[cloud['id']] = GCP_KEYS2
		elif cloud['name'] == 'AWS':
			cloud_ids[cloud['id']] = AWS_KEYS2
		elif cloud['name'] == 'AZURE_ARM':
			cloud_ids[cloud['id']] = AZURE_KEYS2
	
	return cloud_ids
	
###################################################################################################	
def _process_bp(bp, session, new_blueprints, failed, project, keys):

# This function process the new bluepint and patch the changes
# 
# Usage: _process_bp(bp, session, new_blueprints) 	
# 
# Returns: 	None.
	project_id = project['id']
	new_bp = True
	try:
		log.info('Processing blueprint for Machine ' + bp['machineName'])

		new_bps = [x for x in new_blueprints if x['machineName'] == bp['machineName'] and project['name'] == x['projectName']]
		if len(new_bps) == 0:
			return
		new_bp = new_bps[0]

		# Sync
		changes_made = False
		for k in new_bp.keys():
			if (k in ['id', 'machineId', 'project','region']):
				continue
				
			if k not in keys:
				continue

			new_value = new_bp[k]
			if new_value.startswith('[') or new_value.startswith('{') or new_value == 'True' or new_value == 'False':
				new_value = eval(new_bp[k]) # NOTE: FOR SECURITY REASONS, DO *NOT* USE 'EVAL' IN PRODUCTION!!

			if new_value != '' and bp[k] != new_value:
				log.info("-- Key {}: Changed '{}' --> '{}'".format(k, bp[k], new_value))
				if new_value == '\"\"':
					new_value = ''
				bp[k] = new_value
				changes_made = True

		if changes_made:
			log.info('Updating blueprint for ' + bp['machineName'])
			bp.pop('machineName', None) # TODO: Need to add support to change Machine name in GCP
			resp = session.patch(url=HOST+ENDPOINT.format('projects/{}/blueprints/{}'.format(project_id, bp['id'])),
								data=json.dumps(bp))
			if resp.status_code != 200:
				log.error('Error setting blueprint for machine due to invalid parameters')
				log.error(resp.status_code)
				log.error(resp.reason)
				log.error(resp.content)
				failed.append(new_bp)
		else:
			log.info('No change was made\n')
	
	except Exception as ex:
		print("failed")
		log.error('Unexpected error occured')
		log.error(ex.message+'\n')
		failed.append(new_bp)
		pass




###################################################################################################


###################################################################################################
def main(args):
	print(chr(27) + "[2J")
	print("CloudEndure AutomaTAG")
	while (args.user is None) or (args.user == ''):
		args.user = input("email address:")
	while (args.password is None) or (args.password == ''):
		args.password = getpass.getpass()
	while (args.task is None) or (args.user == '') or ((args.task != 'add') and (args.task != 'del') and (args.task != 'dryrun')):
		args.task = input("Enter Task <add> or <del> or <dryrun> :")
	while (args.wave is None) or (args.wave == ''):
		args.wave = input("Enter wave. Ex. R02W03:")
	


# This main function gets all the servers and peoject in an account and create a csv file to be 
# used as a template input file to update the blueprints
# 
# Usage: main(args)

	session = _login(args)
	
	clouds = _get_cloud_ids(session)

	projects_resp = session.get(url=HOST+ENDPOINT.format('projects'))	
	if projects_resp.status_code != 200:
		print ('Failed to fetch the project')
		sys.exit(2)
	projects = json.loads(projects_resp.content)['items']
	machines=[]
	for project in projects:
		print ('Processing project ' + project['name'])
		project_id = project['id']
		targetCloud=''
		keys =()
		if project['targetCloudId'] in clouds:
			keys=clouds[project['targetCloudId']][1]
			targetCloud=clouds[project['targetCloudId']][0]
		else:			
			print ('Could not identify target cloud for Project ' + project['name'] + '. Target cloud id is ' + project['targetCloudId'])
			continue
		
		r = session.get(url=HOST+ENDPOINT.format('projects/{}/machines').format(project_id))
		if r.status_code != 200:
			print ('Failed to fetch the machines')
			continue

		
		for machine in json.loads(r.text)['items']:
			machines.append((project['name'], targetCloud, machine['sourceProperties']['name'])+keys)
	_dump_csv(TITLES+machines,open('myMachinesFromCloudendure.csv', mode='w'))
	print ('Done!')

	###### ExcelToCE start ########

	with open('config.json') as json_file:
		data = json.load(json_file)
		for param in data['config']:
			wb_machine_names = xlrd.open_workbook(param['Excel_File_Name'])
			sheet = wb_machine_names.sheet_by_name(param['Excel_Tab_Name'])
			server_col_string = param['Server_Column_Name']
			wave_col_string = param['Wave_Column_Name']
			row_with_field_names = param['Row_With_Field_Names'] - 1
			first_tag = param['First_Tag_Name']
			last_tag = param['Last_Tag_Name']

    #1
	csv_machines = put_machine_names_from_csvfiles_in_array()
    #2
	servers_col = get_server_location(sheet, server_col_string, row_with_field_names)
    #3
	waves_col = get_wave_location(sheet, wave_col_string, row_with_field_names)
    #4
	excel_servers = put_server_names_from_excelfile_in_array(sheet, servers_col, waves_col, args.wave, row_with_field_names)
    #5
	machines = compare_arrays_of_machine_names(csv_machines, excel_servers)
    #6
	create_csv(machines, sheet, args.task, servers_col, first_tag, last_tag, row_with_field_names)

###### excelToCE end ########
	session = _login2(args)
	
	log.info('Reading configuration from CSV file...')
	new_blueprints = _read_blueprints_csv("CE-blueprint.csv")
	
	clouds = _get_cloud_ids2(session)
	
	log.info('Reading data...')
	projects_resp = session.get(url=HOST+ENDPOINT.format('projects'))
	
	if projects_resp.status_code != 200:
		log.error('Failed to fetch projects. Aborting!')
		return -1	
		
	projects = json.loads(projects_resp.content)['items']
	
	datestringNow = datetime.strftime(datetime.now(), '%m-%d-%Y-%H-%M-%S')
	failed = []
	for project in projects:
		log.info('=' * 100)
		log.info('Processing blueprint for Project ' + project['name'])
		project_id = project['id']
		
		keys =[]
		if project['targetCloudId'] in clouds:
			keys=clouds[project['targetCloudId']]
		else:			
			log.warn('Could not identify target cloud for Project ' + project['name'] + '. Target cloud id is ' + project['targetCloudId'])
			continue
			
		machines_resp = session.get(url=HOST+ENDPOINT.format('projects/{}/machines'.format(project_id)))
		if machines_resp.status_code != 200:
			log.error('Failed to fetch machines for project ' + project['name'] + '. Skipping this project')
			continue
		machines = json.loads(machines_resp.content)['items']

		blueprints_resp = session.get(url=HOST+ENDPOINT.format('projects/{}/blueprints'.format(project_id)))
		if blueprints_resp.status_code != 200:
			log.error('Failed to fetch bluprints for project ' + project['name'] + '. Skipping this project')
			continue
		blueprints = json.loads(blueprints_resp.content)['items']

		for bp in blueprints:
			machine_blueprints = None
			machine_blueprints = [m['sourceProperties']['name'] for m in machines if m['id']==bp['machineId']]
			if not machine_blueprints:
				continue
			bp['machineName'] = machine_blueprints[0]

		if args.outputfile and blueprints:
			log.info('Writing '+ project['name'] +' blueprints to CSV file...')
			_write_blueprints_csv(project['name']+'_'+datestringNow+'_'+args.outputfile, blueprints)

		for bp in blueprints:
			if 'machineName' in bp.keys():
				_process_bp(bp, session, new_blueprints, failed, project ,keys)
	
	if failed:
		print("failed1")
		log.info('Not all Blueprints were set. Please fix parameters in FailedBlueprints csv file and re-run')
		_write_blueprints_csv('FailedBlueprints_'+datestringNow+'.csv', failed)
	else:
		log.info('All blueprints were set!')
		### delete files ###
		os.remove("CE-blueprint.csv")
		os.remove("myMachinesFromCloudendure.csv")
		print("All blueprints were set!")
	return 0

	

	


###################################################################################################
if __name__ == '__main__':
	
    parser = argparse.ArgumentParser()
    parser.add_argument('-u', '--user', required=False, help='User name')
    parser.add_argument('-p', '--password', required=False, help='Password')
    ##### excelToCE start ######
    parser.add_argument('-t', '--task', required=False, help='<add>, <del> or <dryrun>')
    parser.add_argument('-w', '--wave', required=False, help='example R0xWxx, Pilot or <all> for all server on spreadsheet')
    ###### excelToCE end ########
    parser.add_argument('-o', '--outputfile', required=False, help='Output CSV file for backup before change')
    main(args = parser.parse_args())