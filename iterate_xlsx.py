from openpyxl import load_workbook
import os
import re

join = os.path.join
working_dir = os.path.normpath("C:/Users/SQA-AGrudev/Desktop/TAF_WORKING_DIR")
working_dir_json = join(working_dir, 'Jsons')
working_dir_xlsx = join(working_dir, 'XLSX')
working_dir_sorted = join(working_dir, 'Sorted')


def rename_files(name_to_be_ranamed):
    """
    Renames files in the current working directory directory to match the current convention.
    """
    os.chdir(working_dir)
    working_name = str(name_to_be_ranamed)
    if re.search("PostPaid_CON", working_name):
        working_name = re.sub('PostPaid_CON', '', working_name)
        working_name = re.sub('_', "_PostPaid_CON_", working_name, 1)
    if re.search("PostPaid_BUS", working_name):
        working_name = re.sub('PostPaid_BUS', '', working_name)
        working_name = re.sub('_', "_PostPaid_BUS_", working_name, 1)
    if re.search("PostPaid_SP", working_name):
        working_name = re.sub('PostPaid_SP', '', working_name)
        working_name = re.sub('_', "_PostPaid_SP_", working_name, 1)
    if re.search("_LTE", working_name):
        working_name = re.sub('_LTE', '', working_name)
        working_name = re.sub('_', "_LTE_", working_name, 1)
    working_name = re.sub('Terminate', 'Delete', working_name)
    working_name = re.sub('Register', 'Create', working_name)
    working_name = re.sub('04947', '', working_name)
    working_name = re.sub('04451', '', working_name)
    working_name = re.sub('04452', '', working_name)
    working_name = re.sub('04166', '', working_name)
    working_name = re.sub('04946', '', working_name)
    working_name = re.sub('04334', '', working_name)
    working_name = re.sub('04460', '', working_name)
    working_name = re.sub('__', '_', working_name)
    working_name = re.sub('_\s', '_', working_name)
    working_name = re.sub('_\.', '.', working_name)
    working_name = re.sub('Change_IMSI', 'ChangeIMSI', working_name)
    print(working_name)
    os.rename(name_to_be_ranamed, working_name)


def operate_cwd_old(work_on):
    """
    Operates over the files in the given working directory comment and uncomment needed functions
    returns a list of file names for later sorting
    Also returns a list of files later use for numbering them
    """
    # JSON
    if work_on == 'json':
        for root, dirs, files in os.walk(working_dir_json, topdown=False):
            if root != working_dir_json:
                continue
            for name in files:
                if name[-4:] == 'json':
                    pass
                    remove_request_trace(name)
                    get_service_names(name)
                    replace_service_names(name)
                    parametrize_payload(name)
    if work_on == 'json':
        for root, dirs, files in os.walk(working_dir_json, topdown=False):
            if root != working_dir_json:
                continue
            for name in files:
                if name[-4:] == 'json':
                    replace_var_value(False, name)
    # XLS
    if work_on == 'xlsx':
        for root, dirs, files in os.walk(working_dir_xlsx, topdown=False):
            if root != working_dir_xlsx:
                continue
            for name in files:
                if name[-4:] == 'xlsx':
                    pass
                    workbook = load_workbook(filename=join(root, name))
                    ws = load_sheet(workbook)
                    if not ws:
                        print(f'Failed to load sheet for file {name}')
                        break
                    remove_resp_add_first_four_lines(ws)
                    update_names(ws, name[:-4] + 'json')
                    remove_sri(ws)
                    remove_antics(ws)
                    replace_var_value(ws, name)
                    workbook.save(filename=join(root, name))
                print(f"Successfully modified File:{name}")


def insert_in_nvar(size_to_match, string_to_insert, modified_list):
    """
    Fills the modified list with as many values as the size of the first list
    """
    for _ in size_to_match:
        modified_list = modified_list + [string_to_insert]
    return modified_list


def replace_var_value(ws, name_working):
    """
    Parametrazis the excels and json files.
    """
    ICCID = ['89470311210210000681', "89470311210210000699", "89470311210210000707", "89470311210210000715",
             "89470311210210000723", "89470311210210000731", "89470311210210000749", "89470311210210000756",
             "89470311210210000764", "89470311210210000772", "89470311210210000780", "89470311210210000798",
             "89470311210210000806", "89470311210210000814", "89470311210210000822", "89470311210210000830",
             "89470311210210000996", "89470000190625003999", "89470000190625003981", "89470000190625003973",
             "89470000190625003965", "89470000190625003957", "89470000190625003940", "89470000190625003932",
             "89470000190625003924", "89470000190625003890", "89470000190625003882", "89470000190625003874",
             "89470000190625003866", "89470000190625003825", "89470000190625003817", "89470000190625003809",
             "89470000190625003791", "89470311210210000962", "89470311210210000970", "89470311210210000988"]
    IMSI = ["242013083407751", "242013083407752", "242013083407753", "242013083407754", "242013083407758",
            "242013083407759", "242013083407760", "242013083407771", "242013083407764", "242013083407765",
            "242013083407766", "242013083407767", "242013083407768", "242013083407769", "242013083407770",
            "242013083407771", "242013083407762", "242013083407763", "242013083407757", "242013083407756",
            "242013083407755", "242013112998926", "242013112998927", "242013112998928", "242013112998929",
            "242013112998930", "242013112998931", "242013112998932", "242013112998933", "242013112998934",
            "242013112998935", "242013112998936", "242013112998937", "242013112998938", "242013112998939",
            "242013112998940", "242013112998941", "242013112998942", "242013112998943", "242013112998944",
            "242013112998945", "242013112998946", "242013112998947", "242013112998948", "242013112998949",
            "242013112998962", "242013112998963", "242013112998964", "242013112998965"]
    order_subsceriber_desc = ["580002010006", "580002010007", "580002010008", "580002010009", "580002010010",
                              "580002010011", "580002010012", "580002010013", "580002010005", "580002010014",
                              "580002010015", "580002010016", "580002010017", "580002010018", "580002010019",
                              "580002010020", "580004217291", "580004217292", "580004217293", "580004217294",
                              "580004217295", "580004217296", "580004217297", "580004217298", "580004217299",
                              "580000001404", "580002150328", "580002150990", "580002150964", "580002150963",
                              "580002150960", "580002150712", "580002150388", "580002150328", "580002150304",
                              "580002150290", "580002150284", "99899862", "99899824", "99899548", "99899499",
                              "99899444", "99898929", "99898689", "99897498", "99897152", "99896844", "580000084347",
                              "580000051184", "580000041302", "580000041274", "580000041270", "580000015006",
                              "580000010850", "580000001404", "580000000972", "580000000304", "99792775", "99749252",
                              "99748432", "99741336", "99723770", "99721033", "99640312", "99628853", "99617374",
                              "99547073", ]
    order_subscriber_id = ["41147337", "41186302", "41147373", "41135924", "41171632", "41171641", "41171646",
                           "41171647", "41135929", "51547639", "34572037", "34944493", "34174958", "37210572",
                           "43927785", "35019217", "48505103", "38865130", "34723909", "51593973", "46308554",
                           "44838652", "50569389", "44906460", "35178340", "43997935", "41559450", '20929630',
                           '47989748', '48270620', '47988315', '47989812', '48271069', '47987444', '47990536',
                           '47990098', '48270844', '51751826', '51751842', '51751847', '48270280', '29410992',
                           '48207456', '48207305', '48207327', '48537495', '47485638', '51709328', '40737621',
                           '45588396', '50335525', '41713772', '50177158', '46447744', '48643763', '46705964',
                           '45927525', '42647761', '42651740', '42653822', '42648906', '42652680', '42652015',
                           '42994967', '42995459', '42995945', '42995341']
    service_name = []
    with open('service_names.txt', 'r') as file:
        for line in file:
            line = line.strip('\n')
            line = f'"{line}"'
            service_name.append(line)
    system_component_id = ['pcrf_1384193797', "pcrf_1084193793", "RoamingZone", ]
    valid_to_date = ['05-09-2022']
    roaming_zone = ['RZ1']

    o_var = order_subsceriber_desc + order_subscriber_id + service_name + system_component_id + valid_to_date + roaming_zone + ICCID + IMSI
    n_var = []

    n_var = insert_in_nvar(order_subsceriber_desc, '${ORDER_SUBSCRIBER_DESC}', n_var)
    n_var = insert_in_nvar(order_subscriber_id, '${ORDER_SUBSCRIBER_ID}', n_var)
    n_var = insert_in_nvar(service_name, '${SERVICE_NAME}', n_var)
    n_var = insert_in_nvar(system_component_id, '${SYSTEM_COMPONENT_ID}', n_var)
    n_var = insert_in_nvar(valid_to_date, "${VALID_TO_DATE}", n_var)
    n_var = insert_in_nvar(roaming_zone, "${ROAMING_ZONE}", n_var)
    n_var = insert_in_nvar(ICCID, "${ICCID}", n_var)
    n_var = insert_in_nvar(IMSI, "${IMSI}", n_var)
    if name_working[-4:] == 'json':
        replace_in_json(name_working, o_var, n_var)
    elif name_working[-4:] == 'xlsx':
        replace_in_xls(ws, o_var, n_var)


def replace_in_json(name_working, o_var, n_var):
    """
    Parametrize the JSON file
    """
    with open(join(working_dir_json, name_working), "r") as file:
        f = file.read()
        for idx, _ in enumerate(o_var):
            f = re.sub(o_var[idx], n_var[idx], f)
    with open(join(working_dir_json, name_working), "w") as file:
        file.write(f)


def replace_in_xls(ws, old_var, new_var):
    """
    Parametrize the JSON file
    """
    if not ws:
        return
    collum = 'J'
    for idx, _ in enumerate(old_var):
        i = 8
        while ws[collum + str(i)].value is not None:
            ws[collum + str(i)].value = re.sub(old_var[idx], new_var[idx], ws[collum + str(i)].value)
            if ws[collum + str(i)].value == old_var[idx]:
                ws[collum + str(i)] = new_var[idx]
            i += 1


def remove_sri(ws):
    """
    Removes the SRI tasks.
    """
    collum_key = 'I'
    collum_value = 'J'
    collum_task = 'K'
    i = 2
    f = 0
    while ws[collum_value + str(i)].value is not None:
        if (ws[collum_value + str(i)].value == 'SRI') and (ws[collum_key + str(i)].value == 'NE_TYPE'):
            task_id = ws[collum_task + str(i)].value
            while ws[collum_task + str(i)].value == task_id:
                ws.delete_rows(i, 1)
                f = 1
            if f == 1:
                i -= 1
        i += 1


def remove_resp_add_first_four_lines(ws):
    """
    Removes the response parameters and add additional 4 lines to each excel file.
    """
    if ws["I4"].value != "TimeOut":
        ws.insert_rows(4)
        ws.insert_rows(4)
        ws.insert_rows(4)
        ws.insert_rows(4)

        ws['D' + str(4)].value = 'IL'
        ws['E' + str(4)].value = 'FlowoneAPI'
        ws['F' + str(4)].value = 'Setup'
        ws['G' + str(4)].value = 'AsyncReceiver'
        ws['I' + str(4)].value = 'TimeOut'
        ws['J' + str(4)].value = '${TimeOut}'

        ws['D' + str(5)].value = 'IL'
        ws['E' + str(5)].value = 'FlowoneAPI'
        ws['F' + str(5)].value = 'Verify'
        ws['G' + str(5)].value = 'AckResponse'
        ws['I' + str(5)].value = 'statusMessage'
        ws['J' + str(5)].value = 'InstantLink accepted request with request id: ${Request_id} for order no: 1234567890'

        ws['D' + str(6)].value = 'IL'
        ws['E' + str(6)].value = 'FlowoneAPI'
        ws['F' + str(6)].value = 'Verify'
        ws['G' + str(6)].value = 'Response'
        ws['H' + str(6)].value = 'serviceOrder'
        ws['I' + str(6)].value = 'statusMessage'
        ws['J' + str(6)].value = 'Order delivered'
        ws['K' + str(6)].value = 'Resp'

        ws['D' + str(7)].value = 'IL'
        ws['E' + str(7)].value = 'FlowoneAPI'
        ws['F' + str(7)].value = 'Verify'
        ws['G' + str(7)].value = 'Response'
        ws['H' + str(7)].value = 'serviceOrder'
        ws['I' + str(7)].value = 'state'
        ws['J' + str(7)].value = 'completed'
        ws['K' + str(7)].value = 'Resp'

    collum_value = 'H'
    i = 8
    while ws['G' + str(i)].value is not None:
        if ws[collum_value + str(i)].value == 'TaskResponse':
            ws.delete_rows(i, 1)
        if ws['G' + str(i)].value == 'total_ne_tasks':
            ws.delete_rows(i, 1)
        i += 1


def update_names(ws, name_to_be_updated):
    """
    Upadate the test name in the excel to match the name of the excel file.
    Also gives the same name to the json that is required as input.
    """
    collum_value = 'I'
    i = 2
    while ws[collum_value + str(i)].value is not None:
        if ws["H" + str(i)].value == "RequestPayload":
            ws["I" + str(i)].value = "PayloadFile"
            ws["J" + str(i)].value = name_to_be_updated
        i += 1
    ws['A3'].value = name_to_be_updated[:-5]
    ws['B3'].value = name_to_be_updated[:-5]


def remove_antics(ws):
    """
    Removes duplicate variables in the excel files
    """
    collum_key = 'I'
    collum_value = 'J'
    collum_task = 'K'
    i = 2
    while ws[collum_value + str(i + 1)].value is not None:
        task_id = ws[collum_task + str(i)].value
        param_value = ws[collum_key + str(i)].value
        j = i + 1
        while task_id == ws[collum_task + str(j)].value:
            if param_value == ws[collum_key + str(j)].value:
                ws.delete_rows(j, 1)
            j += 1
        i += 1


def sort_files(work_on):
    """
    Sorts the payload files into Subscribers
    """
    wd = False
    groups = ["PostPaid_CON", "PostPaid_BUS", "PostPaid_SP", "LTE", "PrePaid_SP", "PrePaid", "M2M"]

    if work_on == 'json':
        wd = working_dir_json
    if work_on == 'xlsx':
        wd = working_dir_xlsx

    for root, dirs, files in os.walk(wd, topdown=False):
        if root != wd:
            continue
        for name in files:
            os.chdir(wd)
            for group in groups:
                if re.search(group, name):
                    os.replace(join(wd, name),
                               join(working_dir_sorted, group, name))


def give_number_to_files(starting_number):
    """
    Renames files in the current working directory directory to be sorted in order and in create delete pairs
    """
    i = starting_number

    for root, dirs, files in os.walk(working_dir_sorted, topdown=False):
        file_names = files
        k = 0
        os.chdir(root)
        while k < len(file_names) and i < len(file_names) + starting_number:
            j = 0
            while j < len(file_names):
                if file_names[k][6:] == file_names[j][6:] and k != j:
                    try:
                        if i < 9:
                            os.rename(file_names[k], '000' + str(i) + file_names[k])
                            os.rename(file_names[j], '000' + str(i + 1) + file_names[j])
                        elif i == 9:
                            os.rename(file_names[k], '000' + str(i) + file_names[k])
                            os.rename(file_names[j], '00' + str(i + 1) + file_names[j])
                        else:
                            os.rename(file_names[k], '00' + str(i) + file_names[k])
                            os.rename(file_names[j], '00' + str(i + 1) + file_names[j])
                        i += 2
                    except OSError as err:
                        print(err.errno)
                    break
                j += 1
            k += 1
        break


def remove_number_from_files():
    """
    Renames files in the current working directory directory to remove the sorting numbers
    """
    os.chdir(working_dir_xlsx)
    file_names = []
    for root, dirs, files in os.walk(working_dir_xlsx, topdown=False):
        file_names = files
        break
    k = 0
    while k < len(file_names):
        if re.match('\d{4}', file_names[k]) is not None:
            os.rename(file_names[k], file_names[k][4:])
        k += 1


def remove_request_trace(name_working):
    """
    Removes the request trace parameters from the json files.
    """
    with open(join(working_dir_json, name_working), "r") as file:
        f = file.read()
        f = re.sub(',\s*{\s*"name":\s"REQUEST_TRACE",\s*"value":\s"EVERYTHING"\s*}', '', f)
    with open(join(working_dir_json, name_working), "w") as file:
        file.write(f)


def replace_service_names(name_working):
    """
    Replaces the Service name value with ${SERVICE_NAME} in all json. Effectivly parametrizing them.
    """

    with open(join(working_dir_json, name_working), "r") as file:
        file_string = file.read()
        service_match = re.search(r'"service":\s{\s*"id":\s"([a-zA-z0-9]*)",\s*"name":\s"([0-9]*)"', file_string)
        if service_match is not None:
            file_string = re.sub('"service":\s{\s*"id":\s(.*)\s*"name":\s"([0-9]*)"',
                                 '"service": {\n        "id":"%s",\n        "name":"${SERVICE_NAME}"' % (
                                     service_match.group(1)), file_string)
    with open(join(working_dir_json, name_working), "w") as file:
        file.write(file_string)


def get_service_names(file_name):
    """
    Reads the service name from the json files and prints them to the screen.
    """
    os.chdir(working_dir_json)
    with open(join(working_dir_json, file_name), "r") as file:
        file_string = file.read()
        service_match = re.search(r'"service":\s{\s*"id":\s"([a-zA-z0-9]*)",\s*"name":\s"([0-9]*)"', file_string)

    if service_match is not None:
        print(f',"{service_match.group(2)}"')
        with open('service_names.txt', 'a') as file:
            file.write(f'{service_match.group(2)}\n')


def parametrize_payload(name_working):
    """
    Parametrizes the payload with the default values
    """
    with open(join(working_dir_json, name_working), "r") as file:
        file_string = file.read()
        file_string = re.sub('"externalId":\s"([^"]*)"', '"externalId": "${externalId}"', file_string)
        file_string = re.sub('"orderDate":\s"([^"]*)"', '"orderDate": "${orderDate}"', file_string)
        file_string = re.sub('"orderType":\s"([^"]*)"', '"orderType": "${orderType}"', file_string)

        match = re.search(r'"AsyncResponse":\s{[\s\S]*"replyToAddress":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${AsyncResponse_replyToAddress}', file_string)

        match = re.search(r'"WSSec":\s{[\s\S]*"username":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${WSSec_username}', file_string)

        match = re.search(r'"WSSec":\s{[\s\S]*"password":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${WSSec_password}', file_string)

        match = re.search(r'"OMRequestSpec":\s{[\s\S]*"neType":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${OMRequestSpec_neType}', file_string)

        match = re.search(r'"name":\s"IL_REQ_GROUP",\s*"value":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${IL_REQ_GROUP}', file_string)

        match = re.search(r'"name":\s"ORDER_SUBSCRIBER_ID",\s*"value":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${ORDER_SUBSCRIBER_ID}', file_string)

        match = re.search(r'"name":\s"ORDER_SUBSCRIBER_DESC",\s*"value":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${ORDER_SUBSCRIBER_DESC}', file_string)

        match = re.search(r'"name":\s"ORDER_SYS_ID",\s*"value":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${ORDER_SYS_ID}', file_string)

        match = re.search(r'"name":\s"CASE_ID",\s*"value":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${CASE_ID}', file_string)

        match = re.search(r'"name":\s"ORDER_USER_ID",\s*"value":\s"([^"]*)"', file_string)
        file_string = re.sub(match.group(1), '${ORDER_USER_ID}', file_string)

    with open(join(working_dir_json, name_working), "w") as file:
        file.write(file_string)


def get_payload(sheet_name, workbook):
    """
    Reads the Dayly report xcel file and creates payload from it. It requres a name collum to be set up at collum C.
    """
    ws = workbook[sheet_name]
    i = 3
    os.chdir(working_dir)
    while ws['F' + str(i)].value is not None:
        match = re.search(
            r'[jJ][sS][oO][nN][\s.]?[Rr][eE][Qq][uUeEsStT.]{0,4}:?\s{0,3}([\S\s]*)[Jj][Ss][Oo][Nn][\s.]?[Rr][Ee][Ss][pPoOnNsSeE.]{0,5}:?',
            ws['F' + str(i)].value)
        test_name = ws['C' + str(i)].value
        if test_name is None:
            print(f'Name missing for line {i}')
        elif match is None:
            print(f'Payload missing for line {i}')
        else:
            with open(test_name + '.json', 'w') as file:
                file.write(match.group(1))
        i += 1


def get_request_id_cfg_file(sheet_name, workbook, open_mode='w'):
    """
    Reads the Dayly report xcel file creates the reqest_ids file for the task generator
    and the cfg file for the excel generator
    """
    ws = workbook[sheet_name]
    i = 3
    os.chdir(working_dir)
    request_ids = []
    test_cfg = []
    while ws['F' + str(i)].value is not None:
        match = re.match('[Rr][Ee][Qq][UuEeSsTt]*.?\s?[Ii][Dd][.:]{0,2}\s{0,4}(\d+)', ws['F' + str(i)].value)
        test_name = ws['C' + str(i)].value
        if test_name is None:
            print(f'Name missing for line {i}')
        elif match is None:
            print(f'Request missing for line {i}')
        else:
            request_ids.append(match.group(1))
            test_cfg.append(f'{test_name}')
        i += 1
    with open('request_ids.txt', open_mode) as file:
        temp_string = ''
        for el in request_ids:
            temp_string += f'{el}\n'
        file.write(temp_string)
    with open('cfg_file.txt', open_mode) as file:
        temp_string = ''
        for el in test_cfg:
            temp_string += f'{el}\n'
        file.write(temp_string)


def load_sheet(workbook):
    """
    Some excels are generated with different sheetnames
    """
    try:
        ws = workbook['Test cases1']
    except:
        try:
            ws = workbook['Test cases']
        except:
            return
    return ws


def operate_cwd(work_on):
    """
    Operates over the files in the given working directory comment and uncomment needed functions
    returns a list of file names for later sorting
    Also returns a list of files later use for numbering them
    """
    # JSON

    # XLS
    for root, dirs, files in os.walk(working_dir_xlsx, topdown=False):
        if root != working_dir_xlsx:
            continue
        for name in files:
            if name[-4:] == 'xlsx' and work_on == "xlsx":
                pass
                workbook = load_workbook(filename=join(root, name))
                ws = load_sheet(workbook)
                if not ws:
                    print(f'Failed to load sheet for file {name}')
                    break
                remove_resp_add_first_four_lines(ws)
                update_names(ws, name[:-4] + 'json')
                remove_sri(ws)
                remove_antics(ws)
                replace_var_value(ws, name)
                workbook.save(filename=join(root, name))
            if name[-4:] == 'json' and work_on == 'json':
                pass
                remove_request_trace(name)
                get_service_names(name)
                replace_service_names(name)
                parametrize_payload(name)

    if work_on == 'json':
        for root, dirs, files in os.walk(working_dir_json, topdown=False):
            if root != working_dir_json:
                continue
            for name in files:
                if name[-4:] == 'json':
                    replace_var_value(False, name)
    # print(f"Successfully modified File:{name}")


def operate_workbook():
    """
    Operates the Test Report file, generates cfg file for TAF regression generator,
    service names files later used for parametraization and request ids file used to get request from DB
    """
    workbook = load_workbook(join(working_dir, 'Test_Report.xlsx'))
    curret_day = 3
    last_day = 13
    while curret_day < last_day:
        get_payload(f'Day {curret_day}', workbook)
        print(f'Payload for Day {curret_day} Finished \n --------------------------------------')
        get_request_id_cfg_file(f'Day {curret_day}', workbook)
        print(f'Request IDS for Day {curret_day} Finished \n --------------------------------------')
        curret_day += 1


def init_dir():
    """
    Creates the necessary files to operated
    """
    if not os.path.isdir(working_dir_json):
        os.mkdir(working_dir_json)
    if not os.path.isdir(working_dir_sorted):
        os.mkdir(working_dir_sorted)
    if not os.path.isdir(working_dir_xlsx):
        os.mkdir(working_dir_xlsx)
    l = ['LTE', 'M2M', "PostPaid_BUS", "PostPaid_CON", "PostPaid_SP", "PrePaid", "PrePaid_SP"]
    for i in l:
        if not os.path.isdir(join(working_dir_sorted, i)):
            os.mkdir(join(working_dir_sorted, i))


# init_dir()
# operate_workbook()

# operate_cwd('json')
# operate_cwd('xlsx')
# sort_files('xlsx')
# give_number_to_files(2)
# sort_files('json')

# remove_number_from_files()
