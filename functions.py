from ports_list import *
import re
import ipaddress

def extract_ports_for_fw():
    pattern = r'\s+'
    with open('PaloPortsByVfw.txt') as port_file:
        ports = port_file.readlines()
        for line in ports:
            row = re.split(pattern, line)
            if row[0][:4] == 'ae1.':
                if row[4] == 'vr:INTERNET-TRANSFER':
                    internet_transfer.append(row[0])
                elif row[4] == 'vr:Infra':
                    infra.append(row[0])
                elif row[4] == 'vr:MSVX-NSX-T-DMZ':
                    msvx_nsx_t_dmz.append(row[0])
                elif row[4] == 'vr:MSVX-NSX-T-Default':
                    msvx_nsx_t_default.append(row[0])
                elif row[4] == 'vr:Sap01':
                    sap01.append(row[0])
                elif row[4] == 'vr:deedvbasf005':
                    deedvbasf005.append(row[0])
                elif row[4] == 'vr:vmpchbi01':
                    vmpchbi01.append(row[0])
                elif row[4] == 'vr:vmpcsdcn01':
                    vmpcsdcn01.append(row[0])


def create_new_worksheet(name, wb) -> None:
    headings = ['Interface (including VLAN)', 'IP-Adress', 'MAC-Adress']
    ws = wb.create_sheet(name)
    ws.append(headings)


def append_data_to_worksheet(data, worksheet_name, wb) -> None:
    ws = wb[worksheet_name]
    ws.append(data)


def get_lb_name(first_row) -> str:
    row = str(first_row)
    return row.split('@(')[1].split(')')[0]


def validate_ip_address(ip):
    try:
        ipaddress.ip_address(ip)
        return True

    except ValueError:
        return False
