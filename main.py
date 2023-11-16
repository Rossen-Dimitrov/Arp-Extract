import re
import openpyxl
from ports_list import *
from openpyxl import Workbook, load_workbook

from extract_port_fn import extract_ports_for_fw


def extract_ip_arp_for_each_fw():
    extract_ports_for_fw()
    pattern = r'\s+'

    wb = Workbook()

    with open('Palo Arp.txt') as arp_file:
        arp = arp_file.readlines()
        headings = ['Interface', 'IP', 'MAC ADDRESS']

        for line in arp:
            row = re.split(pattern, line)
            if row[0][:4] == 'ae1.':
                data = [row[0], row[1], row[2]]

                if row[0] in internet_transfer:
                    if not 'internet_transfer' in wb.sheetnames:
                        ws = wb.create_sheet('internet_transfer')
                        ws.append(headings)
                    ws = wb["internet_transfer"]
                    ws.append(data)

                elif row[0] in infra:
                    if not 'infra' in wb.sheetnames:
                        ws = wb.create_sheet('infra')
                        ws.append(headings)
                    ws = wb["infra"]
                    ws.append(data)

                elif row[0] in msvx_nsx_t_dmz:
                    if not 'msvx_nsx_t_dmz' in wb.sheetnames:
                        ws = wb.create_sheet('msvx_nsx_t_dmz')
                        ws.append(headings)
                    ws = wb["msvx_nsx_t_dmz"]
                    ws.append(data)

                elif row[0] in msvx_nsx_t_default:
                    if not 'msvx_nsx_t_default' in wb.sheetnames:
                        ws = wb.create_sheet('msvx_nsx_t_default')
                        ws.append(headings)
                    ws = wb["msvx_nsx_t_default"]
                    ws.append(data)

                elif row[0] in sap01:
                    if not 'sap01' in wb.sheetnames:
                        ws = wb.create_sheet('sap01')
                        ws.append(headings)
                    ws = wb["sap01"]
                    ws.append(data)

                elif row[0] in deedvbasf005:
                    if not 'deedvbasf005' in wb.sheetnames:
                        ws = wb.create_sheet('deedvbasf005')
                        ws.append(headings)
                    ws = wb["deedvbasf005"]
                    ws.append(data)

                elif row[0] in vmpchbi01:
                    if not 'vmpchbi01' in wb.sheetnames:
                        ws = wb.create_sheet('vmpchbi01')
                        ws.append(headings)
                    ws = wb["vmpchbi01"]
                    ws.append(data)

                elif row[0] in vmpcsdcn01:
                    if not 'vmpcsdcn01' in wb.sheetnames:
                        ws = wb.create_sheet('vmpcsdcn01')
                        ws.append(headings)
                    ws = wb["vmpcsdcn01"]
                    ws.append(data)

    wb.save('Firewall ARP Table.xlsx')


extract_ip_arp_for_each_fw()
