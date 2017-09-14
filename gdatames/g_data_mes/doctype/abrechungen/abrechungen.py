# -*- coding: utf-8 -*-
# Copyright (c) 2017, itsdave GmbH and contributors
# For license information, please see license.txt

from __future__ import unicode_literals
import frappe
from frappe.model.document import Document
import zipfile
from frappe.utils import file_manager
import fnmatch
import xml.etree.ElementTree as ET


class Abrechungen(Document):
	def start_processing_zip(self):
		attached_file = frappe.get_all('File', {"attached_to_doctype": "Abrechungen",
												"attached_to_name": self.name})
		gdata_zipfile = frappe.utils.file_manager.get_file_path(attached_file[0].name)
		zip_xml_content =  self.extract_xml_from_zip(gdata_zipfile)
		mes_xml = zip_xml_content[0]
		self.xml_data = mes_xml
		xml_et = ET.fromstring(self.xml_data)
		reportentrys = xml_et.findall('ReportEntry')
		#Momentan ist die Verarbeitung auf einen Report pro XML Datei begrenzt.
		if len(reportentrys) == 1:
			log = 'Beginne mit der Verarbeitung der Reports\n'
			for reportentry in reportentrys:
				counter_MaxActiveClients = 0
				log = log + 'Report für Firma: ' + reportentry.attrib['Company'] + '\n'
				log = log + 'G Data Kundenummer: ' + reportentry.attrib['GDCustomerNr'] + '\n'
				log = log + 'Login Name: ' + reportentry.attrib['Login'] + '\n'
				log = log + 'Produkt: ' + reportentry.attrib['Product'] + '\n'
				log = log + 'Gesamtanzahl reporterter Clients: ' + reportentry.attrib['MaxActiveClients'] + '\n'
				managementservers = reportentry.findall('ManagementServer')
				for managementserver in managementservers:
					log = log + 'Server mit ID: ' + managementserver.attrib['id'] + ' mit ' + managementserver.attrib['MaxActiveClients'] + ' aktiven Clients gefunden.\n'
					counter_MaxActiveClients = counter_MaxActiveClients + int(managementserver.attrib['MaxActiveClients'])
				log = log + 'Gesamtanzahl Clients gezählt: ' + str(counter_MaxActiveClients)
				self.log = log


		else:
			msgprint('Keinen oder mehr als ein ReportEntry im XML-Code gefunden, breche ab.')




	def extract_xml_from_zip(self, inputzip):
		zf = zipfile.ZipFile(inputzip, 'r')
		files = []
		for name in zf.namelist():
			if fnmatch.fnmatch(name, '*.xml'):
				files.append(zf.read(name))
		return files
