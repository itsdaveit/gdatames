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
import re
from erpnext.accounts.party import set_taxes as party_st



class Abrechungen(Document):
	@frappe.whitelist()
	def start_processing_zip(self):
		attached_file = frappe.get_all('File', {"attached_to_doctype": "Abrechungen",
												"attached_to_name": self.name})
		gdata_zipfile = frappe.utils.file_manager.get_file_path(attached_file[0].name)
		zip_xml_content =  self.extract_xml_from_zip(gdata_zipfile)
		mes_xml = zip_xml_content[0]
		self.xml_data = mes_xml
		xml_et = ET.fromstring(self.xml_data)
		date = re.search('(?<=files/)(\d\d)_(\d\d\d\d)_mes_usage_export.zip', gdata_zipfile)
		invoice_month = date.group(1)
		invoice_year = date.group(2)
		self.month = invoice_month + "." + invoice_year
		reportentrys = xml_et.findall('ReportEntry')
		#Momentan ist die Verarbeitung auf einen Report pro XML Datei begrenzt.
		#if self.check_report_date():
		if self.check_report_date():
			if len(reportentrys) == 1:
				#if self.xml_data != '':
				#	frappe.utils.file_manager.remove_file(attached_file[0].name)
				log = 'Beginne mit der Verarbeitung der Reports\n'
				for reportentry in reportentrys:
					counter_MaxActiveClients = 0
					log = log + 'Report für Firma: ' + reportentry.attrib['Company'] + '\n'
					log = log + 'Abrechnungsmonat: ' + invoice_month + '.' + invoice_year + '\n'
					log = log + 'G Data Kundenummer: ' + reportentry.attrib['GDCustomerNr'] + '\n'
					log = log + 'Login Name: ' + reportentry.attrib['Login'] + '\n'
					log = log + 'Produkt: ' + reportentry.attrib['Product'] + '\n'
					log = log + 'Gesamtanzahl reporterter Clients: ' + reportentry.attrib['MaxActiveClients'] + '\n'
					managementservers = reportentry.findall('ManagementServer')
					for managementserver in managementservers:
						mes_id_upper = managementserver.attrib['id'].upper()
						log = log + 'Server mit ID: ' + managementserver.attrib['id'] + ' mit ' + managementserver.attrib['MaxActiveClients'] + ' aktiven Clients gefunden.\n'
						counter_MaxActiveClients = counter_MaxActiveClients + int(managementserver.attrib['MaxActiveClients'])
						#Wenn Server keine MaxActiveClients reporten, aber dennoch auftauchen, sind andere Produkte im Spiel
						#Hierfür haben wir noch keine Behandlung
						if managementserver.attrib['MaxActiveClients'] == 0 or managementserver.attrib['MaxActiveClients'] == "0":
							msgtext = "Für Server " + managementserver.attrib['id'] + " wurden keine Clients, aber andere Produkte reported. Bitte prüfen"
							log += msgtext +"\n"
							frappe.msgprint(msgtext)
						else:
							self.create_mes_invoice(mes_id_upper, managementserver.attrib['MaxActiveClients'], invoice_month, invoice_year )
					log = log + 'Gesamtanzahl Clients gezählt: ' + str(counter_MaxActiveClients)
					self.log = log
					self.status = "Ausgangsrechnungen erstellt"
					self.anzahl_clients = str(counter_MaxActiveClients)
					self.save()
			else:
				frappe.msgprint('Keinen oder mehr als ein ReportEntry im XML-Code gefunden, breche ab.')


	def create_mes_invoice(self, mes_id, max_active_clients, invoice_month, invoice_year):
		if max_active_clients == 0:
			frappe.throw("Managementserver ")
		GDATAMES_Settings = frappe.get_doc("GDATAMES Settings")
		management_server = frappe.get_all('Management Server', {"management_server_id": mes_id})
		if len(management_server) == 0:
			frappe.throw('Management Server ID ' + mes_id + ' nicht gefunden.')
		if len(management_server) == 1:
			doc_management_server = frappe.get_doc("Management Server", management_server[0].name)
			product = frappe.get_doc("Produkte", doc_management_server.product)
			item = frappe.get_doc("Item", product.item)
			sales_invoice_item = frappe.get_doc({"doctype": "Sales Invoice Item",
									"item_code": item.name,
									"qty": int(max_active_clients),
									})

			introduction_text = GDATAMES_Settings.introduction_text + "<div><br></div><div>Leistungszeitraum " + invoice_month + "." + invoice_year + "<br>Ihre Management Server ID: " + mes_id + "</div>"

			sales_invoice_doc = frappe.get_doc({"doctype": "Sales Invoice",
									"title": "MES " + invoice_month + "." + invoice_year + " " + frappe.get_doc("Customer", doc_management_server.customer).customer_name,
									"customer": doc_management_server.customer,
									"status": "Draft",
									#"payment_terms_template": GDATAMES_Settings.payment_terms_template,
									"tc_name": GDATAMES_Settings.terms_and_conditions,
									"company": frappe.get_doc("Global Defaults").default_company,
									"introduction_text": introduction_text
									})
			customer_doc = frappe.get_doc("Customer", doc_management_server.customer )
			#print(customer_doc.payment_terms)
			if customer_doc.payment_terms:
				sales_invoice_doc.payment_terms_template = customer_doc.payment_terms
			else:
				sales_invoice_doc.payment_terms_template = GDATAMES_Settings.payment_terms_template
			sales_invoice_doc.append("items", sales_invoice_item)
			SINV = frappe.get_doc("Sales Invoice", sales_invoice_doc.insert().name)

			# für Kunden gültiges Sales Taxes and Charges Template mit ERPNext eigener Funktion ermitteln
			# Berücksichtigt die Anwendung von steuerreglen
			SINV.taxes_and_charges = party_st(SINV.customer, "Customer", SINV.posting_date, SINV.company)
			#entsprechend als Doctype laden
			taxes = frappe.get_doc("Sales Taxes and Charges Template", SINV.taxes_and_charges).taxes
			#im Doctype hinterlegte steuern müssen der Rechnung angefügt werden.
			#Die Berechnung der Beträge geschieht automatisch
			for tax in taxes:
				new_tax = frappe.get_doc({
					"doctype": "Sales Taxes and Charges",
					"charge_type": tax.charge_type,
					"account_head": tax.account_head,
					"rate": tax.rate,
					"description": tax.description
				})
				SINV.append("taxes", new_tax)
			
			tac_doc = frappe.get_doc("Terms and Conditions", GDATAMES_Settings.terms_and_conditions)
			SINV.terms = terms = tac_doc.terms
			SINV.save()
		else:
			frappe.throw('Management Server ID ' + mes_id + ' nich einmalig.')


	def extract_xml_from_zip(self, inputzip):
		zf = zipfile.ZipFile(inputzip, 'r')
		files = []
		for name in zf.namelist():
			if fnmatch.fnmatch(name, '*.xml'):
				files.append(zf.read(name))
		return files
	
	def check_report_date(self):
		abrechnung_list =  frappe.get_all("Abrechungen", filters={"month": self.month })
		if len(abrechnung_list) == 0:
			return True
		else:

			self.status = "fehlerhaft"
			self.log ="Zu dem Abrechnungsmonat "+ self.month +" existiert bereits eine Abrechnung ("+abrechnung_list[0]["name"]+"), es wurde keine neue Abrechnung erstellt."     
            
			return False