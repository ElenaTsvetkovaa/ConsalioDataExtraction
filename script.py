import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import os

class InvoiceParser:
    def __init__(self, xml_file, save_folder='templates'):
        self.xml_file = xml_file
        self.save_folder = save_folder
        self.ns = {
        'rsm': "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
        'ram': "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100" ,
        'udt': "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100"
    }

        self.global_invoice_data = []
        self.line_items_data = []
        self.parse_xml()

    def parse_xml(self):
        tree = ET.parse('1300457821.xml')
        root = tree.getroot()


        for invoice in root.findall('rsm:ExchangedDocument', self.ns):
            invoice_number = invoice.find('ram:ID', self.ns).text
            raw_date = datetime.strptime(
                invoice.find('ram:IssueDateTime/udt:DateTimeString', self.ns).text, "%Y%m%d")
            invoice_date = raw_date.strftime("%d.%m.%Y")

            #   Recipient Info
            recipient_path = root.find('.//ram:BuyerTradeParty', self.ns)
            recipient = recipient_path.find('ram:Name', self.ns).text
            recipient_person = recipient_path.find('ram:DefinedTradeContact/ram:PersonName', self.ns).text

            trade_settlement = root.find('.//ram:ApplicableHeaderTradeSettlement', self.ns)
            currency = trade_settlement.find('ram:InvoiceCurrencyCode', self.ns).text
            total_vat_excluded = trade_settlement.find('.//ram:LineTotalAmount', self.ns).text
            total_vat_included = trade_settlement.find('.//ram:GrandTotalAmount', self.ns).text

            self.global_invoice_data.append(
                [recipient, "", invoice_number, "", "", "", "", currency,
                 total_vat_excluded, total_vat_included, invoice_date, "", "", "",
                 recipient_person, "", "", "", "", "", "", "", ""]
            )
            # Line items

        for line_data in root.findall('.//ram:IncludedSupplyChainTradeLineItem', self.ns):

            description = line_data.find('.//ram:SpecifiedTradeProduct/ram:Name', self.ns).text
            unit_price = line_data.find('.//ram:SpecifiedLineTradeAgreement/ram:NetPriceProductTradePrice/ram:ChargeAmount', self.ns).text
            quantity = line_data.find('.//ram:SpecifiedLineTradeDelivery/ram:BilledQuantity', self.ns).text
            line_price = line_data.find('.//ram:SpecifiedTradeSettlementLineMonetarySummation/ram:LineTotalAmount', self.ns).text
            vat = line_data.find('.//ram:SpecifiedLineTradeSettlement/ram:ApplicableTradeTax/ram:RateApplicablePercent', self.ns).text

            self.line_items_data.append(
                ["", "", "", "", description, quantity, unit_price, line_price, vat, "", "", "", "",
                 "", "", "", ""]
            )

    def create_global_to_excel(self):

        global_tem_columns = [
            "Recipient", "Recipient Parent", "Invoice Number", "Invoice Type", "Pricing Type",
            "RVG - Dispute Value", "RVG - Fee Rate", "Currency", "Total VAT Excluded",
            "Total VAT Included", "Invoice Date", "Due Date", "Period Start", "Period End",
            "Recipient Display Name", "Recipient First Name", "Recipient Last Name",
            "Issuer Display Name", "Issuer First Name", "Issuer Last Name", "Reference / PO Number",
            "Is Draft", "Debit Invoice Number"
        ]
        global_temp = pd.DataFrame(self.global_invoice_data, columns=global_tem_columns)
        output_file = os.path.join(self.save_folder, "global.xlsx")

        global_temp.to_excel(output_file, index=False)

    def create_line_to_excel(self):
        line_items_columns = [
            "Date Start", "Date End", "Is Hours", "Code", "Description",
            "Quantity", "Unit Price", "Line Price", "VAT Rate",
            "Cost Type", "Cost Centre", "Project 1", "Project 2", "Project 3",
            "Workstream 1", "Workstream 2", "Workstream 3"
        ]

        line_temp = pd.DataFrame(self.line_items_data, columns=line_items_columns)
        output_file = os.path.join(self.save_folder, "line_items.xlsx")

        line_temp.to_excel(output_file, index=False)
        print("Excel file created successfully!")

    def save_templates(self):
        self.create_global_to_excel()
        self.create_line_to_excel()


if __name__ == "__main__":
    parser = InvoiceParser('1300457821.xml')
    parser.save_templates()
