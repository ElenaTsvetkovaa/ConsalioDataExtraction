import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import os



class XMLParser:
    def __init__(self, xml_file):
        self.xml_file = xml_file
        self.ns = {
            'rsm': "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
            'ram': "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
            'udt': "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100"
        }
        self.root = self.get_root()

    def get_root(self):
        tree = ET.parse(self.xml_file)

        return tree.getroot()

    def find(self, path):
        return self.root.find(path , self.ns)

    def findall(self, path):
        return self.root.findall(path, self.ns)



class GlobalExtractor:

    def __init__(self, parser):
        self.parser = parser
        self.global_invoice_data = []

    def extract(self):

        for invoice in parser.root.findall('rsm:ExchangedDocument', parser.ns):
            invoice_number = invoice.find('ram:ID', parser.ns).text
            raw_date = datetime.strptime(
                invoice.find('ram:IssueDateTime/udt:DateTimeString', parser.ns).text, "%Y%m%d")
            invoice_date = raw_date.strftime("%d.%m.%Y")

            #   Recipient Info
            recipient_path = parser.root.find('.//ram:BuyerTradeParty', parser.ns)
            recipient = recipient_path.find('ram:Name', parser.ns).text
            recipient_person = recipient_path.find('ram:DefinedTradeContact/ram:PersonName', parser.ns).text

            trade_settlement = parser.root.find('.//ram:ApplicableHeaderTradeSettlement', parser.ns)
            currency = trade_settlement.find('ram:InvoiceCurrencyCode', parser.ns).text
            total_vat_excluded = trade_settlement.find('.//ram:LineTotalAmount', parser.ns).text
            total_vat_included = trade_settlement.find('.//ram:GrandTotalAmount', parser.ns).text

            self.global_invoice_data.append(
                [recipient, "", invoice_number, "", "", "", "", currency,
                 total_vat_excluded, total_vat_included, invoice_date, "", "", "",
                 recipient_person, "", "", "", "", "", "", "", ""]
            )

        return self.global_invoice_data



class LineItemExtractor:

    def __init__(self, parser):
        self.parser = parser
        self.line_items_data = []

    def extract(self):

        for line_data in parser.root.findall('.//ram:IncludedSupplyChainTradeLineItem', parser.ns):
            description = line_data.find('.//ram:SpecifiedTradeProduct/ram:Name', parser.ns).text
            unit_price = line_data.find(
                './/ram:SpecifiedLineTradeAgreement/ram:NetPriceProductTradePrice/ram:ChargeAmount', parser.ns).text
            quantity = line_data.find('.//ram:SpecifiedLineTradeDelivery/ram:BilledQuantity', parser.ns).text
            line_price = line_data.find('.//ram:SpecifiedTradeSettlementLineMonetarySummation/ram:LineTotalAmount',
                                        parser.ns).text
            vat = line_data.find('.//ram:SpecifiedLineTradeSettlement/ram:ApplicableTradeTax/ram:RateApplicablePercent',
                                 parser.ns).text

            self.line_items_data.append(
                ["", "", "", "", description, quantity, unit_price, line_price, vat, "", "", "", "",
                 "", "", "", ""]
            )

        return self.line_items_data


class ExcelExporter:

    def __init__(self, save_folder='templates'):
        self.save_folder = save_folder


    def create_global_to_excel(self, global_invoice_data):
        global_tem_columns = [
            "Recipient", "Recipient Parent", "Invoice Number", "Invoice Type", "Pricing Type",
            "RVG - Dispute Value", "RVG - Fee Rate", "Currency", "Total VAT Excluded",
            "Total VAT Included", "Invoice Date", "Due Date", "Period Start", "Period End",
            "Recipient Display Name", "Recipient First Name", "Recipient Last Name",
            "Issuer Display Name", "Issuer First Name", "Issuer Last Name", "Reference / PO Number",
            "Is Draft", "Debit Invoice Number"
        ]
        global_temp = pd.DataFrame(global_invoice_data, columns=global_tem_columns)
        output_file = os.path.join(self.save_folder, "global.xlsx")

        global_temp.to_excel(output_file, index=False)

    def create_line_to_excel(self, line_items_data):
        line_items_columns = [
            "Date Start", "Date End", "Is Hours", "Code", "Description",
            "Quantity", "Unit Price", "Line Price", "VAT Rate",
            "Cost Type", "Cost Centre", "Project 1", "Project 2", "Project 3",
            "Workstream 1", "Workstream 2", "Workstream 3"
        ]

        line_temp = pd.DataFrame(line_items_data, columns=line_items_columns)
        output_file = os.path.join(self.save_folder, "line_items.xlsx")

        line_temp.to_excel(output_file, index=False)
        print("Excel file created successfully!")

    def save_templates(self, global_invoice_data, line_items_data):
        self.create_global_to_excel(global_invoice_data)
        self.create_line_to_excel(line_items_data)


class TemplateManager:

    def __init__(self, global_extractor, line_extractor, exporter):
        self.global_extractor = global_extractor
        self.line_extractor = line_extractor
        self.exporter = exporter

    def process(self):
        """Extracts data and saves it to Excel."""
        global_data = self.global_extractor.extract()
        line_items = self.line_extractor.extract()
        self.exporter.save_templates(global_data, line_items)


if __name__ == "__main__":
    parser = XMLParser('1300457821.xml')

    global_extractor = GlobalExtractor(parser)
    line_extractor = LineItemExtractor(parser)

    exporter = ExcelExporter()
    processor = TemplateManager(global_extractor, line_extractor, exporter)
    processor.process()

