import os
from django.shortcuts import render
from openpyxl import Workbook
from django.http import HttpResponse, JsonResponse
#from xml.etree.ElementTree import ET
import xml.etree.ElementTree as ET
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser,FormParser


class convert_xml_to_xlsx_APIView(APIView):
    parser_classes = (MultiPartParser,FormParser)
    # Create your views here.
    def post(self, request, *args, **kwargs):
            #get the file
                xml_file = request.FILES['file']
                if not xml_file:
                    return Response({"error": "No file provided"}, status=400)
            
                arc = ET.parse(xml_file)
                root = arc.getroot()
                '''print(head)
                print(len(head))
                print(dir(head))'''
                #create sheet
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "XML Transaction Data"
                headers = ['DATE', 'TRANSACTIONTYPE', 'VCHNO', 'REFTYPE','REFNO',
                  'REFDATE', 'DEBTOR', 'REFAMOUNT', 'AMOUNT', 'PARTICULARS','VCHTYPE','AMOUNTVERIFIED']
                
                sheet.append(headers)
                for voucher in root.findall('.//VOUCHER'):
                    # Extract voucher-level information
                    voucher_date = voucher.find('DATE').text if voucher.find('DATE') is not None else "NA"
                    voucher_no = voucher.find('VOUCHERNUMBER').text if voucher.find('VOUCHERNUMBER') is not None else "NA"
                    
                    voucher_type = voucher.find('VOUCHERTYPENAME').text if voucher.find('VOUCHERTYPENAME') is not None else "NA"
                    
                    #print(voucher_no,"sssssss")
                    
                    
                    for ledger_entry in voucher.findall('.//ALLLEDGERENTRIES.LIST'):
                            #ref_no = ledger_entry.find('NAME').text if ledger_entry.find('NAME') is not None else "NA"
                            ref_amount = ledger_entry.find('AMOUNT').text if ledger_entry.find('AMOUNT') is not None else "NA"
                            amount = ledger_entry.find('VATEXPAMOUNT').text if ledger_entry.find('VATEXPAMOUNT') is not None else "NA"
                            amount_verified = ledger_entry.find('ISINVOICE').text if ledger_entry.find('ISINVOICE') is not None else "NA"
                            
                            for allocate_entry in ledger_entry.findall('.//BILLALLOCATIONS.LIST'):
                                ref_no = allocate_entry.find('NAME').text if allocate_entry.find('NAME') is not None else "NA"
                                ref_type = allocate_entry.find('BILLTYPE').text if allocate_entry.find('BILLTYPE') is not None else "NA" 
                                particulars = voucher.find('PARTYNAME').text if voucher.find('PARTYNAME') is not None else "NA"
                                debtor = voucher.find('PARTYNAME').text if voucher.find('PARTYNAME') is not None else "NA"
                                ref_date = voucher.find('REFERENCEDATE').text if voucher.find('REFERENCEDATE') is not None else "NA"
                                print(debtor,"sssss")
                                if ref_type == "Agst Ref":
                                            transaction_type = "Child"
                                            amount_verified = "yes"
                                elif ref_type == "NA":
                                             transaction_type = "Parent" 
                                             amount_verified = "No"             
                                else:
                                        transaction_type = "Other" 
                                print(debtor,"partyname")        
                                row = [voucher_date,transaction_type,voucher_no, ref_type,ref_no,ref_date,
                            debtor, ref_amount, amount,particulars,voucher_type,amount_verified] 
                                if voucher_type == "Receipt":
                                    sheet.append(row)                         
                                #print(transaction_type,"ssss")
                            #print(ref_no,"rrrr")
                            #print(ref_no,"ssssssssssssssss")
                            
                    
                file_path = 'output.xlsx'
                #media_path = os.path.join('media', 'output.xlsx')

                workbook.save(file_path)
                return JsonResponse({'Success': 'File saved Successfully'}, status=200)


                
                


def sample_api(request):
    return HttpResponse("Sample response")


