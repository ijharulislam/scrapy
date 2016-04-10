from lxml import html
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re
from slimit import ast
from slimit.parser import Parser
from slimit.visitors import nodevisitor
from datetime import datetime
import json
from bson import json_util


def download(url):
    print('Downloading: ' + url)
    r = requests.get(url)

    print('Status code: ' + str(r.status_code))

    if r.status_code == requests.codes.ok: 
        return r.content
    else:
        return None



def write_to_excel(workbook,worksheet,outputs):
        
        # w = tzwhere.tzwhere()
        bold = workbook.add_format({'bold': True})
        bold_italic = workbook.add_format({'bold': True, 'italic':True})
        border_bold = workbook.add_format({'border':True,'bold':True})
        border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
        border = workbook.add_format({'border':True,'center_across':True})
        
        #worksheet = workbook.add_worksheet('%s_%s'%(a,j))
        worksheet.set_column('B:D', 22)
        worksheet.set_column('E:F', 33)
        row = 0
        col = 0


        worksheet.write(row,col,'Industry Data',bold)
        row = row + 1

        row = row + 2

        worksheet.write(row,col,'Sl No',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Name',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Work Phone',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Toll Free Phone',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Crisis Phone',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Fax',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Email',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Website',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Address 1',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Address 2',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Location',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Fees',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Hours',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Last Modified Date',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Language of Service',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Organization Type',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Eligibility',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'How To Apply',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Physical Access',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'About',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Tag',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Service Description Title',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Service Description',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Contact',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Latitude',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Longitude',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'City',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Province',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Postal Code',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'FB URL',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Twitter URL',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Google Plus',border_bold_grey)
        row = row + 1
        i = 0

        for out in outputs:
            for output in out:
                
                i = i + 1
                col = 0
                worksheet.write(row, col, i, border)
                col = col + 1
                worksheet.write(row, col, output["Name"] if output.has_key('Name') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Phone"] if output.has_key('Phone') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Toll-free phone"] if output.has_key('Toll-free phone') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Crisis phone"] if output.has_key('Crisis phone') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Fax"] if output.has_key('Fax') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Email"] if output.has_key('Email') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Website"] if output.has_key('Website') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Address 1"] if output.has_key('Address 1') else '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, output["Location (Intersection)"] if output.has_key('Location (Intersection)') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Fees"] if output.has_key('Fees') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Hours"] if output.has_key('Hours') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Last modified date"] if output.has_key('Last modified date') else '',border)
                #################
                col = col + 1
                worksheet.write(row, col, output["Language of service"] if output.has_key('Language of service') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Organization type"] if output.has_key('Organization type') else '',border)
                #################
                col = col + 1
                worksheet.write(row, col, output["Eligibility"] if output.has_key('Eligibility') else '',border)
                col = col + 1
                worksheet.write(row, col, output["How to apply"] if output.has_key('How to apply') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Physical access"] if output.has_key('Physical access') else '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, output["Service Description"] if output.has_key('Service Description') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Contact"] if output.has_key('Contact') else '',border)
                col = col + 1
                worksheet.write(row, col, output["lat"] if output.has_key('lat') else '',border)
                col = col + 1
                worksheet.write(row, col, output["lon"] if output.has_key('lon') else '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                worksheet.write(row, col, '',border)
                col = col + 1
                row = row + 1



def scrap(url):
    def get_url(href):
        return href and re.compile("searchNav=F").search(href)

    datas = []
    domain = 'http://www.ic.gc.ca'
    content = download(domain + url)
    if content:
        soup = BeautifulSoup(content, "lxml")
        first_url = soup.find_all(href=get_url)
        i = 0

        for link in first_url:
            output = {}
            print link.get('href')
            p = download(domain + link.get('href'))
        
            if p:
                sou = BeautifulSoup(p, "lxml")
                name = sou.find("h1", {"id": "cn-cont"})
                pretified_name = name.get_text().replace("\n","").strip()
                output["Name"] = pretified_name if pretified_name else ""
                # phone = sou.find("div", {"class": "ic2col2 width-50"})
                # pretified_phone = phone.get_text().replace("\n","").replace("\t","")
                # out["Phone"] = pretified_phone if pretified_phone else ""
                phone_and_fax = sou.find_all("div", {"class": "ic2col2 width-50"})
                if len(phone_and_fax) == 3 :

                    phone = phone_and_fax[0].get_text().replace("\n","").replace("\t","").strip()
                    output["Phone"] = phone if phone else ""
                    toll_free_phone = phone_and_fax[1].get_text().replace("\n","").replace("\t","").strip()
                    output["Toll-free phone"] = toll_free_phone if toll_free_phone else ""
                    fax = phone_and_fax[2].get_text().replace("\n","").replace("\t","").strip()
                    output["Fax"] = fax if fax else ""
                else:
                    phone = phone_and_fax[0].get_text().replace("\n","").replace("\t","").strip()
                    output["Phone"] = phone if phone else ""
                    fax = phone_and_fax[1].get_text().replace("\n","").replace("\t","").strip()
                    output["Fax"] = fax if fax else ""

                website = sou.find("a", {"class": "marginLeft20 noIcon font-small"})
                if website:
                    web = website.get_text()
                    output["Website"] = web if web else ""
                email = sou.find("a", {"class": "noIcon font-small"})
                if email:
                    mail = email.get_text()
                    output["Email"] = mail if mail else ""
                service = sou.find("p", {"class": "comment more"})
                if service:
                    service_description = service.get_text().replace("\n","").replace("\t","").replace("https://www.collegesinstitutes.ca/about/","").strip()

                    output["Service Description"] = service_description if service_description else ""
                add = sou.find_all("div", {"class": "span-4"})
                for ad in add:
                    text = ad.get_text()
                    search = re.search(r"Location Address:(?P<Address>[\w\-,.\s]+)",text)
                    if search:
                        loc = search.group(1).replace("Tel.","")
                        output["Address 1"]= " ".join(loc.split()) if loc else ""
                datas.append(output)
                # outs =[datas]
                # for out in outs:
                #     for o in out:
                #         print o
        return datas
    
                
                # outputs.append(out)
                # print outputs
        # return outputs
      



outputs = []

alpha = map(chr, range(ord('A'), ord('B')+1))
for i in alpha:
    url = '/app/ccc/sld/cmpny.do?letter=%s&lang=eng&profileId=21&tag=221001&letter=A'%(i)
    out = scrap(url)
    outputs.append(out)


workbook = xlsxwriter.Workbook('industry_canada_A_to_B.xlsx')
worksheet = workbook.add_worksheet('Industry')
write_to_excel(workbook,worksheet,outputs)
workbook.close()
# json_data = json.dumps(data,default=json_util.default)
# print data
# print type(data)
# for d in data:
#     print d 

#     data_list = scrap(url)
#     datas.append(data_list)
#     for data in datas:
#         for d in data:
#             print d
#     # print datas
# workbook = xlsxwriter.Workbook('industry_canada.xlsx')
# worksheet = workbook.add_worksheet('Industry')

# write_to_excel(workbook,worksheet,datas)
# workbook.close()
# import json
# from bson import json_util
# print json.dumps(datas,default=json_util.default)



