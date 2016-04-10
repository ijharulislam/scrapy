
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
# page = requests.get('https://services.findhelp.ca/eo/tcu/displayOrgInfo?resultsBean.selectedOrgID=133285&resultsBean.selectedLang=en')
# tree = html.fromstring(page.text)

# print tree
"""
    namedict = ({"first_name":"Joshua", "last_name":"Drake"},
            {"first_name":"Steven", "last_name":"Foo"},
            {"first_name":"David", "last_name":"Bar"})

    name = models.CharField(max_length=200)
    lat = models.FloatField(default=0.0)
    lon = models.FloatField(default=0.0)
    phone = models.CharField(max_length=100,blank=True,null=True)
    toll_free_phone = models.CharField(max_length=100,blank=True,null=True)
    crisis_phone = models.CharField(max_length=100,blank=True,null=True)
    fax = models.CharField(max_length=150,blank=True,null=True)
    email = models.EmailField(max_length=150,blank=True,null=True)
    website = models.CharField(max_length=150,null=True,blank=True)
    address = models.CharField(max_length=250,blank=True,null=True)
    mail_address = models.CharField(max_length=250,blank=True,null=True)
    fees = models.CharField(max_length=150,blank=True,null=True)
    hours = models.CharField(max_length=150,blank=True,null=True)
    last_modified = models.DateTimeField()
    language_of_service = models.CharField(max_length=300,blank=True,null=True)
    organization_type = models.CharField(max_length=300,blank=True,null=True)
    eligibility = models.CharField(max_length=200,blank=True,null=True)
    how_to_apply = models.TextField(max_length=500,blank=True,null=True)
    physical_access = models.CharField(max_length=300,blank=True,null=True)
    service_description = models.TextField(max_length=500,blank=True,null=True)
    approved = models.BooleanField(default=False)

{u'Website': 'http://www.cadets.ca', 
u'Physical access': u'Partially Accessible', 
u'Location (Intersection)': u'Borden CFB (By S-Parade Square Across the street from obstacle course and water tower)', 
u'How to apply': u'Registration required', 
u'Eligibility': u'12 years - 18 years', 
u'Fees': u'None', 
'lon': -79.896681, 
u'Hours': u'Office 8:30 am - 4:30 pm', 
'lat': 44.27697, 
u'Contact': u'Captain Edward Ross, Commanding Officer, 2408 CFLTC Royal Canadian Army Cadets, ph:705-424-1200 ext 7315, Hodskins.DC@forces.gc.ca', 
u'Organization type': u'Non Profit', 
u'Address': u'T-83, 51 Golan Rd, Borden, ON L0M 1C0', 
u'Office phone': u'705-424-1200 Ext 7361', 
u'Last modified date': u'30-Mar-15', 
u'Service description': u'Teaches citizenship, fitness, sensible living, leadership, discipline and skills through training in sports, first aid, adventure, band instruction, and military drills.  Activities include a biathlon team, a shooting program, summer camps, exchanges, and community service with local organizations. Meet Thursday evenings 6:45 pm-9:30 pm', 
u'Mail Address': u'Capt D. Hodskins,  PO Box 398, Borden, ON L0M 1C0', 
'Name': u'\n2408 CFSAL Royal Canadian Army Cadets\n\t\t\t'}

"""
def static_vars(**kwargs):
    def decorate(func):
        for k in kwargs:
            setattr(func, k, kwargs[k])
        return func
    return decorate

def download(url, max_retries=10):
    for i in range(max_retries):
        print('Downloading: ' + url)
        r = requests.get(url)

        print('Status code: ' + str(r.status_code))

        if r.status_code == requests.codes.ok: return r.content
    return None

def extract_to_dict(bocs):
    out = {}
    for boc in bocs:
        print "##################"
        print boc.prettify()
        a = []
        i = 0
        print "Number of Children %s"%len(list(boc.parent.children))
        for b in boc.parent.children:
            i = i + 1
            if b.string:
                print "%s::::%s"%(i,b.string.replace("\n","").replace("\t",""))
                # if b.string.replace("\n","").replace("\t","")=="Website":
                #     print b
                #     #a.append(b.string.replace("\n","").replace("\t",""))
                if i==2 or i==4:
                    a.append(b.string.replace("\n","").replace("\t",""))
                if len(a)==1 and a[0]=="Website" and i>2:
                    a.append(b.parent.a["href"])
        print a
        if len(a)==2:
            out["%s"%a[0]] = "%s"%a[1]
            print out
        print "##################"
    return out


def write_to_excel(workbook,worksheet,outputs,j,a):
        
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


        worksheet.write(row,col,'%s Page:%s'%(a,j+1),bold)
        row = row + 1

        row = row + 2

        i = 0
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
        worksheet.write(row,col,'Address',border_bold_grey)
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
        worksheet.write(row,col,'Modification Date',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Eligibility',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'How To Apply',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Physical Access',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Service Description',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Latitude',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Longitude',border_bold_grey)
        row = row + 1
        i = 0
        """
        {u'Physical access': u'Fully Accessible; Wheelchair accessible building', 
        u'Fees': u'None', 
        u'Location (Intersection)': u'Kitchener (Frederick and Irvin Sts)', 
        u'Area served': u'Waterloo Region * Wellington County', 
        u'Crisis phone': u'519-742-0867', 
        u'Email': u'erc@anishnabegoutreach.org', 
        u'Hours': u'Mon-Fri 9am-5pm * summer hours Mon-Fri 8am-4pm', 
        u'Contact': u'Christine Restoule, Employment Counsellor', 
        u'Organization type': u'Non Profit', 
        u'Address': u'151 Frederick St Unit 501, Kitchener, ON N2H 2M2', 
        u'Office phone': u'1-866-888-8808', 
        u'Last modified date': u'29-May-2015', 
        u'Toll-free phone': u'1-866-888-8808'}
        """
        for output in outputs:
            print output
            i =  i + 1
            col = 0
            worksheet.write(row, col, i, border)
            col = col + 1
            worksheet.write(row, col, output["Name"] if output.has_key('Name') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Office phone"] if output.has_key('Office phone') else '',border)
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
            worksheet.write(row, col, output["Address"] if output.has_key('Address') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Location (Intersection)"] if output.has_key('Location (Intersection)') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Fees"] if output.has_key('Fees') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Hours"] if output.has_key('Hours') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Last modified date"] if output.has_key('Last modified date') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Eligibility"] if output.has_key('Eligibility') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Physical access"] if output.has_key('Physical access') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Contact"] if output.has_key('Contact') else '',border)
            col = col + 1
            worksheet.write(row, col, output["lat"] if output.has_key('lat') else '',border)
            col = col + 1
            worksheet.write(row, col, output["lon"] if output.has_key('lon') else '',border)
            col = col + 1
            # worksheet.write(row, col, output[""],border)
            # col = col + 1

            row = row + 1
    
@static_vars(row=0)
def write_to_excel1(workbook,worksheet,outputs,j,a):
        
    # w = tzwhere.tzwhere()
    bold = workbook.add_format({'bold': True})
    bold_italic = workbook.add_format({'bold': True, 'italic':True})
    border_bold = workbook.add_format({'border':True,'bold':True})
    border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
    border = workbook.add_format({'border':True,'center_across':True})
    
    #worksheet = workbook.add_worksheet('%s_%s'%(a,j))
    worksheet.set_column('B:D', 22)
    worksheet.set_column('E:F', 33)
    #write_to_excel1.row = 0
    col = 0
    i = 0
    if j==0:
        worksheet.write(write_to_excel1.row,col,'%s Page:%s'%(a,j+1),bold)
        write_to_excel1.row = write_to_excel1.row + 1

        write_to_excel1.row = write_to_excel1.row + 2

        
        worksheet.write(write_to_excel1.row,col,'Sl No',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Area',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Name',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Work Phone',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Toll Free Phone',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Crisis Phone',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Fax',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Email',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Website',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Address',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Location',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Fees',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Hours',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Last Modified Date',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Language of Service',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Organization Type',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Modification Date',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Eligibility',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'How To Apply',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Physical Access',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Service Description',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Latitude',border_bold_grey)
        col = col + 1
        worksheet.write(write_to_excel1.row,col,'Longitude',border_bold_grey)
    write_to_excel1.row = write_to_excel1.row + 1
    #i = 0
    """
    {u'Physical access': u'Fully Accessible; Wheelchair accessible building', 
    u'Fees': u'None', 
    u'Location (Intersection)': u'Kitchener (Frederick and Irvin Sts)', 
    u'Area served': u'Waterloo Region * Wellington County', 
    u'Crisis phone': u'519-742-0867', 
    u'Email': u'erc@anishnabegoutreach.org', 
    u'Hours': u'Mon-Fri 9am-5pm * summer hours Mon-Fri 8am-4pm', 
    u'Contact': u'Christine Restoule, Employment Counsellor', 
    u'Organization type': u'Non Profit', 
    u'Address': u'151 Frederick St Unit 501, Kitchener, ON N2H 2M2', 
    u'Office phone': u'1-866-888-8808', 
    u'Last modified date': u'29-May-2015', 
    u'Toll-free phone': u'1-866-888-8808'}
    """
    for output in outputs:
        print output
        i =  i + 1
        col = 0
        worksheet.write(write_to_excel1.row, col, i, border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, a, border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Name"] if output.has_key('Name') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Office phone"] if output.has_key('Office phone') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Toll-free phone"] if output.has_key('Toll-free phone') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Crisis phone"] if output.has_key('Crisis phone') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Fax"] if output.has_key('Fax') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Email"] if output.has_key('Email') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Website"] if output.has_key('Website') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Address"] if output.has_key('Address') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Location (Intersection)"] if output.has_key('Location (Intersection)') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Fees"] if output.has_key('Fees') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Hours"] if output.has_key('Hours') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Last modified date"] if output.has_key('Last modified date') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Eligibility"] if output.has_key('Eligibility') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Physical access"] if output.has_key('Physical access') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["Contact"] if output.has_key('Contact') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["lat"] if output.has_key('lat') else '',border)
        col = col + 1
        worksheet.write(write_to_excel1.row, col, output["lon"] if output.has_key('lon') else '',border)
        col = col + 1
        # worksheet.write(write_to_excel1.row, col, output[""],border)
        # col = col + 1

        write_to_excel1.row = write_to_excel1.row + 1
    
def enter_into_db(data):
    import psycopg2

    # Try to connect
    """
    namedict = ({"first_name":"Joshua", "last_name":"Drake"},
            {"first_name":"Steven", "last_name":"Foo"},
            {"first_name":"David", "last_name":"Bar"})

    name = models.CharField(max_length=200)
    lat = models.FloatField(default=0.0)
    lon = models.FloatField(default=0.0)
    phone = models.CharField(max_length=100,blank=True,null=True)
    toll_free_phone = models.CharField(max_length=100,blank=True,null=True)
    crisis_phone = models.CharField(max_length=100,blank=True,null=True)
    fax = models.CharField(max_length=150,blank=True,null=True)
    email = models.EmailField(max_length=150,blank=True,null=True)
    website = models.CharField(max_length=150,null=True,blank=True)
    address = models.CharField(max_length=250,blank=True,null=True)
    mail_address = models.CharField(max_length=250,blank=True,null=True)
    fees = models.CharField(max_length=150,blank=True,null=True)
    hours = models.CharField(max_length=150,blank=True,null=True)
    last_modified = models.DateTimeField()
    language_of_service = models.CharField(max_length=300,blank=True,null=True)
    organization_type = models.CharField(max_length=300,blank=True,null=True)
    eligibility = models.CharField(max_length=200,blank=True,null=True)
    how_to_apply = models.TextField(max_length=500,blank=True,null=True)
    physical_access = models.CharField(max_length=300,blank=True,null=True)
    service_description = models.TextField(max_length=500,blank=True,null=True)
    approved = models.BooleanField(default=False)


    """
    try:
        conn=psycopg2.connect("dbname='jobdb' user='jobmin' password='f1d3r!@#' host='localhost'")
        print "Database Connected"
    except:
        print "I am unable to connect to the database."

    cur = conn.cursor()
    print data
    for o in data:
        try:
            print "$$$$$$$$$$$$$$$$$$$"
            print "%s,%s"%(o["lat"],o["lon"])
            print "$$$$$$$$$$$$$$$$$$$"
            name=o["Name"][0:200]
            lat=o["lat"]
            lon=o["lon"]
            phone=o["Office phone"][0:100] if o.has_key("Office phone") else ""
            toll_free_phone=o["Toll-free phone"][0:100] if o.has_key("Toll-free phone") else ""
            if o.has_key("Crisis phone"):
                crisis_phone = o["Crisis phone"][0:100]
            else:
                if o.has_key("Office phone"):
                    crisis_phone=o["Office phone"][0:100]
                else:
                    if o.has_key("Toll-free phone"):
                        crisis_phone=o["Toll-free phone"][0:100]
                    else:
                        crisis_phone = ""
            fax=o["Fax"][0:100] if o.has_key("Fax") else ""
            email=o["Email"][0:150] if o.has_key("Email") else "sample@sample.com"
            website=o["Website"][0:150] if o.has_key("Website") else "http://www.sample.com"
            if o.has_key("Location (Intersection)"):
                address_location = o["Location (Intersection)"][0:250]
            else:
                address_location = o["Address"][0:250] if o.has_key("Address") else ""
            address=o["Address"][0:250] if o.has_key("Address") else ""
            
            if o.has_key("Mail Address"): 
                mail_address=o["Mail Address"][0:250] 
            else: 
                mail_address=o["Address"][0:250] if o.has_key("Address") else ""
            fees=o["Fees"][0:150] if o.has_key("Fees") else "NA"
            hours=o["Hours"][0:150] if o.has_key("Hours") else ""
            language_of_service=o["Language of service"][0:300] if o.has_key("Language of service") else "English"
            organization_type=o["Organization type"][0:300] if o.has_key("Organization type") else "NA"
            eligibility=o["Eligibility"][0:200] if o.has_key("Eligibility") else "NA"
            how_to_apply=o["How to apply"][0:500] if o.has_key("How to apply") else "NA"
            physical_access=o["Physical access"][0:300] if o.has_key("Physical access") else "NA"
            service_description=o["Service description"][0:500] if o.has_key("Service description") else "NA"
            contact = o["Contact"][0:500] if o.has_key("Contact") else "NA"
            zip_code = o["Address"][-7:] if o.has_key("Address") else ""
            approved=True
            last_modified=datetime.now()
            cur.execute(
                """INSERT INTO 
                agency_agency(name,lat,lon,phone,toll_free_phone,crisis_phone,fax,email,website,address,address_location,mail_address,fees,hours,language_of_service,organization_type,eligibility,how_to_apply,physical_access,service_description,contact,approved,last_modified,zip_code) 
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",(
                    name,lat,lon,phone,toll_free_phone,crisis_phone,fax,email,website,address,address_location,mail_address,fees,hours,language_of_service,organization_type,eligibility,how_to_apply,physical_access,service_description,contact,approved,last_modified,zip_code))
            print "Inserted to Agency"
        except Exception, e:
            print "I can't INSERT into Agency", e.message
    
    conn.commit()

def write_to_excel2(workbook,worksheet,datas):
        
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


        worksheet.write(row,col,'Ontario Data',bold)
        row = row + 1

        row = row + 2

        i = 0
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
        for outputs in datas:
            for output in outputs:
                print output
                i =  i + 1
                col = 0
                worksheet.write(row, col, i, border)
                col = col + 1
                worksheet.write(row, col, output["Name"] if output.has_key('Name') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Office phone"] if output.has_key('Office phone') else '',border)
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
                worksheet.write(row, col, output["Address"] if output.has_key('Address') else '',border)
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
                worksheet.write(row, col, output["Service description"] if output.has_key('Service description') else '',border)
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

def call_main_url(domain,url,j=0,a="Ontario"):
    outputs = []
    page = download(domain+url)
    #print page
    if page:
        soup = BeautifulSoup(page, "lxml")
        boccats = soup.find("tbody")
        i = 0
        # for boccat in boccats:
        text_to_find = "addPointToMap"
        bc = soup.find_all("td", " namewidth")
        #print bc

        soupscript = BeautifulSoup(page)
        scripts = soupscript.find_all('script')
        locations = []
        for script in scripts:
            pattern = r"addPointToMap"
            pattern1 = r'addPointToMap\("(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?)"\)'
            list_coords = re.findall(pattern1, script.text)

            if list_coords:
                for j in list_coords:
                    print j
                    print "%s:[%s,%s]"%(j[1],j[3],j[4])
                    try:
                        locations.append({
                            "id":re.sub('["\s]', '', j[1]),
                            "lat":float(re.sub('["\s]', '', j[3])),
                            "lon":float(re.sub('["\s]', '', j[4])),
                            "name":re.sub('["\s]', '', j[6])
                            })
                    except Exception, e:
                        locations.append({
                            "id":re.sub('["\s]', '', j[1]),
                            "lat":0.0,
                            "lon":0.0,
                            "name":re.sub('["\s]', '', j[6])
                            })
                        pass
        print locations

        for boccat in bc:
            i = i + 1
            print "CALLING "+boccat.a["href"]
            splitted_hrefs = boccat.a["href"].split("/")
            print splitted_hrefs[4]
            print "#####################"
            p = download(domain+boccat.a["href"])
            if p:
                sou = BeautifulSoup(p, "lxml")
                bocs = sou.find_all("td","resulttext")
                print "*******  DATA  %s      *********"%i
                output = extract_to_dict(bocs)
                print "*******  LARGEBOLDENDDATA  %s   *********"%i
                b = sou.find("td","largebold")
                print b.contents
                print b.get_text()
                print dir(b)
                output["Name"] = b.get_text() if b.get_text() else ""
                for locs in locations:
                    if locs["id"] == splitted_hrefs[4]:
                        output["lat"] = locs["lat"]
                        output["lon"] = locs["lon"]
                print "*******  ENDDATA  %s   *********"%i
                print output

                outputs.append(output)
        #write_to_excel(workbook,outputs,j,a)
        #write_to_excel1(workbook,worksheet,outputs,j,a)
        #enter_into_db(outputs)
    return outputs

    #for boc in bocs:
    
domain = 'http://services.findhelp.ca'
url = '/eo/en/quick?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=Ontario&resultsBean.currentPage=0&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career'
# workbook = xlsxwriter.Workbook('Canada_Combined.xlsx')
workbook = xlsxwriter.Workbook('ontario05.xlsx')
worksheet = workbook.add_worksheet('Canada')
areas = [
    'Ontario',
    # 'Toronto',
    # 'Toronto+South',
    # 'Toronto+Central',
    # 'Toronto+East',
    # 'Toronto+West',
    # 'Greater+Toronto+Area',
    # 'Edmonton',
    # 'Edmunston',
    # 'Windsor',
    # 'Windsor+South',
    # 'Windsor+Central',
    # 'Windsor+East',
    # 'Windsor+West',
    # 'Toronto+(City+of)',
    ]
datas = []
for a in areas:
    a = ""
    #for i in range(478):    
    for i in range(400,550):    
        # if i==0:
        #     url1 = '/eo/en/quick?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=%s&resultsBean.currentPage=%s&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career'%(a,i)

        # else:
        #     url1 = '/eo/en/quick?multiBean.message=Employment+Service+including+Second+Career&resultsBean.index=11&resultsBean.showAreaServed=false&resultsBean.currentPage=%s&resultsBean.textView=false&multiBean.program=PR034&commonBean.langTxt=ENGLISH&multiBean.newSearch=false&multiBean.client=CL000&multiBean.zipCode=&commonBean.location=%s&multiBean.includeEng=Y'%(i,a)
        url1 = '/eo/en/multi?program=PR000&client=&location=%s&servingRegion=false&orderBy=ORGANIZATION&includeEnglish=true&print=false&pageNum=%s&showMap=true'%(a,i)
        data_list = call_main_url(domain,url1,i,a)
        datas.append(data_list)
import json
from bson import json_util
print json.dumps(datas,default=json_util.default)

write_to_excel2(workbook,worksheet,datas)


workbook.close()
#tree = html.fromstring(page)
"""
http://services.findhelp.ca/eo/en/multi?program=PR034&client=&location=Ontario&servingRegion=false&orderBy=ORGANIZATION&includeEnglish=true&print=false&pageNum=1&showMap=true
http://services.findhelp.ca/eo/en/multi?print=false&servingRegion=false&showMap=true&client=&orderBy=ORGANIZATION&location=Ontario&program=PR034&pageNum=2&includeEnglish=true


http://services.findhelp.ca/eo/en/multi?print=false&servingRegion=false&showMap=true&client=&orderBy=ORGANIZATION&location=&program=PR000&pageNum=2&includeEnglish=true

http://services.findhelp.ca/eo/en/multi?print=false&servingRegion=false&showMap=true&client=&orderBy=ORGANIZATION&location=Ontario&program=PR000&pageNum=1&includeEnglish=true
http://services.findhelp.ca/eo/en/multi?print=false&servingRegion=false&showMap=true&client=&orderBy=ORGANIZATION&location=Ontario&program=PR000&pageNum=2&includeEnglish=true
http://services.findhelp.ca/eo/en/multi?print=false&servingRegion=false&showMap=true&client=&orderBy=ORGANIZATION&location=Ontario&program=PR000&pageNum=3&includeEnglish=true
http://services.findhelp.ca/eo/en/multi?print=false&servingRegion=false&showMap=true&client=&orderBy=ORGANIZATION&location=Ontario&program=PR000&pageNum=4&includeEnglish=true

##http://services.findhelp.ca/eo/en/quick?query=&location=Ontario&servingRegion=false&orderBy=PROXIMITY&includeEnglish=true&print=false&pageNum=1&showMap=true
http://services.findhelp.ca/eo/en/quick?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=Edmonton&resultsBean.currentPage=0&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career

http://services.findhelp.ca/eo/tcu/msearch?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=Edmonton&resultsBean.currentPage=0&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career
http://services.findhelp.ca/eo/tcu/msearch?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=Ontario&resultsBean.currentPage=0&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career
http://services.findhelp.ca/eo/tcu/msearch?multiBean.message=Employment%2BService%2Bincluding%2BSecond%2BCareer&resultsBean.index=11&resultsBean.showAreaServed=false&resultsBean.currentPage=1&resultsBean.textView=false&multiBean.program=PR034&commonBean.langTxt=ENGLISH&multiBean.newSearch=false&multiBean.client=CL000&multiBean.zipCode=&commonBean.location=Ontario&multiBean.includeEng=Y
http://services.findhelp.ca/eo/tcu/msearch?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=Toronto&resultsBean.currentPage=0&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career
http://services.findhelp.ca/eo/tcu/msearch?multiBean.message=Employment%2BService%2Bincluding%2BSecond%2BCareer&resultsBean.index=11&resultsBean.showAreaServed=false&resultsBean.currentPage=1&resultsBean.textView=false&multiBean.program=PR034&commonBean.langTxt=ENGLISH&multiBean.newSearch=false&multiBean.client=CL000&multiBean.zipCode=&commonBean.location=Toronto&multiBean.includeEng=Y
http://services.findhelp.ca/eo/tcu/msearch?resultsBean.textView=false&resultsBean.showAreaServed=false&commonBean.location=Windsor%20South&resultsBean.currentPage=0&resultsBean.orderType=ORGANIZATION&multiBean.program=PR034&multiBean.client=CL000&resultsBean.index=1&multiBean.includeEng=Y&commonBean.langTxt=ENGLISH&multiBean.zipCode=&multiBean.newSearch=true&multiBean.message=Employment+Service+including+Second+Career
http://services.findhelp.ca/eo/tcu/msearch?multiBean.message=Employment%2BService%2Bincluding%2BSecond%2BCareer&resultsBean.index=11&resultsBean.showAreaServed=false&resultsBean.currentPage=1&resultsBean.textView=false&multiBean.program=PR034&commonBean.langTxt=ENGLISH&multiBean.newSearch=false&multiBean.client=CL000&multiBean.zipCode=&commonBean.location=Windsor+South&multiBean.includeEng=Y
"""
#buyers = tree.xpath('//td[@class=" namewidth"]/text()')
#print tree
#print buyers

