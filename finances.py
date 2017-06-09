from settings import *

import sys

sys.path.append(PYOOCALCPATH)

import pyoocalc
import datetime
import json
import re
import logging
import os
import PyPDF2


def open_spreadsheet(fname=FILENAME):
    '''Tries to start LibreOffice and open fname. Returns fname as a document.'''
    
    doc = None
    
    try:
        doc = pyoocalc.Document(autostart=True)
    except OSError as e:
        logger.error("ERROR:", e.errno, e.strerror)
    except pyoocalc.NoConnectException as e:
        logger.error("ERROR: The OpenOffice.org process is not started or does not listen on the resource.\n\ {0}\n\n\Start LibreOffice/OpenOffice in listening mode, example:\n\soffice \--accept=\"socket,host=localhost,port=2002;urp;\"\n".format(e.Message))

    try:
        doc.open_document(fname)
        return doc
    except OSError as e:
        logger.error("ERROR: Could not open supplied file name", e.errno, e.strerror)
        return None
##    finally:
##            doc.close_document()


def get_categories(fname=CATEGORIES_FILENAME):
    with open(fname) as json_data:
        d = json.load(json_data)
    return d

def import_new(*im_funcs):
    '''Calls specific import functions for different pieces of data'''

    for i in im_funcs:
        i()
    
def im_tsp(tsp_path=TSP_PATH):
    '''Reads any TSP asset lists, and updates then relevant fields in the spreadsheet'''

    sheet_icpr = doc.sheets.sheet(SHEET_ICPR)
    
    #Enumerate all files in target directory    
    _ = [f for f in os.listdir(tsp_path) if os.path.isfile(os.path.join(tsp_path, f))]

    #Cull for correct file names (properly formatted file should have the name TSP_Balance_YYYY_MM_DD.csv)
    __ = [f for f in _ if (f[-3:]=="csv" and f[:11]=="TSP_Balance")]

    #Find most recent entry
    i, v = 0, 0
    
    while v!=None:
        v = sheet_icpr.cell_value_by_index(0,2+i)
        i = i + 1

    most_recent_date = datetime.datetime.strptime(sheet_icpr._oSheet.getCellByPosition(0,i).getString(),'%m/%d/%y').date()
    next_row = i + 1
    
    #Cull for unrecorded TSP statements
    files_to_record = [f for f in __ if datetime.datetime.strptime(f[12:22],"%Y_%m_%d").date() > most_recent_date]

    #Get data from files
    for f in files_to_record:
        full_path = os.path.join(tsp_path,f)
        
        logger.info("Importing TSP data from: " + full_path)
        
        tsp_file = open_spreadsheet(full_path)
        sh = tsp_file.sheets.sheet(0)

        #Pull out data
        L2050 = sh.cell_value_by_index(3,3)
        G = sh.cell_value_by_index(3,8)
        C = sh.cell_value_by_index(3,10)
        S = sh.cell_value_by_index(3,11)
        I = sh.cell_value_by_index(3,12)

        tsp_file.close_document()
        
        #Put data in main spreadsheet
        logger.info("Writing TSP data from: " + full_path + " to main spreadsheet")

        sheet_icpr.set_cell_value_by_index(L2050, 9, next_row ,is_formula=False)     
        sheet_icpr.set_cell_value_by_index(G, 5, next_row, is_formula=False)
        sheet_icpr.set_cell_value_by_index(C, 6, next_row, is_formula=False)
        sheet_icpr.set_cell_value_by_index(S, 7, next_row, is_formula=False)
        sheet_icpr.set_cell_value_by_index(I, 8, next_row, is_formula=False)

        formula_str = "=SUM(F" + str(next_row + 1) + ":J" + str(next_row + 1) + ")"
        sheet_icpr.set_cell_value_by_index(formula_str, 3, next_row ,is_formula=True)

        sheet_icpr.set_cell_value_by_index("Biweekly", 1, next_row ,is_formula=False)

        #FIX: Date is being goofy
        sheet_icpr.set_cell_value_by_index(f[17:19] + "/" + f[20:22] + "/" + f[12:16], 0, next_row ,is_formula=False)
        
        next_row = next_row + 1

def im_cc():
    pass

def im_cash():
    pass

def im_paystub(force_reimport=False, paystub_path=PAYSTUB_PATH):
    '''Reads any unread paystubs; if force_reimport is True then rereads all existing paystubs'''
    
    sheet_ect = doc.sheets.sheet(SHEET_ECT)
    PAYSTUB_OFFSET=25
    
    #TODO: Reimport
    if force_reimport:
        pass

    try:
        target_date = datetime.datetime.strptime(sheet_ect._oSheet.getCellByPosition(2,1).getString(),'%m/%d/%y').date()
        filename = datetime.datetime.strftime(target_date + datetime.timedelta(days=PAYSTUB_OFFSET),'NFC_Paystub_%Y_%m_%d.pdf')
        
        full_path = os.path.join(paystub_path, filename)
    except:
        logger.error("Error trying to find paystub for biweek beginning: " + str(target_date))          
    
    pdf_obj = open(full_path, 'rb')
    logger.info("Importing paystub from: " + full_path)

    pdfr = PyPDF2.PdfFileReader(pdf_obj)

    #TODO: Find a better way to extract data than messy regex
    pdf_txt = pdfr.getPage(0).extractText()

    def validate_pdf_data(field, regex):
        try:
            _ = re.search(regex,pdf_txt)

            if _ != None:
                if field != "tmp_b" and field != "Beginning Date":
                    return float(_.group(0))
                elif field == "tmp_b" or field == "Beginning Date":
                    return str(_.group(0))                    
            else:
                logger.error("No value found for " + field + " while parsing: " + full_path)
                return None
        except:
          logger.error("PDF parsing error while parsing " + field + " in: " + full_path)  

    #Temporary holding; TODO: refactor by giving them individual expense/income lines
          
    #Hours Worked
    hours_worked = validate_pdf_data("Hours Worked", "(?<=GROSS PAY \*\*\*\*)\d\d.\d\d")
    print("Hours Worked: " + str(hours_worked))

    #Assumes 7 digit pay string
    gross_pay = validate_pdf_data("Gross Pay", "(?<=GROSS PAY \*\*\*\*\d\d.\d\d)(\d){4}.\d\d")
    print("Gross Pay: " + str(gross_pay))

    #Bonus
    inc_bonus_str = validate_pdf_data("Bonus", "(?<=CASH AWARD).{3,5}\.\d\d")
    tmp_b_str = validate_pdf_data("tmp_b", "(?<=CASH AWARD).{3,5}\.\d\d(.)*?\d\d(?=([A-Z]){3})")[:-2]

    #if inc_bonus and tmp_b are the same, there was a bonus in a previous pay period but not this one (regex matches the YTD bnous total)
    #TODO: Either find a better comaprison or validate
    if len(tmp_b_str) - len(str(inc_bonus_str)) <= 1:
        inc_bonus = 0
    else:
        inc_bonus = inc_bonus_str
    print("Bonus: " + str(inc_bonus))
    #print("TMPBonus: " + str(tmp_b_str))

    #TODO: OT

    #Retirement Contribution
    retirement = validate_pdf_data("Retirement Contribution", "(?<=RETIREMENT)(\d){3,5}\.\d\d")
    print("Retirement Contribution: " + str(retirement))

    #TSP Contribution
    tsp = validate_pdf_data("TSP Contribution", "(?<=ROTH TSP-FERS)(\d){3,5}\.\d\d")
    print("TSP Contribution: " + str(tsp))

    #Social Security
    ss = validate_pdf_data("Social Security", "(?<=SOCIAL SECURITY \(OASDI\))(\d){3,5}\.\d\d")
    print("Social Security: " + str(ss))

    #Federal Tax
    fed_tax = validate_pdf_data("Federal Tax", "(?<=FEDERAL TAX EXEMPTS S03)(\d){3,5}\.\d\d")
    print("Federal Tax: " + str(fed_tax))

    #State Tax
    state_tax = validate_pdf_data("State Tax", "(?<=ST TAX \w\w   EXEMPTS 001)(\d){3,5}\.\d\d")
    print("StateTax: " + str(state_tax))

    #Health Insurance Premium
    health_ins_premium = validate_pdf_data("Health Insurance Premium", "(?<=FEHBA - ENROLL CODE  \d\d\d)(\d){2,3}\.\d\d")
    print("Health Insurance Premium: " + str(health_ins_premium))

    #Vision Insurance Premium
    vision_ins_premium = validate_pdf_data("Vision Insurance Premium", "(?<=VISION PLAN)(\d){1,3}\.\d\d")
    print("Vision Insurance Premium: " + str(vision_ins_premium))

    #Misc
    misc = validate_pdf_data("Misc", "(?<=UNION.ASSOCIATION DUES \d\d \d\d\d\d)(\d){1,3}\.\d\d")
    print("Misc: " + str(misc))

    #Medicare
    medicare = validate_pdf_data("Medicare", "(?<=MEDICARE TAX WITHHELD)(\d){2}\.\d\d")
    print("Medicare: " + str(medicare))

    #Gym
    gym = validate_pdf_data("Gym", "(?<=DISCRETIONARY ALLOTMENT)(\d){2}\.\d\d")
    print("Gym: " + str(gym))

    #Beginning Date
    beg_date = validate_pdf_data("Beginning Date", "(?<=\*\*\*\*)\d\d/\d\d/\d\d\d\d")
    print("Beginning Date: " + str(beg_date))
    
    print("Deductions: " + str(retirement + tsp + ss + fed_tax + state_tax + health_ins_premium + vision_ins_premium + misc + medicare + gym))    
    print("MISC: " + str(misc))
    print("HC: " + str(health_ins_premium + vision_ins_premium + gym))
    print("RET: " + str(retirement + tsp))
    print("TAX: " + str(fed_tax + state_tax + medicare))
    
def read_old():
    expenses, income = [], []
    expenses.append(read_exp())
    income.append(read_inc())

def read_exp():
    sheet_cel = doc.sheets.sheet(SHEET_CEL)
    sheet_ccel = doc.sheets.sheet(SHEET_CCEL)
    sheet_oel = doc.sheets.sheet(SHEET_OEL)
    i, j, k = 0, 0, 0
    
    cel_row_offset, cel_col_offset, ccel_row_offset, ccel_col_offset, oel_row_offset, oel_col_offset  = 2, 2, 4, 5, 2, 2

    expenses = []   #List of tuples. Each tuple consists of (amount, date, category)

    logger.debug("Reading cash expenses...")
    try:
        #Get cash expenses
        while sheet_cel.cell_value_by_index(cel_col_offset, i + cel_row_offset) != None:
            ex = (sheet_cel.cell_value_by_index(cel_col_offset, i + cel_row_offset), #Amount
                 datetime.datetime.strptime(sheet_cel._oSheet.getCellByPosition(cel_col_offset + 1, i + cel_row_offset).getString(),'%m/%d/%y').date(), #Date. Note: Uses UNO method directly because OO returns a weird int (not the datetime ordinal?) instead of a date string
                 sheet_cel.cell_value_by_index(cel_col_offset + 2, i + cel_row_offset)) #Category
            expenses.append(ex)
            i=i+1
    except ValueError:
        logger.error("ValueError in cash expenses after: " + str(ex))

        logger.debug("Reading credit card expenses...")
    try:
        #Get credit card expenses
        while sheet_ccel.cell_value_by_index(ccel_col_offset, j + ccel_row_offset) != None:
            ex = (sheet_ccel.cell_value_by_index(ccel_col_offset, j + ccel_row_offset), #Amount
                 datetime.datetime.strptime(sheet_ccel._oSheet.getCellByPosition(ccel_col_offset - 3, j + ccel_row_offset).getString(),'%m/%d/%Y').date(), #NOTE DIFFERENT STRPTIME FMT. Also: Date. Note: Uses UNO method directly because OO returns a weird int (not the datetime ordinal?) instead of a date string
                 sheet_ccel.cell_value_by_index(ccel_col_offset + 2, j + ccel_row_offset)) #Category
            expenses.append(ex)
            j=j+1
    except ValueError:
        logger.error("ValueError in credit card expenses after: " + str(ex))

        logger.debug("Reading other expenses...")
    try:
        #Get other expenses
        while sheet_oel.cell_value_by_index(oel_col_offset, k + oel_row_offset) != None:
            ex = (sheet_oel.cell_value_by_index(oel_col_offset, k + oel_row_offset), #Amount
                 datetime.datetime.strptime(sheet_oel._oSheet.getCellByPosition(oel_col_offset + 1, k + oel_row_offset).getString(),'%m/%d/%y').date(), #NOTE DIFFERENT STRPTIME FMT. Also: Date. Note: Uses UNO method directly because OO returns a weird int (not the datetime ordinal?) instead of a date string
                 sheet_oel.cell_value_by_index(oel_col_offset + 2, k + oel_row_offset)) #Category
            expenses.append(ex)
            k=k+1
    except ValueError:
        logger.error("ValueError in other expenses after: " + str(ex))
        
    return expenses

def read_inc():
    pass

#def bin_period(dates, category_list, income, expenses):
#    pass

def bin_period_category(dates, category_name, expenses):
    beg_date = datetime.datetime.strptime(dates[0],'%m/%d/%y').date()  #Inclusive
    end_date = datetime.datetime.strptime(dates[1],'%m/%d/%y').date()  #Exclusive
    
    return [e for e in expenses if ((e[2] == category_name) and (e[1] >= beg_date) and (e[1] < end_date))]

def find_headers():
    sheet_ect = doc.sheets.sheet(SHEET_ECT)
    
    #Get offsets
    raw_headers = []

    #Load everything in the left column
    for i in range(10,100):
        raw_headers.append((sheet_ect.cell_value_by_index(0,i),i))

    #Regex to parse which are actual category headers
    def reg_mat(r):
        if r[0] != None:
            a = re.match("\[[A-Z]*\]",r[0])
            if a != None:
                return (a.group(0)[1:-1],r[1])
        else:
            return None
        
    actual_headers = [i for i in list(map(reg_mat,raw_headers)) if i != None]    

    return actual_headers
    
def populate_targets():
    sheet_ect = doc.sheets.sheet(SHEET_ECT)

    actual_headers = find_headers()

    cats = get_categories()

    for h in actual_headers:
        #Total target for the category is the sum of subcategory targets
        total = sum(list(map(float,cats[h[0]].values())))

        #Add target to spreadsheet
        sheet_ect.set_cell_value_by_index(total, 3, h[1] + 1,is_formula=False)

    #Add formulas
    #TODO: Un-hardcode these
    sheet_ect.set_cell_value_by_index("=D13-D12", 3, 13,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D16-D15", 3, 16,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D19-D18", 3, 19,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D22-D21", 3, 22,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D25-D24", 3, 25,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D28-D27", 3, 28,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D31-D30", 3, 31,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D34-D33", 3, 34,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D37-D36", 3, 37,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D40-D39", 3, 40,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D43-D42", 3, 43,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D46-D45", 3, 46,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D49-D48", 3, 49,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D52-D51", 3, 52,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D55-D54", 3, 55,is_formula=True)
    
def update_ect():
    sheet_ect = doc.sheets.sheet(SHEET_ECT)

    #Get latest date
    latest_date = datetime.datetime.strptime(sheet_ect._oSheet.getCellByPosition(2, 1).getString(),'%m/%d/%y').date()
    if latest_date + datetime.timedelta(days=TIME_INTERVAL) <= datetime.date.today():
        new_date = datetime.datetime.strftime(latest_date + datetime.timedelta(days=TIME_INTERVAL),'%m/%d/%y')
        logger.info("Adding new column for period beginning " + str(latest_date))
        
        sheet_ect._oSheet.Columns.insertByIndex(2,1) #Params are index, num cols?
        sheet_ect.set_cell_value_by_index(new_date, 2, 1, is_formula=False)

        #Perform this for each period
        populate_targets()
        update_expenses()
        update_main_fields()
        update_ect()    #Check again

def update_expenses():
    sheet_ect = doc.sheets.sheet(SHEET_ECT)
    headers = find_headers()

    cats = get_categories()

    beg_date = sheet_ect._oSheet.getCellByPosition(3, 1).getString()
    end_date = sheet_ect._oSheet.getCellByPosition(2, 1).getString()

    for h in headers:
        #Total target for the category is the sum of subcategory targets

        #Make a temporary function to sum a subcategory
        def tmp_bin_fnc(cat):
            return sum([e[0] for e in bin_period_category((beg_date, end_date),cat,exp)])

        #Sum all of the subcategories for the grand total
        total = sum(list(map(tmp_bin_fnc,list(cats[h[0]].keys()))))

        #Update spreadsheet with total
        sheet_ect.set_cell_value_by_index(total, 3, h[1],is_formula=False)

        print(h,total)

def update_main_fields():
    '''Updates income, and adds formulas for total expenses and net savings'''
    sheet_ect = doc.sheets.sheet(SHEET_ECT)

    #income = get_income(get_most_recent_period())
    sheet_ect.set_cell_value_by_index("=SUM(D5:D8)", 3, 3,is_formula=True)
    #TODO: Fill out cells 5-8

    sheet_ect.set_cell_value_by_index("=D12+D15+D18+D21+D24+D27+D30+D33+D36+D39+D42+D45+D51+D54", 3, 8,is_formula=True)
    sheet_ect.set_cell_value_by_index("=D4-D9", 3, 9,is_formula=True)
    
def get_most_recent_period():
    sheet_ect = doc.sheets.sheet(SHEET_ECT)

    beg_date = sheet_ect._oSheet.getCellByPosition(3, 1).getString()
    end_date = sheet_ect._oSheet.getCellByPosition(2, 1).getString()

    return (beg_date,end_date)

def get_income(dates):
    '''Returns the sum of all income for the given time period (inclusive, exclusive)'''

    #TODO: Get actual value
    return 0

def update_misc():
    '''Updates a few miscellaneous values on different sheets'''

    #logger.info("Misc. updates for period beginning " + str(new_date))
    #
    pass

def init_logger(l_fname=LOG_FILENAME):
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    
    # create console handler and set level to info
    handler = logging.StreamHandler()
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter("%(levelname)s - %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    # create error file handler and set level to error
    handler = logging.FileHandler(LOG_FILENAME, encoding=None, delay="true")
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter(fmt="%(asctime)s - %(levelname)s: %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    return logger

if __name__ == "__main__":

    #Init logger
    logger = init_logger(LOG_FILENAME)
    logger.info("--------------------------------")
    logger.info("Opening spreadsheet: " + FILENAME)
    logger.info("--------------------------------")
    
    #Open files
    doc = open_spreadsheet(FILENAME)
    cats = get_categories()

    #Import any new data available
    import_new(im_tsp, im_paystub, im_cash, im_cc)
    
    exp = read_exp()    
    update_ect()
    #update_misc()

    doc.save_document()
    doc.close_document()
