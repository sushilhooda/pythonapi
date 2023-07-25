from flask import Flask, jsonify
import pythoncom
app = Flask(__name__)
pythoncom.CoInitialize()
# Your existing Python code goes here (e.g., the function that processes the data)

import pandas as pd
import numpy as np
import os
from datetime import date
from datetime import timedelta
import calendar
desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive - PROCLOZ SERVICES PVT LTD\Desktop\iBPRO\Payroll Invoices")
os.chdir(desktop_path)
input_path=desktop_path
excel_raw_data=input_path+"\word invoices and Excel Raw Data"

def start_invoice_num():
    #starting_invoice_number=int(input("ENTER the starting invoice number example:10479 for 479 invoice number\n")) #enter here the starting invoice number
    starting_invoice_number=10444 #enter here the starting invoice number
    return starting_invoice_number


def customer_sheet_auto():
    os.chdir(desktop_path+"\Do_Not_Delete\system files")
    import cusotmer_contact_format_changes


def india_process_data():
    try:
        import pandas as pd
        import numpy as np
        import os
        from datetime import date
        from datetime import timedelta
        import calendar
        desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive - PROCLOZ SERVICES PVT LTD\Desktop\iBPRO\Payroll Invoices")
        os.chdir(desktop_path)
        input_path=desktop_path
        india=pd.read_excel("India RAW Data Format.xlsx",skiprows=1)
        india["costcenter"]=india["costcenter"].apply(lambda x:x.replace("Papaya Global, Inc. (External)","Papaya Global, Inc.  "))
        india["costcenter"]=india["costcenter"].apply(lambda x:x.replace("Papaya Global (HK) Limited (External)","Papaya Global (HK) Limited   "))
        india["costcenter"]=india["costcenter"].apply(lambda x:x.replace("Papaya Global, Inc. (Internal)","Papaya Global, Inc.    "))
        india["costcenter"]=india["costcenter"].apply(lambda x:x.replace("Papaya Global (HK) Limited (Internal)","Papaya Global (HK) Limited     "))
        india_col=pd.read_excel("India Col Mapping.xlsx")
        
        india_col['Comment']=india_col['Comment'].apply(lambda x: "DO NOT IMPORT" if pd.isnull(x)==True else x)
        india_col=india_col[india_col['Comment']!="DO NOT IMPORT"]
        #india_col['Country Name']='India'
        
        ind_col=list(india_col['Column Name'])    
        
        #final India Data using loop
        
        India=pd.DataFrame()
        for col in ind_col:
            if col in india.columns:
                India[col]=india[col]
        
        #India=india[ind_col] before above loop this was the logic to build final dataframe
        India['Country Name']="India"
        India["Entity Name"]="India"
        India["Local Currency"]="INR"
        India["Supplier Name"]="Procloz Services Private Limited"
        #India["Service_Fee_TBC"]=India["Total in Foreign Currency"]
        
        #renaming costcenter and client to cost center and division to have column name same across
        
        India.rename(columns ={"costcenter":"Cost Center"},inplace=True)
        
        # =============================================================================
        # Reading India Rate Sheet here and Exchange Rate will be mapped basis this
        # =============================================================================
        
        india_ex_rate=pd.read_excel("Rate Sheet.xlsx",sheet_name="India")
        india_ex_rate.drop(columns=["Local Currency"],inplace=True)
        
        India=India.merge(india_ex_rate,on="Invoicing Currency",how="left")
        
        # =============================================================================
        # Generating Multiple Excel File based on Raw Data Basis Division for India
        # =============================================================================
        
        for i in range(len(India["Division"].unique())):
            division_name=India["Division"].unique()[i]
            new_dataframe=India[India['Division']==division_name]
            new_dataframe.to_excel(excel_raw_data+"\\"+division_name.replace("/","")+".xlsx",index=False)    
        
        #India.rename(columns={"Client":"Division"},inplace=True)
        
        return India,india_col
    except Exception as e:
        return None, print(f"error occured while reading India file {e}")


def bangladesh_process_data():
    try:
        import pandas as pd
        import numpy as np
        import os
        from datetime import date
        from datetime import timedelta
        import calendar
        desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive - PROCLOZ SERVICES PVT LTD\Desktop\iBPRO\Payroll Invoices")
        os.chdir(desktop_path)
        input_path=desktop_path
        bang=pd.read_excel("Bangladesh RAW Data Format.xlsx",skiprows=1)
        bang["Cost Center"]=bang["Cost Center"].apply(lambda x:x.replace("Papaya Global, Inc. (External)","Papaya Global, Inc.  "))
        bang["Cost Center"]=bang["Cost Center"].apply(lambda x:x.replace("Papaya Global (HK) Limited (External)","Papaya Global (HK) Limited   "))
        bang["Cost Center"]=bang["Cost Center"].apply(lambda x:x.replace("Papaya Global, Inc. (Internal)","Papaya Global, Inc.    "))
        bang["Cost Center"]=bang["Cost Center"].apply(lambda x:x.replace("Papaya Global (HK) Limited (Internal)","Papaya Global (HK) Limited     "))
        
        bang_col=pd.read_excel("Bangla Col Mapping.xlsx")
        
        bang_col['Comment']=bang_col['Comment'].apply(lambda x: "DO NOT IMPORT" if pd.isnull(x)==True else x)
        bang_col=bang_col[bang_col['Comment']!="DO NOT IMPORT"]
        #bang_col['Country Name']='Bangladesh'
        
        ban_col=list(bang_col['Column Name'])
        
        #final Bangladesh Data
        Bangladesh=pd.DataFrame()
        for col in ban_col:
            if col in bang.columns:
                Bangladesh[col]=bang[col]
        
        #Bangladesh=bang[ban_col] #before above loop this was the logic
        Bangladesh['Country Name']="Bangladesh"
        Bangladesh["Entity Name"]="Bangladesh"
        Bangladesh["Local Currency"]="BDT"
        Bangladesh["Supplier Name"]="Procloz Private Limited"
        #Bangladesh["Service_Fee_TBC"]=Bangladesh["Total Gross"]
        
        bangla_ex_rate=pd.read_excel("Rate Sheet.xlsx",sheet_name="Bangladesh")
        bangla_ex_rate.drop(columns=["Local Currency"],inplace=True)
        
        Bangladesh=Bangladesh.merge(bangla_ex_rate,on="Invoicing Currency",how="left")

        # =============================================================================
        # Generating Multiple Excel File based on Raw Data Basis Division for Bangladesh
        # =============================================================================  
        for i in range(len(Bangladesh["Division"].unique())):
            division_name=Bangladesh["Division"].unique()[i]
            new_dataframe=Bangladesh[Bangladesh['Division']==division_name]
            new_dataframe.to_excel(excel_raw_data+"\\"+division_name.replace("/","")+".xlsx",index=False)  
        
        return Bangladesh,bang_col  
    except FileNotFoundError:
        return None, print("Bangladesh File or column mapping not found")

        
def australia_process_data():
    try:
        import pandas as pd
        import numpy as np
        import os
        from datetime import date
        from datetime import timedelta
        import calendar
        desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive - PROCLOZ SERVICES PVT LTD\Desktop\iBPRO\Payroll Invoices")
        os.chdir(desktop_path)
        input_path=desktop_path
        aus=pd.read_excel("Australia RAW Data Format.xlsx",skiprows=1)
        aus["Cost Center"]=aus["Cost Center"].apply(lambda x:x.replace("Papaya Global, Inc. (External)","Papaya Global, Inc.  "))
        aus["Cost Center"]=aus["Cost Center"].apply(lambda x:x.replace("Papaya Global (HK) Limited (External)","Papaya Global (HK) Limited   "))
        aus["Cost Center"]=aus["Cost Center"].apply(lambda x:x.replace("Papaya Global, Inc. (Internal)","Papaya Global, Inc.    "))
        aus["Cost Center"]=aus["Cost Center"].apply(lambda x:x.replace("Papaya Global (HK) Limited (Internal)","Papaya Global (HK) Limited     "))
        
        aus_col=pd.read_excel("Aus Col Mapping.xlsx")
        
        aus_col['Comment']=aus_col['Comment'].apply(lambda x:"DO NOT IMPORT" if pd.isnull(x)==True else x)
        aus_col=aus_col[aus_col['Comment']!="DO NOT IMPORT"]
        #aus_col['Country Name']="Australia"
        
        aust_col=list(aus_col['Column Name'])
        
        # generating australia final data
        Australia=pd.DataFrame()
        for col in aust_col:
            if col in aus.columns:
                Australia[col]=aus[col]

        #Australia=aus[aust_col] # before above loop this was the logic for final data
        Australia['Country Name']="Australia"
        Australia["Entity Name"]="Australia"
        Australia["Local Currency"]="AUD"
        Australia["Supplier Name"]="Procloz Pty Ltd"
        #Australia["Service_Fee_TBC"]=Australia["Total Gross Plus Total ER Contribution"]
        
        Australia.rename(columns={"Date of Joining":"Join Date"},inplace=True)
        
        aust_ex_rate=pd.read_excel("Rate Sheet.xlsx",sheet_name="Australia")
        aust_ex_rate.drop(columns=["Local Currency"],inplace=True)
        
        Australia=Australia.merge(aust_ex_rate,on="Invoicing Currency",how="left")

        # =============================================================================
        # Generating Multiple Excel File based on Raw Data Basis Division for Australia
        # =============================================================================
        
        for i in range(len(Australia["Division"].unique())):
            division_name=Australia["Division"].unique()[i]
            new_dataframe=Australia[Australia['Division']==division_name]
            new_dataframe.to_excel(excel_raw_data+"\\"+division_name.replace("/","")+".xlsx",index=False)  
        
        return Australia,aus_col    

    except FileNotFoundError:
        return None, print("Australia File or column mapping not found")


def is_valid_dataframe(df):
    return isinstance(df, pd.DataFrame) and not df.empty


def service_fee_divide_exchange_var_rate(master):
    if master["Local Currency"]=="AUD":
        USD=1.00
    elif (master["Local Currency"]=="INR") and (master["Invoicing Currency"]=="INR"):
        USD=1.00
    else:
        USD=round((master["EXCHANGE RATE"]/(1+master["Exchange Var"]*100/100)),2)
    return USD


def master_combine(dataframes):
    master=pd.concat(dataframes,axis=0)
    return master


def my_python():
    customer_sheet_auto()
    India,india_col=india_process_data()
    Bangladesh,bang_col=bangladesh_process_data()
    Australia,aus_col=australia_process_data()
    
    desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive - PROCLOZ SERVICES PVT LTD\Desktop\iBPRO\Payroll Invoices")
    os.chdir(desktop_path)
    input_path=desktop_path
    
    starting_invoice_number=start_invoice_num()

    master=pd.DataFrame()
    dataframes = [df for df in [India, Bangladesh, Australia] if is_valid_dataframe(df)]
    
    master = master_combine(dataframes)
    
    #master=pd.concat([India,Bangladesh,Australia],axis=0)
    print(master.columns)
    print(master['Division'])
    master['Division'] = master['Division'].apply(lambda x:"--" if pd.isnull(x) else x)

    inv_data=pd.DataFrame(master["Cost Center"]+"R0R0R0"+master["Division"]+"R0R0R0"+master["Bill Date"].astype(str)+"R0R0R0"+master["FY"],columns=["Unique"])
    inv_data=pd.DataFrame(inv_data["Unique"].unique(),columns=["Unique"])
    inv_data[["Cost Center","Division","Bill Date","FY"]]=inv_data["Unique"].str.split("R0R0R0",expand=True)
    inv_data["invoice_text"]="TI"
    inv_data["Bill Date"]=pd.to_datetime(inv_data["Bill Date"])
    inv_data['Month_name']=inv_data['Bill Date'].apply(lambda x:calendar.month_abbr[x.month])
    #print(calendar.month_abbr[3])

    inv_data.reset_index(inplace=True)
    #inv_data.drop("index",axis=1,inplace=True)
    inv_data.rename(columns={"index":"Sr No"},inplace=True)
    inv_data["Sr No"]=inv_data["Sr No"]+starting_invoice_number

    inv_data["Sr No"]=inv_data["Sr No"].apply(lambda x:str(x)[1:])

    #country_name_invoice_number=input("Enter the country name for invoice number: India, Bangladesh, Australia\n")
    country_name_invoice_number="Australia"

    if country_name_invoice_number=="India":
        inv_data["Invoice_number"]=inv_data["invoice_text"]+"/"+inv_data["FY"]+"/"+"IN"+"-"+inv_data["Sr No"]
    elif country_name_invoice_number=="Bangladesh":
        inv_data["Invoice_number"]=inv_data["invoice_text"]+"/"+inv_data["FY"]+"/"+"BD"+"-"+inv_data["Sr No"]
    elif country_name_invoice_number=="Australia":
        inv_data["Invoice_number"]=inv_data["invoice_text"]+"/"+inv_data["FY"]+"/"+"AU"+"-"+inv_data["Sr No"]

    # =============================================================================

    inv_data.reset_index(inplace=True)
    #inv_data.to_excel("inv_data.xlsx")
    #inv_data

    master["Unique"]=master["Cost Center"]+"R0R0R0"+master["Division"]+"R0R0R0"+master["Bill Date"].astype(str)+"R0R0R0"+master["FY"]
    inv=inv_data[["Unique","Invoice_number"]]

    # =============================================================================
    # Final touch up to master database
    # =============================================================================

    master=master.merge(inv,on="Unique",how="left")
    #master.drop("Invoice_number_y",axis=1,inplace=True)
    #master.rename(columns={"Invoice_number_x":"Invoice_number"},inplace=True)
    master.reset_index(inplace=True)
    #master.to_excel("master.xlsx")
    #writing salary file below to this section

    # =============================================================================
    # combining all the column mapping sheet in one master mapping sheet
    # =============================================================================

    all_col=pd.DataFrame()
    col_dataframes=[df for df in [india_col,aus_col,bang_col] if is_valid_dataframe(df)]
    all_col=master_combine(col_dataframes)
    #col_dataframes=[[globals().get('aus_col'),globals().get('bang_col'),globals().get('india_col')]]
    #all_col=pd.concat([aus_col,bang_col,india_col],axis=0)

    # for j in col_dataframes:
    #     try:
    #         if is_valid_dataframe(j):
    #             all_col=pd.concat([all_col,j],axis=0)
    #     except Exception as e:
    #         print(f"error occured while comining mapping col {e}")


    #all_col.to_excel('all columns.xlsx')

    # filling zero to blank rows in master file

    master=master.fillna(0)

    # =============================================================================
    # Lets create salary components for invoices
    # =============================================================================

    #creating Salary Column mapping, clubbed all other to salary

    sal=list(set(all_col[(all_col['Comment']=="ProEmp:Revenue - Salary External") | (all_col['Comment']=="PROEMP:Salary External")]['Column Name']))

    #len(sal)
    master['Salary']=0
    for col in sal:
        if col in master.columns:
            master['Salary']=master['Salary']+master[col]

    #master['Salary']=master[sal].sum(axis=1)
    # master.to_excel("testing.xlsx")
    #creating Salary other column mapping, to be added separately

    sal_other1=list(set(all_col[(all_col['Comment']=="PROEMP:Salary Others External") | (all_col['Comment']=="ProEmp:Revenue - Salary Others External")]['Column Name']))

    sal_other=[]
    for col in sal_other1:
        if col in master.columns:
            sal_other.append(col)
    #len(sal_other)

    #creating Bonus/Commission column mapping, to be called separately

    Bonus_com1=list(set(all_col[(all_col['Comment']=="PROEMP:Bonus/Commission External") | (all_col['Comment']=="ProEmp:Revenue - Bonus/Commission External")]['Column Name']))

    Bonus_com=[]
    for col in Bonus_com1:
        if col in master.columns:
            Bonus_com.append(col)
    #len(Bonus_com)

    #creating expense column to be added separately

    Expense1=list(set(all_col[(all_col['Comment']=="PROEMP:Expense Reimbursement External") | (all_col['Comment']=="ProEmp:Revenue - Expense Reimbursement External")]['Column Name']))

    Expense=[]
    for col in Expense1:
        if col in master.columns:
            Expense.append(col)

    #len(Expense)

    master['Expense_Excluded']=0
    for col in Expense:
        if col in master.columns:
            master['Expense_Excluded']=master['Expense_Excluded']+master[col]

    #master["Expense_Excluded"]=master[Expense].sum(axis=1)
    #master.to_excel("master.xlsx",index=False)

    # Creating social cost mapping column for invoices, to be added separately

    Social1=list(set(all_col[(all_col['Comment']=="ProEmp:Revenue - Social External") | (all_col['Comment']=="PROEMP:Employer Superannuation External (PF)")]['Column Name']))

    Social=[]
    for col in Social1:
        if col in master.columns:
            Social.append(col)

    #len(Social)

    #calculating Payroll Tax Rate Logic for Master
    payroll_tax=list(set(all_col[all_col["Comment"]=="PROEMP:Employer Payroll Tax External"]['Column Name']))
    payroll_tax_amt=[]
    for col in payroll_tax:
        if col in master.columns:
            payroll_tax_amt.append(col)

    #creating insurance cost mapping column for invoices, to be added separately as items

    insurance1=list(set(all_col[(all_col['Comment']=="ProEmp:Revenue - Insurance External") | (all_col['Comment']=="PROEMP:Workman Compensation External")] ["Column Name"]))


    insurance=[]
    for col in insurance1:
        if col in master.columns:
            insurance.append(col)

    #len(insurance)

    #creating Independent Contractor Charges i.e. IC Charges nominal code 

    independent_chr=list(set(all_col[all_col['Comment']=="ProEmp:COGS - IC Charges"]["Column Name"]))

    inde_charge=[]
    for col in independent_chr:
        if col in master.columns:
            inde_charge.append(col)
            
    #creating IT Procurement nominal code section here

    it_pro=list(set(all_col[all_col['Comment']=="ProEmp:Revenue - IT Procurement"]["Column Name"]))

    it_procurement=[]
    for col in it_pro:
        if col in master.columns:
            it_procurement.append(col)

    #creating adhoc cost mapping column for invoices, to be added separately as items

    adhoc1=list(set(all_col[all_col['Comment']=="ProEmp:Revenue - Adhoc"]["Column Name"]))

    adhoc=[]
    for col in adhoc1:
        if col in master.columns:
            adhoc.append(col)
    #len(adhoc)

    #creating this section for service fee
    service_fee=["Service Fees"]
    #len(service_fee)

    #creating this section for Set Up
    setup=["Set Up"]
    #len(setup)

    #creating total for service fee calculation master["Service_Fee_TBC"]

    service_fee_list=sal+sal_other1+Bonus_com1+Social1

    master["Service_Fee_TBC"]=0
    for col in service_fee_list:
        if col in master.columns:
            master["Service_Fee_TBC"]=master["Service_Fee_TBC"]+master[col]

    # =============================================================================
    # salary invoices component created in the above section 
    # =============================================================================

    # =============================================================================
    # reading Customer Contact Format and adding  state code to it
    # =============================================================================

    cf=pd.read_excel("Customer Contact Format.xlsx")
    cf["Customer"]=cf["Customer"].apply(lambda x:x.replace("Papaya Global, Inc. (External)","Papaya Global, Inc.  "))
    cf["Customer"]=cf["Customer"].apply(lambda x:x.replace("Papaya Global (HK) Limited (External)","Papaya Global (HK) Limited   "))
    cf["Customer"]=cf["Customer"].apply(lambda x:x.replace("Papaya Global, Inc. (Internal)","Papaya Global, Inc.    "))
    cf["Customer"]=cf["Customer"].apply(lambda x:x.replace("Papaya Global (HK) Limited (Internal)","Papaya Global (HK) Limited     "))

    cf["Billing State"]=cf["Billing State"].apply(lambda x:"--" if pd.isnull(x) else x)
    cf["Billing State"]=cf["Billing State"].apply(lambda x:x.lower())
    st=pd.read_excel("state_code.xlsx")
    st.rename(columns={"State":"Billing State"},inplace=True)
    st["Billing State"]=st["Billing State"].apply(lambda x:x.lower())

    customer_format=cf.merge(st,on="Billing State",how="left")
    customer_format.drop("State Code",axis=1,inplace=True)
    customer_format.rename(columns={"TIN":"State_Code"},inplace=True)

    customer_format["Billing State"]=customer_format["Billing State"].apply(lambda x:x.title())
    customer_format["State_Code"]=customer_format["State_Code"].fillna(0)
    customer_format["State_Code"]=customer_format["State_Code"].astype(int)
    customer_format["Terms"]=customer_format["Terms"].apply(lambda x:"Net 0" if pd.isnull(x) else x)
    customer_format.reset_index(inplace=True)

    customer_format["customer_location"]=customer_format["Customer"]+customer_format["Location"]
    customer_format["customer_location"]=customer_format["customer_location"].apply(lambda x:x.lower())

    #customer_format.to_excel("Customer_format.xlsx")


    # =============================================================================
    # merging customer format and master data for setup fees and service fee
    # =============================================================================

    customer=customer_format[["Bank Name","Customer","Billing Currency","Exchange Var","Billing State","Setup Fees","%age Markup","Minimum Fees","Location"]]
    customer.rename(columns={"%age Markup":"Service_fee_percentage","Minimum Fees":"Service_fee_minimu_fee"},inplace=True)
    customer["Setup Fees"]=customer["Setup Fees"].apply(lambda x:0 if pd.isnull(x) else str(x)[4:])
    customer["Setup Fees"]=customer["Setup Fees"].astype(int)
    customer["Service_fee_minimu_fee"]=customer["Service_fee_minimu_fee"].apply(lambda x: 0 if pd.isnull(x) else str(x)[4:])
    customer["Service_fee_minimu_fee"]=customer["Service_fee_minimu_fee"].astype(int)

    #customer.columns

    #customer.rename(columns={"Customer":"Cost Center"},inplace=True)
    customer["Customer1"]=customer["Customer"].apply(lambda x:x.lower())
    customer["Location1"]=customer["Location"].apply(lambda x:x.lower())
    customer["Unique_key_for_customer"]=customer["Customer1"]+customer["Location1"]
    customer.drop(["Customer1","Location1"],axis=1,inplace=True)

    master["Cost Center1"]=master["Cost Center"].apply(lambda x:x.lower())
    master["Country Name1"]=master["Country Name"].apply(lambda x:x.lower())
    master["Unique_key_for_customer"]=master["Cost Center1"]+master["Country Name1"]
    master.drop(["Cost Center1","Country Name1"],axis=1,inplace=True)

    #master.to_excel("testing.xlsx")

    master=master.merge(customer,on="Unique_key_for_customer",how="left")

    #place holder for master file to store exchange variance rate
    master["TBC_USD_Amt_for_Service_Fee"]=master.apply(service_fee_divide_exchange_var_rate,axis=1)

    #dividing exchange variance rate to SErvice Fee TBC column before computation

    master["Service_Fee_TBC"]=master["Service_Fee_TBC"]/master["TBC_USD_Amt_for_Service_Fee"]


    def service_fee_func(master):
        if (master["Service_Fee_TBC"]*master["Service_fee_percentage"])>master["Service_fee_minimu_fee"]:
            return (master["Service_Fee_TBC"]*master["Service_fee_percentage"])
        else:
            return master["Service_fee_minimu_fee"]

    master["Service Fees"]=master.apply(service_fee_func,axis=1)

    #fixing service fees to 12000 for TargetCW customer

    def service_targetcw_func(master):
        if master['Cost Center']=="TargetCW":
            return 12000/round((master["EXCHANGE RATE"]/(1+master["Exchange Var"]*100/100)),2)
        else:
            return master['Service Fees']
            

    master["Service Fees"]=master.apply(service_targetcw_func,axis=1)

    # =============================================================================
    #working on setup fee, 
    # created setup column which has setup fees
    # =============================================================================

    master["Join Date"]=pd.to_datetime(master["Join Date"])

    from datetime import datetime
    current_month=pd.to_datetime(datetime.now()).month
    current_year=pd.to_datetime(datetime.now()).year
    current_month_year=str(current_month)+str(current_year)

    def setup_fee(master):
        if str(master["Join Date"].month)+str(master["Join Date"].year)==current_month_year:
            return master["Setup Fees"]
        else:
            return 0

    master["Set Up"]=master.apply(setup_fee,axis=1)

    #writing master to xlsx file
    master=master.fillna(0)

    # =============================================================================
    # working with invoices
    # =============================================================================
    #state code will go blank for non Indian state
    # place for supply will be state code + Billing State, billing state and state code 
    #can be picked up from customer format file

    #single space for (external) and double space for (internal)

    #master["Cost Center"]=master["Cost Center"].apply(lambda x: x.replace("(External)","."))
    #master["Cost Center"]=master["Cost Center"].apply(lambda x: x.replace("(Internal)",".."))

    from docxtpl import DocxTemplate
    from docx2pdf import convert
    path=input_path+"\pdf invoices"
    #word_path=r"C:\Users\RahulTiwari\Documents\Rahul Tiwari\Python\Invoices Tool\Working files\word invoices"

    list_dict={}
    invoice_list=[]
    sub_dict={}
    total_inv_value={}
    charging_exch_rate={}
    bonus_comm={}
    exp_reimb={}
    contract_charges={}
    adhoc_charging={}
    capital_exp={}
    ins_sheet={}
    sal_xt={}
    sal_ot_xt={}
    ser_fee_sheet={}
    set_chg={}
    social_xt={}
    pr_tax={}
    for i in range(len(master)):
        for k in range(len(customer_format)):
            #(master["Cost Center"][i]==customer_format["Customer"][k]) # earlier this code was under if condition
            if (master["Unique_key_for_customer"][i]==customer_format["customer_location"][k]):
                cs=master["Cost Center"][i]
                #Printing bank details on invoices as modeling is taking time so printing the name
                #via loop here itself to test how much time it will take
                if (master["Local Currency"][i]=="AUD") and (master["Invoicing Currency"][i]=="AUD"):
                    beneficiary_name="Procloz Pty Ltd"
                    bank_name="NAB"
                    bank_address="Orion Shopping Centre, Shop 2B, 1 Main Street, Springfield Central, QLD, 4300"
                    bank_branch="--"
                    account_number="40-847-0970"
                    bic_swift_title_name="BSB Code: 084-740"
                    ifsc_code_title="IFSC Code: NATAAU3303M"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

                elif (master["Local Currency"][i]=="BDT"):
                    beneficiary_name="Procloz Private Limited"
                    bank_name="Standard Chartered"
                    bank_address="67 Gulshan Avenue, Gulshan, Dhaka 1212"
                    bank_branch="Gulshan"
                    account_number="01356078902"
                    bic_swift_title_name="Swift Code: SCBLBDDX"
                    ifsc_code_title=""
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   
                    
                elif (master["Local Currency"][i]=="BDT") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135305001997"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""                 

                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="INR") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135305001997"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="EUR") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135306000052"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="CAD") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135306000062"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="USD") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135306000044"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="GBP") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135306000050"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="AUD") and (master["Bank Name"][i]=="ICICI"):
                    beneficiary_name="Procloz Services Private Limited"
                    bank_name="ICICI Bank Limited"
                    bank_address="--"
                    bank_branch="New Delhi - Pitampura, Delhi"
                    account_number="135306000051"
                    bic_swift_title_name="Swift Code: ICICINBB"
                    ifsc_code_title="IFSC Code: ICIC0001353"
                    #hiding with no text for SCB bank details for non SCB bank
                    SCB_corrospondant_bank_title=""
                    SCB_corroospondant_swift_title=""
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   

    #generating SCB bank details from here 
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="USD") and (master["Bank Name"][i]=="SCB"):
                    beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                    bank_name="Standard Chartered Bank, India"
                    bank_address="--"
                    bank_branch="DLF Cyber City Gurugram Branch"
                    account_number="53105127765"
                    bic_swift_title_name="BIC/Swift Code: SCBLINBBCON"
                    ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                    SCB_corrospondant_bank_title="Correspondent Bank Name: Standard Chartered Bank, New York"
                    SCB_corroospondant_swift_title="BIC/Swift CODE: SCBLUS33XXX"
                    scb_account_read="SCB India’s account with SCB US will be read as 3582088635001"
                    scb_fed="The FED ABA No of SCB US is : 026002561"
                    scb_chips="The CHIPS ABA No of SCB US is : 0256"
                    #local currency INR, invoicing currency AUD and bank SCB
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="AUD") and (master["Bank Name"][i]=="SCB"):
                    beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                    bank_name="Standard Chartered Bank, India"
                    bank_address="--"
                    bank_branch="DLF Cyber City Gurugram Branch"
                    account_number="53105127838"
                    bic_swift_title_name="BIC/Swift Code: SCBLINBBCON"
                    ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                    SCB_corrospondant_bank_title="Correspondent Bank Name: National Australia Bank, Melbourne"
                    SCB_corroospondant_swift_title="BIC/Swift CODE: NATAAU33XXX"
                    scb_account_read="SCB India’s account with SCB US will be read as 1803011894500"
                    scb_fed=""
                    scb_chips=""            
                    #local currency INR, invoicing currency CAD and bank SCB
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="CAD") and (master["Bank Name"][i]=="SCB"):
                    beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                    bank_name="Standard Chartered Bank, India"
                    bank_address="--"
                    bank_branch="DLF Cyber City Gurugram Branch"
                    account_number="53105127811"
                    bic_swift_title_name="BIC/Swift Code: SCBLINBBCON"
                    ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                    SCB_corrospondant_bank_title="Correspondent Bank Name: Canadian Imperial Bank of Commerce, Toronto"
                    SCB_corroospondant_swift_title="BIC/Swift CODE: CIBCCATTXXX"
                    scb_account_read="SCB India’s account with SCB US will be read as 1775413"
                    scb_fed=""
                    scb_chips=""      
                    #local currency INR, invoicing currency EUR and bank SCB
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="EUR") and (master["Bank Name"][i]=="SCB"):
                    beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                    bank_name="Standard Chartered Bank, India"
                    bank_address="--"
                    bank_branch="DLF Cyber City Gurugram Branch"
                    account_number="53105127781"
                    bic_swift_title_name="BIC/Swift Code: SCBLINBBCON"
                    ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                    SCB_corrospondant_bank_title="Correspondent Bank Name: Standard Chartered Bank GERMANY"
                    SCB_corroospondant_swift_title="BIC/Swift CODE: SCBLDEFX"
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""      
                    #local currency INR, invoicing currency GBP and bank SCB
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="GBP") and (master["Bank Name"][i]=="SCB"):
                    beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                    bank_name="Standard Chartered Bank, India"
                    bank_address="--"
                    bank_branch="DLF Cyber City Gurugram Branch"
                    account_number="53105127803"
                    bic_swift_title_name="BIC/Swift Code: SCBLINBBCON"
                    ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                    SCB_corrospondant_bank_title="Correspondent Bank Name: Standard Chartered Bank London"
                    SCB_corroospondant_swift_title="BIC/Swift CODE: SCBLGB2LXXX"
                    scb_account_read=""
                    scb_fed=""
                    scb_chips=""   
                    #local currency INR, invoicing currency SGD and bank SCB
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="SGD") and (master["Bank Name"][i]=="SCB"):
                    beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                    bank_name="Standard Chartered Bank, India"
                    bank_address="--"
                    bank_branch="DLF Cyber City Gurugram Branch"
                    account_number="53105127773"
                    bic_swift_title_name="BIC/Swift Code: SCBLINBBCON"
                    ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                    SCB_corrospondant_bank_title="Correspondent Bank Name: Standard Chartered Bank SINGAPORE"
                    SCB_corroospondant_swift_title="BIC/Swift CODE: SCBLSGSGXXX"
                    scb_account_read="SCB India’s account with SCB US will be read as 5100029346"
                    scb_fed=""
                    scb_chips=""   
                    #local currency INR, invoicing currency INR and bank SCB
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="INR") and (master["Bank Name"][i]=="SCB"):
                   beneficiary_name="PROCLOZ SERVICES PRIVATE LIMITED"
                   bank_name="Standard Chartered Bank, India"
                   bank_address="--"
                   bank_branch="EXPRESS TOWERS"
                   account_number="53105127757"
                   bic_swift_title_name="BIC/Swift Code: SCBL0036086"
                   ifsc_code_title="NOSTRO/ INTERMEDIARY/ CORRESPONDENT/ FOREIGN Bank Details"
                   SCB_corrospondant_bank_title="Correspondent Bank Name: "
                   SCB_corroospondant_swift_title="BIC/Swift CODE: "
                   scb_account_read=""
                   scb_fed=""
                   scb_chips=""   
                        
                ba=customer_format["Billing Address"][k]
                dv=master["Division"][i]
                inv=master["Invoice_number"][i]
                bill_curr=master["Invoicing Currency"][i]
                local_curr=master["Local Currency"][i]
                if master["Local Currency"][i]=="AUD":
                    USD=1.00
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="INR"):
                    USD=1.00
                else:
                    USD=round((master["EXCHANGE RATE"][i]/(1+master["Exchange Var"][i]*100/100)),2)
                        #round((master["EXCHANGE RATE"][i]-(master["EXCHANGE RATE"][i]*master["Exchange Var"][i])),2) 
                        #round((master["EXCHANGE RATE"][i]/(1+master["Exchange Var"][i]*100/100)),2)
                #Storing Invoice number and exchange variance rate
                if inv not in charging_exch_rate:
                    charging_exch_rate[inv]=USD
                else:
                    charging_exch_rate[inv]=USD
                net_term=customer_format["Terms"][k]
                due_term1=customer_format["Terms"][k].split(" ")[1]
                current_date=master["Bill Date"][i].strftime("%d/%m/%Y")
                due_term=master["Bill Date"][i]+timedelta(days=int(due_term1))
                due_term=due_term.strftime("%d/%m/%Y")
                #generating place of supply
                if (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="INR"):
                    state=str(customer_format["State_Code"][k])+" - "+master["Billing State"][i]
                    place_suppy=""
                    place_dsc=""
                    reverse_chage=""
                    st_code=customer_format["State_Code"][k]
                    state_code_title="State Code"
                    state_colon=":"
                    gst_address_name="GSTIN"
                    gst_colon=":"
                    gst_address_number=customer_format["GSTIN"][k]
                elif (master["Local Currency"][i]=="INR"):
                    #(master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]!="INR") | (master["Invoicing Currency"][i]!="AUD"):
                    place_suppy="Place of Supply : "
                    state="Other Territory"
                    place_dsc="Supply is meant for Export under letter of understaking without payment of IGST"
                    reverse_chage="Reverse Charge : No"
                    st_code=""
                    state_code_title=""
                    state_colon=""
                    gst_address_name=""
                    gst_colon=""
                    gst_address_number=""
                elif (master["Local Currency"][i]=="AUD"):
                    #(master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]!="INR") | (master["Invoicing Currency"][i]!="AUD"):
                    place_suppy=""
                    state=""
                    place_dsc=""
                    reverse_chage=""
                    st_code=""
                    state_code_title=""
                    state_colon=""
                    gst_address_name=""
                    gst_colon=""
                    gst_address_number=""   
                elif (master["Local Currency"][i]=="BDT"):
                    #(master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]!="INR") | (master["Invoicing Currency"][i]!="AUD"):
                    place_suppy=""
                    state=""
                    place_dsc=""
                    reverse_chage=""
                    st_code=""
                    state_code_title=""
                    state_colon=""
                    gst_address_name=""
                    gst_colon=""
                    gst_address_number=""                
                else:
                    place_suppy=""
                    state=""
                    place_dsc=""
                    reverse_chage=""
                    st_code=""
                    state_code_title=""
                    state_colon=""
                    gst_address_name=""
                    gst_colon=""
                    gst_address_number=""
                #generating Procloz name and address for different entity
                if master["Country Name"][i]=="India":
                    pro_name="Procloz Services Private Limited"
                    pro_Add="7th Floor, Ambience Mall, National Highway-8 Gurgaon, Haryana  122002 IN"
                    gst_name="GSTIN"
                    gst_no="06AAECI2902E1Z3"
                    cin_name="CIN"
                    dot_bracket=":"
                    cin_num="U93090DL2016PTC305912"
                elif master["Country Name"][i]=="Australia":
                    pro_name="Procloz Pty Ltd"
                    pro_Add="161 Lakeside Avenue Springfield Lakes QLD 4300 +61 733819754"
                    gst_name="ABN"
                    gst_no="41641543008"
                    cin_name=""
                    dot_bracket=""
                    cin_num=""
                elif master["Country Name"][i]=="Bangladesh":
                    pro_name="Procloz Private Limited"
                    pro_Add="JCX Buisness Tower 1136/A, Japan Street, Level-5, Suite-H Block-I, Vatara, Bashunshara Dhaka 1212, Bangladesh"
                    gst_name="BIN"
                    gst_no="004557327-0101"
                    cin_name=""
                    dot_bracket=""
                    cin_num=""
                gst=customer_format["GSTIN"][k]
                emp_name=master["Name"][i]
                
                if master["Salary"][i]==0:
                    columnSeriesObj=master["Salary"]
                    pass
                else:
                    columnSeriesObj=master["Salary"]
                    # Generating this seciton for Salary,
                    #Createed a new list with serial number 1, as it could be first item for the list
                    if inv not in list_dict:
                        list_dict[inv] = [[1, emp_name + " " + columnSeriesObj.name, "{:,.2f}".format(master["Salary"][i]), "{:,.2f}".format(USD), "{:,.2f}".format(master["Salary"][i]/float(USD))]]
                    else:
                        # Get the last serial number used for this invoice number and add 1
                        last_counter = list_dict[inv][-1][0]
                        counter = last_counter + 1
                        description = emp_name + " " + columnSeriesObj.name
                        amount = master["Salary"][i]
                        exchange_rate = USD
                        final_amount = round(amount/float(USD), 2)
                        lst = [counter, description, "{:,.2f}".format(amount), "{:,.2f}".format(exchange_rate), "{:,.2f}".format(final_amount)]
                        list_dict[inv].append(lst)
                    if inv not in sub_dict:
                        sub_dict[inv]=round(master["Salary"][i]/float(USD), 2)
                    else:
                        sub_dict[inv]+=(round(master["Salary"][i]/float(USD), 2))
                    # writing salary external to add in charging sheet   
                    if inv not in sal_xt:
                        sal_xt[inv]=round(master["Salary"][i]/float(USD), 2)
                    else:
                        sal_xt[inv]+=round(master["Salary"][i]/float(USD), 2)                   
                        
                for s_ot in range(len(sal_other)):
                    columnSeriesObj=master[sal_other[s_ot]]
                    if master[sal_other[s_ot]][i]==0:
                        pass
                    else:
                        #generating this section for Salary other
                        # Createed a new list with serial number 1, as it could be first item for the list
                        if inv not in list_dict:
                            list_dict[inv] = [[1, emp_name + " " + columnSeriesObj.name, "{:,.2f}".format(master[sal_other[s_ot]][i]), "{:,.2f}".format(USD), "{:,.2f}".format(master[sal_other[s_ot]][i]/float(USD))]]
                        else:
                            # Get the last serial number used for this invoice number and add 1
                            last_counter = list_dict[inv][-1][0]
                            counter = last_counter + 1
                            columnSeriesObj=master[sal_other[s_ot]]
                            description=emp_name+" "+columnSeriesObj.name
                            sal_other_amt=master[sal_other[s_ot]][i]
                            sal_ot_final_amount=round(sal_other_amt/float(USD),2)
                            sal_ot_lst=[counter,description,"{:,.2f}".format(sal_other_amt),"{:,.2f}".format(USD),"{:,.2f}".format(sal_ot_final_amount)]
                            list_dict[inv].append(sal_ot_lst)
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[sal_other[s_ot]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[sal_other[s_ot]][i]/float(USD),2)
                         # writing Salary Others External to add in charging sheet   
                        if inv not in sal_ot_xt:
                            sal_ot_xt[inv]=round(master[sal_other[s_ot]][i]/float(USD),2)
                        else:
                            sal_ot_xt[inv]+=round(master[sal_other[s_ot]][i]/float(USD),2)         
                #generating this seciton for Bonus/Commission
                for bonus in range(len(Bonus_com)):
                    columnSeriesObj=master[Bonus_com[bonus]]
                    if master[Bonus_com[bonus]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[Bonus_com[bonus]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[Bonus_com[bonus]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[Bonus_com[bonus]]
                           description=emp_name+" "+columnSeriesObj.name
                           bonus_amt=master[Bonus_com[bonus]][i]
                           bonus_final_amount=round(bonus_amt/float(USD),2)
                           bonus_lst=[counter,description,"{:,.2f}".format(bonus_amt),"{:,.2f}".format(USD),"{:,.2f}".format(bonus_final_amount)]
                           list_dict[inv].append(bonus_lst)  
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[Bonus_com[bonus]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[Bonus_com[bonus]][i]/float(USD),2)
                         # writing bonus commission to add in charging sheet   
                        if inv not in bonus_comm:
                            bonus_comm[inv]=round(master[Bonus_com[bonus]][i]/float(USD),2)
                        else:
                            bonus_comm[inv]+=round(master[Bonus_com[bonus]][i]/float(USD),2)
                #generating this section for Expense
                for expense in range(len(Expense)):
                    columnSeriesObj=master[Expense[expense]]
                    if master[Expense[expense]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[Expense[expense]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[Expense[expense]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[Expense[expense]]
                           description=emp_name+" "+columnSeriesObj.name
                           expense_amt=master[Expense[expense]][i]
                           expense_final_amount=round(expense_amt/float(USD),2)
                           expense_lst=[counter,description,"{:,.2f}".format(expense_amt),"{:,.2f}".format(USD),"{:,.2f}".format(expense_final_amount)]
                           list_dict[inv].append(expense_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[Expense[expense]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[Expense[expense]][i]/float(USD),2)
                        # writing expense reimbursement to add in charging sheet   
                        if inv not in exp_reimb:
                            exp_reimb[inv]=round(master[Expense[expense]][i]/float(USD),2)
                        else:
                            exp_reimb[inv]+=round(master[Expense[expense]][i]/float(USD),2)
                # generating this seciton for Social
                for social in range(len(Social)):
                    columnSeriesObj=master[Social[social]]
                    if master[Social[social]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[Social[social]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[Social[social]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[Social[social]]
                           description=emp_name+" "+columnSeriesObj.name
                           social_amt=master[Social[social]][i]
                           social_final_amount=round(social_amt/float(USD),2)
                           social_lst=[counter,description,"{:,.2f}".format(social_amt),"{:,.2f}".format(USD),"{:,.2f}".format(social_final_amount)]
                           list_dict[inv].append(social_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[Social[social]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[Social[social]][i]/float(USD),2)
                        # writing Social External to add in charging sheet   
                        if inv not in social_xt:
                            social_xt[inv]=round(master[Social[social]][i]/float(USD),2)
                        else:
                            social_xt[inv]+=round(master[Social[social]][i]/float(USD),2)                            
                # generating this seciton for Payroll Tax Rate
                for pay_tax in range(len(payroll_tax)):
                    columnSeriesObj=master[payroll_tax[pay_tax]]
                    if master[payroll_tax[pay_tax]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name+tax_rate_f,"{:,.2f}".format(master[payroll_tax[pay_tax]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[payroll_tax[pay_tax]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[payroll_tax[pay_tax]]
                           tax_rate=round(master["Payroll Tax Rate"][i]*100,2)
                           tax_rate_f=" Rate "+str(tax_rate)+"%"                       
                           description=emp_name+" "+columnSeriesObj.name+tax_rate_f
                           payroll_tax_amt1=master[payroll_tax[pay_tax]][i]
                           payroll_tax_amt_final=round(payroll_tax_amt1/float(USD),2)
                           payroll_tax_lst=[counter,description,"{:,.2f}".format(payroll_tax_amt1),"{:,.2f}".format(USD),"{:,.2f}".format(payroll_tax_amt_final)]
                           list_dict[inv].append(payroll_tax_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[payroll_tax[pay_tax]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[payroll_tax[pay_tax]][i]/float(USD),2)
                        # writing Payroll tax to add in charging sheet   
                        if inv not in pr_tax:
                            pr_tax[inv]=round(master[payroll_tax[pay_tax]][i]/float(USD),2)
                        else:
                            pr_tax[inv]+=round(master[payroll_tax[pay_tax]][i]/float(USD),2)    
                #generating this section for Service Fee
                for serfee in range(len(service_fee)):
                    columnSeriesObj=master[service_fee[serfee]]
                    if master[service_fee[serfee]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[service_fee[serfee]][i]*USD),"{:,.2f}".format(USD),"{:,.2f}".format(master[service_fee[serfee]][i]*USD/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[service_fee[serfee]]
                           description=emp_name+" "+columnSeriesObj.name
                           service_amt=master[service_fee[serfee]][i]*USD
                           service_final_amount=round(service_amt/float(USD),2)
                           service_lst=[counter,description,"{:,.2f}".format(service_amt),"{:,.2f}".format(USD),"{:,.2f}".format(service_final_amount)]
                           list_dict[inv].append(service_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=master[service_fee[serfee]][i]
                        else:
                            sub_dict[inv]+=master[service_fee[serfee]][i]
                         # writing Service Fees Charging to add in charging sheet   
                        if inv not in ser_fee_sheet:
                            ser_fee_sheet[inv]=master[service_fee[serfee]][i]
                        else:
                            ser_fee_sheet[inv]+=master[service_fee[serfee]][i]                          
                #generating this section for SET UP Fee setup fees setup fee
                for set_up in range(len(setup)):
                    columnSeriesObj=master[setup[set_up]]
                    if master[setup[set_up]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[setup[set_up]][i]*USD),"{:,.2f}".format(USD),"{:,.2f}".format(master[setup[set_up]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[setup[set_up]]
                           description=emp_name+" "+columnSeriesObj.name
                           setup_amt=master[setup[set_up]][i]*USD
                           setup_final_amount=round(setup_amt/float(USD),2)
                           setup_lst=[counter,description,"{:,.2f}".format(setup_amt),"{:,.2f}".format(USD),"{:,.2f}".format(setup_final_amount)]
                           list_dict[inv].append(setup_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=master[setup[set_up]][i]
                        else:
                            sub_dict[inv]+=master[setup[set_up]][i]
                         # writing setup to add in charging sheet   
                        if inv not in set_chg:
                            set_chg[inv]=master[setup[set_up]][i]
                        else:
                            set_chg[inv]+=master[setup[set_up]][i]                        
                #generating this section for adhoc nominal charges adhoc charges adhoc fee
                for ad in range(len(adhoc)):
                    columnSeriesObj=master[adhoc[ad]]
                    if master[adhoc[ad]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[adhoc[ad]][i]*USD),"{:,.2f}".format(USD),"{:,.2f}".format(master[adhoc[ad]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[adhoc[ad]]
                           description=emp_name+" "+columnSeriesObj.name
                           adhoc_amt=master[adhoc[ad]][i]*USD
                           adhoc_final_amount=round(adhoc_amt/float(USD),2)
                           adhoc_lst=[counter,description,"{:,.2f}".format(adhoc_amt),"{:,.2f}".format(USD),"{:,.2f}".format(adhoc_final_amount)]
                           list_dict[inv].append(adhoc_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=master[adhoc[ad]][i]
                        else:
                            sub_dict[inv]+=master[adhoc[ad]][i]
                        # writing Adhoc Charges to add adhoc amount in charging sheet   
                        if inv not in adhoc_charging:
                            adhoc_charging[inv]=master[adhoc[ad]][i]
                        else:
                            adhoc_charging[inv]+=master[adhoc[ad]][i]         
                            
                #generating this section for insurance
                for ins in range(len(insurance)):
                    columnSeriesObj=master[insurance[ins]]
                    if master[insurance[ins]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[insurance[ins]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[insurance[ins]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[insurance[ins]]
                           description=emp_name+" "+columnSeriesObj.name
                           insurance_amt=master[insurance[ins]][i]
                           insurance_final_amount=round(insurance_amt/float(USD),2)
                           insurance_lst=[counter,description,"{:,.2f}".format(insurance_amt),"{:,.2f}".format(USD),"{:,.2f}".format(insurance_final_amount)]
                           list_dict[inv].append(insurance_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[insurance[ins]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[insurance[ins]][i]/float(USD),2)
                         # writing insurance external to add in charging sheet   
                        if inv not in ins_sheet:
                            ins_sheet[inv]=round(master[insurance[ins]][i]/float(USD),2)
                        else:
                            ins_sheet[inv]+=round(master[insurance[ins]][i]/float(USD),2)
                #generating this section for independent contractor charges
                for contchrge in range(len(inde_charge)):
                    columnSeriesObj=master[inde_charge[contchrge]]
                    if master[inde_charge[contchrge]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[inde_charge[contchrge]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[inde_charge[contchrge]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[inde_charge[contchrge]]
                           description=emp_name+" "+columnSeriesObj.name
                           indcharge_amt=master[inde_charge[contchrge]][i]
                           indcharge_amt_final=round(indcharge_amt/float(USD),2)
                           indcharge_lst=[counter,description,"{:,.2f}".format(indcharge_amt),"{:,.2f}".format(USD),"{:,.2f}".format(indcharge_amt_final)]
                           list_dict[inv].append(indcharge_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[inde_charge[contchrge]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[inde_charge[contchrge]][i]/float(USD),2)   
                        # writing Independent Contractor Charges to add in charging sheet   
                        if inv not in contract_charges:
                            contract_charges[inv]=round(master[inde_charge[contchrge]][i]/float(USD),2)
                        else:
                            contract_charges[inv]+=round(master[inde_charge[contchrge]][i]/float(USD),2)         
                #generating this section for IT Procurement Nomimal code (ITP)
                for itp in range(len(it_procurement)):
                    columnSeriesObj=master[it_procurement[itp]]
                    if master[it_procurement[itp]][i]==0:
                        pass
                    else:
                        if inv not in list_dict:
                            list_dict[inv]=[[1,emp_name+" "+columnSeriesObj.name,"{:,.2f}".format(master[it_procurement[itp]][i]),"{:,.2f}".format(USD),"{:,.2f}".format(master[it_procurement[itp]][i]/float(USD))]]
                        else:
                           last_counter = list_dict[inv][-1][0]
                           counter = last_counter + 1
                           columnSeriesObj=master[it_procurement[itp]]
                           description=emp_name+" "+columnSeriesObj.name
                           itp_amt=master[it_procurement[itp]][i]
                           itp_amt_final=round(itp_amt/float(USD),2)
                           itp_lst=[counter,description,"{:,.2f}".format(itp_amt),"{:,.2f}".format(USD),"{:,.2f}".format(itp_amt_final)]
                           list_dict[inv].append(itp_lst) 
                        if inv not in sub_dict:
                            sub_dict[inv]=round(master[it_procurement[itp]][i]/float(USD),2)
                        else:
                            sub_dict[inv]+=round(master[it_procurement[itp]][i]/float(USD),2)   
                         # writing Capital Expenditure External to add in charging sheet   
                        if inv not in capital_exp:
                            capital_exp[inv]=round(master[it_procurement[itp]][i]/float(USD),2)
                        else:
                            capital_exp[inv]+=round(master[it_procurement[itp]][i]/float(USD),2)
                                                    
                #generating IGST VAT SGST CGST calculation section for invoice "Haryana".lower()
                if (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="INR") and (master["Billing State"][i].lower()=="haryana"):
                    tax_name="CGST"
                    at_the_rate="@"
                    tax_percentage="9%"
                    tax_on="on"
                    tax_subtotal="{:,.2f}".format(sub_dict[inv])
                    taxed_amt=sub_dict[inv]*9/100
                    taxed_Amount="{:,.2f}".format(taxed_amt)
                    state_tax_name="SGST"
                    state_rate_at="@"
                    state_percentage="9%"
                    state_on="on"
                    state_subb_total="{:,.2f}".format(sub_dict[inv])
                    state_taxed_amount=sub_dict[inv]*9/100
                    state_taxed_total="{:,.2f}".format(state_taxed_amount)
                    totaldue="{:,.2f}".format((taxed_amt+state_taxed_amount+sub_dict[inv])) 
                    #adding total invoice value to dictionary
                    if inv not in total_inv_value:
                        total_inv_value[inv]=totaldue
                    else:
                        total_inv_value[inv]=totaldue
                elif (master["Local Currency"][i]=="INR") and (master["Invoicing Currency"][i]=="INR") and (master["Billing State"][i].lower()!="haryana"):
                    tax_name="IGST"
                    at_the_rate="@"
                    tax_percentage="18%"
                    tax_on="on"
                    tax_subtotal="{:,.2f}".format(sub_dict[inv])
                    taxed_amt=sub_dict[inv]*18/100
                    taxed_Amount="{:,.2f}".format(taxed_amt)
                   #making state tax zero for other section which is not applicable to print in the invoice
                    state_tax_name=""
                    state_rate_at=""
                    state_percentage=""
                    state_on=""
                    state_subb_total=""
                    state_taxed_amount=""
                    state_taxed_total=""
                    #totaldue="{:,}".format(round((taxed_amt+state_taxed_amount+sub_dict[inv]),2))
                    totaldue="{:,.2f}".format((taxed_amt+sub_dict[inv]))   
                    if inv not in total_inv_value:
                        total_inv_value[inv]=totaldue
                    else:
                        total_inv_value[inv]=totaldue
                elif (master["Local Currency"][i]=="INR"):
                    tax_name="IGST"
                    at_the_rate="@"
                    tax_percentage="0%"
                    tax_on="on"
                    tax_subtotal="{:,.2f}".format(sub_dict[inv])
                    taxed_Amount=(sub_dict[inv]*0/100)
                   #making state tax zero for other section which is not applicable to print in the invoice
                    state_tax_name=""
                    state_rate_at=""
                    state_percentage=""
                    state_on=""
                    state_subb_total=""
                    state_taxed_amount=""
                    state_taxed_total=""
                    #totaldue="{:,}".format(round((taxed_amt+state_taxed_amount+sub_dict[inv]),2))                
                    totaldue="{:,.2f}".format((taxed_Amount+sub_dict[inv]))
                    if inv not in total_inv_value:
                        total_inv_value[inv]=totaldue
                    else:
                        total_inv_value[inv]=totaldue
                elif (master["Local Currency"][i]=="BDT"):
                    tax_name="VAT"
                    at_the_rate="@"
                    tax_percentage="0%"
                    tax_on="on"
                    tax_subtotal="{:,.2f}".format(sub_dict[inv])
                    taxed_Amount=sub_dict[inv]*0/100
                   #making state tax zero for other section which is not applicable to print in the invoice
                    state_tax_name=""
                    state_rate_at=""
                    state_percentage=""
                    state_on=""
                    state_subb_total=""
                    state_taxed_amount=""
                    state_taxed_total=""
                    #totaldue="{:,}".format(round((taxed_amt+state_taxed_amount+sub_dict[inv]),2))                
                    totaldue="{:,.2f}".format((taxed_Amount+sub_dict[inv]))
                    if inv not in total_inv_value:
                        total_inv_value[inv]=totaldue
                    else:
                        total_inv_value[inv]=totaldue
                elif (master["Local Currency"][i]=="AUD"):
                    tax_name="GST"
                    at_the_rate="@"
                    tax_percentage="0%"
                    tax_on="on"
                    tax_subtotal="{:,.2f}".format(sub_dict[inv])
                    taxed_Amount=sub_dict[inv]*0/100
                   #making state tax zero for other section which is not applicable to print in the invoice
                    state_tax_name=""
                    state_rate_at=""
                    state_percentage=""
                    state_on=""
                    state_subb_total=""
                    state_taxed_amount=""
                    state_taxed_total=""
                    #totaldue="{:,}".format(round((taxed_amt+state_taxed_amount+sub_dict[inv]),2))                
                    totaldue="{:,.2f}".format(sub_dict[inv])  
                    if inv not in total_inv_value:
                        total_inv_value[inv]=totaldue
                    else:
                        total_inv_value[inv]=totaldue

                subtotal="{:,.2f}".format(sub_dict[inv]) 
                invoice_list=list_dict[inv]
                doc=DocxTemplate("New_Foreign.docx")
                context={"PName":pro_name,
                         "PAdd":pro_Add,
                         "BIN_ABN_GST":gst_name,
                         "BIN_ABN_no":gst_no,
                         "CIN":cin_name,
                         "dot":dot_bracket,
                         "CIN_No":cin_num,
                         "CustomerName":cs,
                         "BillingAdd1":ba,
                         "State_code":st_code,
                         "code_name":state_code_title,
                         "dt":state_colon,
                         "gst":gst_address_name,
                         "coln":gst_colon,
                         "gst_no":gst_address_number,
                         "GST_No":gst,
                         "Invoice_no":inv,
                         "Inv_Date":current_date,
                         "Due_date":due_term,
                         "Term":net_term,
                         "Place":state,
                         "PLACE_SUPP":place_suppy,
                         "Place_desciption":place_dsc,
                         "reverse_charge":reverse_chage,
                         "Division":dv,
                         "USD":"{:,.2f}".format(USD),
                         "LCUR":local_curr,
                         "invoice_list":invoice_list,
                         "Subtotal":subtotal,
                         "BCUR":bill_curr,
                         "V_NM":tax_name,
                         "at":at_the_rate,
                         "P":tax_percentage,
                         "on":tax_on,
                         "tax_sub":tax_subtotal,
                         "VAT_Amt":taxed_Amount,
                         "S_NM":state_tax_name,
                         "taa":state_rate_at,
                         "PP":state_percentage,
                         "onn":state_on,
                         "tax_subb":state_subb_total,
                         "S_Amt":state_taxed_total,
                         "total_due":totaldue,
                         "Ben_name":beneficiary_name,
                         "Bank_name":bank_name,
                         "Bank_add":bank_address,
                         "Branch_name":bank_branch,
                         "Accountnumber":account_number,
                         "BIC_Swift_BSB_details":bic_swift_title_name,
                         "IFSC_Code_details_Nastro":ifsc_code_title,
                         "SCB_corrospondant_bank_details":SCB_corrospondant_bank_title,
                         "scb_corrospondant_swift_bic_details":SCB_corroospondant_swift_title,
                         "scb_india_read_as_title_account_number":scb_account_read,
                         "the_fed_ABA_number":scb_fed,
                         "the_chips_ABA_number":scb_chips
                         }
                doc.render(context)
                doc.save(excel_raw_data+"\\"+dv.replace("/","")+"_"+cs+".docx")
               # convert(dv+"_"+cs+".docx",path)

    # =============================================================================
    # Writing PDF files in the below code
    # =============================================================================
    #os.getcwd()
    os.chdir(excel_raw_data)
    import glob
    #glob.glob("*.docx")
    for file in glob.glob("*.docx"):
        convert(file,path)

    os.chdir(input_path)

    # =============================================================================
    # Generating the Charging Sheet from here
    # =============================================================================

    #generating total_invoice value or total sale value
    total_invoicing_amount=pd.DataFrame.from_dict(total_inv_value,orient='index',columns=["Total Invoice Value"])
    total_invoicing_amount.reset_index(inplace=True)
    total_invoicing_amount.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(total_invoicing_amount,on="Invoice_number",how="left")
    #master[["Invoice_number","Total Invoice Value"]]

    #Social External code for charging sheet
    social_xtrnal=pd.DataFrame.from_dict(social_xt,orient="index",columns=["Social External"])
    social_xtrnal.reset_index(inplace=True)
    social_xtrnal.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(social_xtrnal,on="Invoice_number",how="left")

    #Payroll Tax Rate amount is getting added here and it will get add to Social External then
    #pr_tax

    payroll_tax_chrg=pd.DataFrame.from_dict(pr_tax,orient="index",columns=["Payroll_Tax_Rate"])
    payroll_tax_chrg.reset_index(inplace=True)
    payroll_tax_chrg.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(payroll_tax_chrg,on="Invoice_number",how="left")

    master["Payroll_Tax_Rate"]=master["Payroll_Tax_Rate"].apply(lambda x: 0 if pd.isnull(x) else x)

    #master["Social External"]=master["Social External"]+master["Payroll_Tax_Rate_chg"]

    #Setup code for charging sheet
    setup_charge=pd.DataFrame.from_dict(set_chg,orient="index",columns=["Setup Charging"])
    setup_charge.reset_index(inplace=True)
    setup_charge.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(setup_charge,on="Invoice_number",how="left")

    #Service Fees Charging code for Charging Sheet
    serice_fee_sheet=pd.DataFrame.from_dict(ser_fee_sheet,orient="index",columns=["Service Fees Charging"])
    serice_fee_sheet.reset_index(inplace=True)
    serice_fee_sheet.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(serice_fee_sheet,on="Invoice_number",how="left")

    #Salary Other External code for charging sheet
    sal_ot_xtrnal=pd.DataFrame.from_dict(sal_ot_xt,orient="index",columns=["Salary Others External"])
    sal_ot_xtrnal.reset_index(inplace=True)
    sal_ot_xtrnal.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(sal_ot_xtrnal,on="Invoice_number",how="left")

    #Salary External code for charging sheet
    sal_xtrnal=pd.DataFrame.from_dict(sal_xt,orient="index",columns=["Salary External"])
    sal_xtrnal.reset_index(inplace=True)
    sal_xtrnal.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(sal_xtrnal,on="Invoice_number",how="left")

    #Insurance External code for Charging sheet
    insurance_exter=pd.DataFrame.from_dict(ins_sheet,orient="index",columns=["Insurance External"])
    insurance_exter.reset_index(inplace=True)
    insurance_exter.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(insurance_exter,on="Invoice_number",how="left")

    #Capital Expenditure External code implementation for Charging Sheet
    capital_expenditure=pd.DataFrame.from_dict(capital_exp,orient='index',columns=["Capital Expenditure External"])
    capital_expenditure.reset_index(inplace=True)
    capital_expenditure.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(capital_expenditure,on="Invoice_number",how="left")

    #client adhoc charges code for sharging sheet
    client_adhoc_charging=pd.DataFrame.from_dict(adhoc_charging,orient="index",columns=["Client Adhoc"])
    client_adhoc_charging.reset_index(inplace=True)
    client_adhoc_charging.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(client_adhoc_charging,on="Invoice_number",how="left")

    #code for Independent contractor charges for charging sheet

    indpendent_con_charg=pd.DataFrame.from_dict(contract_charges,orient='index',columns=["Independent Contractor Charges"])
    indpendent_con_charg.reset_index(inplace=True)
    indpendent_con_charg.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(indpendent_con_charg,on="Invoice_number",how="left")

    #code for Expense Reimbursement to extract the details for charging sheet
    expense_reim=pd.DataFrame.from_dict(exp_reimb,orient='index',columns=["Expense Reimbursement External"])
    expense_reim.reset_index(inplace=True)
    expense_reim.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(expense_reim,on="Invoice_number",how="left")

    # code for Bonus Commission to extract the details for charging sheet
    chg_bonus=pd.DataFrame.from_dict(bonus_comm,orient='index',columns=["Bonus/Commission External"])
    chg_bonus.reset_index(inplace=True)
    chg_bonus.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(chg_bonus,on="Invoice_number",how="left")

    #generating exchange variance rate for charging sheet
    exchange_rate_charging=pd.DataFrame.from_dict(charging_exch_rate,orient='index',columns=["Billing Exchnage Rate"])
    exchange_rate_charging.reset_index(inplace=True)
    exchange_rate_charging.rename(columns={"index":"Invoice_number"},inplace=True)

    master=master.merge(exchange_rate_charging,on="Invoice_number",how="left")
    #generating billing currency
    master["Billing Currency"]=master["Invoicing Currency"]
    #generating Exchnage Rate
    master["Exchnage Rate"]="1 "+master["Billing Currency"]+" = "+master["Billing Exchnage Rate"].apply(str)+" "+master["Local Currency"]

    #create actual month column
    master["Client Deposit"]="" # creating blank column per Gaurav
    master["Others External"]=""
    master["Others Discription"]=""
    #creating following blank column in master list for charing sheet
    master["Funds Received on"]=""
    master["Amount Received (In forrign currency)"]=""
    master["Balance"]=""
    master["Comments"]=""
    master["FIRC No."]=""
    master["Funds Paid on"]=""
    master["Amount Paid"]=""
    master["Balance_supplier"]=""


    #creating following blank column for main invoices however for Partner data to be pulled from GTN
    master["Supplier Deposit"]=""
    master["Supplier Adhoc Cost"]=""
    master["Supplier Service Cost"]=""
    master["Supplier Setup Cost"]=""

    #master["Supplier Name"]=""
    master["Supplier Invoice #"]=""

    #master[["Invoice_number","Total Invoice Value"]]

    #Generating the #of Employees Data here 

    number_empl=pd.DataFrame(master.groupby('Invoice_number')["Invoice_number"].count())
    number_empl.rename(columns={"Invoice_number":"# Employees"},inplace=True)
    number_empl.reset_index(inplace=True)

    master=master.merge(number_empl,on="Invoice_number",how="left")


    #Calculating or Generating Supplier Total Cost and Profit
     
    master["Supplier Total (Cost)"]=master[["Independent Contractor Charges","Bonus/Commission External","Capital Expenditure External","Expense Reimbursement External","Insurance External","Others External","Salary External",
                                            "Salary Others External","Social External","Supplier Deposit","Supplier Adhoc Cost","Supplier Service Cost","Supplier Setup Cost"]].sum(axis=1)

    master["Total Invoice Value"]=master["Total Invoice Value"].apply(lambda x:str(x).replace(",",""))


    master["Total Invoice Value"]=pd.to_numeric(master["Total Invoice Value"],errors='coerce').astype(float)
    master["Profit"]=master["Total Invoice Value"]-master["Supplier Total (Cost)"]
    #writing master data to excel file
    #master.to_excel("master.xlsx",index=False)

    india_master=master[["Entity Name","Bill Date","Actual Month","Invoice_number","Cost Center",
                         "Division","Country Name","# Employees","Type of Invoice","Invoice/Credit","Billing Currency",
                         "Exchnage Rate","Billing Exchnage Rate","Independent Contractor Charges",
                         "Client Deposit","Client Adhoc","Bonus/Commission External","Capital Expenditure External",
                         "Expense Reimbursement External","Insurance External","Others External","Others Discription",
                         "Salary External","Salary Others External","Service Fees Charging","Setup Charging",
                         "Social External","Payroll_Tax_Rate","Total Invoice Value","Funds Received on","Amount Received (In forrign currency)",
                         "Balance","Comments","FIRC No.","Supplier Deposit","Supplier Adhoc Cost","Supplier Service Cost",
                         "Supplier Setup Cost","Supplier Total (Cost)","Funds Paid on","Amount Paid","Balance_supplier",
                         "Supplier Name","Supplier Invoice #","Profit"]]

    india_master.rename(columns={"Bill Date":"Date"},inplace=True)

    india_master=india_master.drop_duplicates()
    final_india=pd.read_excel("Final_India_data.xlsx")
    final_india_data=pd.concat([india_master,final_india])
    final_india_data.to_excel("Final_India_data.xlsx",index=False)

@app.route('/result', methods=['GET'])

def get_data():
    try:
        pythoncom.CoInitialize()
        result=my_python()     
        pythoncom.CoUninitialize()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run()
    
