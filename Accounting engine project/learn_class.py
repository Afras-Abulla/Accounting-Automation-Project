import pandas as pd
import openpyxl

 
#___NOTE THIS BEFORE USE__
# the Accounting object takes a total of 4 arguments as of now which are 
# journal,mapped trial balance, reconcilation data 1 and reconcilation data 2
# every input should be a pandas dataframe
# for the tb methode it takes 3 arguments they are index numbers of the particulars column, debit column and credit column simply enter the numbers and it will generate trial balance (it wont work with ledger as of now only from journal directly to trial balance)
# income statement methode requires no arguments but you should create an mapped trial balance attribute for income statement and balance sheet to work and the mapping should be in a specific order and syntax-
# it should have columns MAPPING1 MAPPING2 MAPPING 3 AND ADJUSTED CLOSING BALANCE in this exact spelling and capitalization
# and in MAPPING1 it should be catogorized into SOPL AND SOFP also in the same spelling and cap
# in MAPPING2 it should be catogoriszed into INCOME EXPENSE CURRENT ASSET, NON CURRENT ASSET, CURRENT LIABILITY and NON CURRENT LIABILITY in the exact spelling and cap
# in mapping 3 you can further classify to your will
# and for reconcilation you create attributes recon data1 and recon data2 after that for reconcilation methode
# you can pass in the index number of first key identifier column then second key identifier column the the index of first value column and second value column first and second being the dataframes or data
# and for export_to_excel methode just pass the path to the excel file as an argument make sure the file is empty and if you want the reconcilation data too just enter True as the second argument 
# _NOTE: this is a project created by me to demonstrate my analetycal skill and it is not fully complete i am still working on it since this is a valuable demonstration of my skill i am showing this on my profile.
# this project is the result of me solving the problems that i faced during my audit internship using python freeing up my valuable time to be used for other value adding activities
# keep in mind i am a beginer programer who is still learning 
# THANK YOU       

class Accounting():

    def __init__(self,Journaldf = None ,Mappedtb = None,Rec1 = None, Rec2 = None):
            self.Journal = Journaldf
            self.Tb = None
            self.Income_statement = None
            self.Balance_sheet = None
            self.Mappedtb = Mappedtb.copy() if Mappedtb is not None else None
            self.Net_income = None
            self.Rec1 = Rec1
            self.Rec2 = Rec2
            self.Reconciledm = None
            self.Reconcilednm = None
        
        


    def colname(self,*indexes):
        cnames = []
        for i in indexes:
            cnames.append(self.Journal.columns[i])
        return cnames

    def tb(self,pindex,debindex,credindex):
        cnames = self.colname(pindex,debindex,credindex)
        td = self.Journal[cnames[1]].sum()
        tc = self.Journal[cnames[2]].sum()
        tbp = self.Journal.groupby(self.Journal[cnames[0]]).sum(numeric_only = True)
        tr = pd.DataFrame([{cnames[1]:td,cnames[2]:tc}], index=["Total"])
        tbp = pd.concat([tbp,tr])
        self.Tb = tbp
        return tbp

    def income_statement(self):
        intb = self.Mappedtb.groupby(["MAPPING1","MAPPING2","MAPPING3"])["ADJUSTED CLOSING"].sum()
        insopl = intb["SOPL"]
        inrev = abs(insopl["INCOME"])
        inexp = insopl["EXPENSE"] * -1 
        income = pd.concat([inrev,inexp])
        ni = income.sum()
        net_income = pd.DataFrame([{"ADJUSTED CLOSING": ni}], index=["NET INCOME"])
        income_state = pd.concat([income,net_income])
        self.Income_statement = income_state
        self.Net_income = ni
        return income_state
    
    def loc_net_income(self,incomestatement):
        net_income = incomestatement.loc["NET INCOME"]
        return net_income
    
    def balance_sheet(self):
        intb = self.Mappedtb.groupby(["MAPPING1","MAPPING2","MAPPING3"])["ADJUSTED CLOSING"].sum()
        insofp = intb["SOFP"]
        inca = insofp["CURRENT ASSET"]
        innca = insofp["NON CURRENT ASSET"]
        asset = pd.concat([inca,innca])
        ta = asset.sum()
        atotal = pd.DataFrame([{"ADJUSTED CLOSING" : ta}], index=["TOTAL ASSETS"])
        ineq = insofp["EQUITY"]
        incl = insofp["CURRENT LIABILITY"]
        try:
            inncl = insofp["NON CURRENT LIABILITY"]
            lande = pd.concat([ineq,incl,inncl])
            let = abs(lande.sum())
            letotal = pd.DataFrame([{"ADJUSTED CLOSING":let}], index=["TOTAL LIABILITY"])
            bs = pd.concat([asset,atotal,lande,letotal])
            self.Balance_sheet = bs
            return bs
        except KeyError:
            nl = pd.concat([ineq,incl])
            nlt = abs(nl.sum())
            adj = nlt + self.Net_income
            ntotal = pd.DataFrame([{"ADJUSTED CLOSING":adj}], index=["TOTAL LIABILITY"])
            liability = pd.concat([ineq,incl,ntotal])
            bs = pd.concat([asset,atotal,liability])
            self.Balance_sheet = bs
            return bs

    def reconcilation(self,uniqid1,uniqueid2,amtclm1,amtclm2):

        # geting collumn names 
        uid1 = self.Rec1.columns[uniqid1]
        uid2 = self.Rec2.columns[uniqueid2]
        amt1 = self.Rec1.columns[amtclm1]
        amt2 = self.Rec2.columns[amtclm2]

        #finding total variance
        td1 = self.Rec1[amt1].sum()
        td2 = self.Rec2[amt2].sum()
        td  = abs(td1 - td2)

        #find and create a mathing dataframe 
        matching = self.Rec1.merge(self.Rec2, on = uid1, how = "inner", suffixes = ("1st", "2nd")).copy()
        mcol = matching.columns
        if amt1+"1st" in mcol:
            mdf = matching[[uid1,amt1+"1st",amt2+"2nd"]]
        else:
            mdf = matching[[uid1,amt1,amt2]].copy()
        #find and create a not matching dataframe
        nmatchings = self.Rec1.merge(self.Rec2, on = uid1, how = "outer", indicator = True, suffixes = ("1st","2nd") ).copy()
        nmatching1 = nmatchings[nmatchings["_merge"] == "left_only"]
        nmatching2 = nmatchings[nmatchings["_merge"] == "right_only"]
        nmc = nmatchings.columns
        if amt1+"1st" in nmc:
            nmatchingm1 = nmatching1[[uid1,amt1+"1st",amt2+"2nd"]]
            nmatchingm2 = nmatching2[[uid1,amt1+"1st",amt2+"2nd"]]
            nmatching = pd.concat([nmatchingm1,nmatchingm2])
            nmatching.columns = ["omitted_items",amt1+"1st",amt2+"2nd"]
        #find difference of matching items
            amc1 = mdf[amt1+"1st"]
            amc2 = mdf[amt2+"2nd"]
        else:
            nmatchingm1 = nmatching1[[uid1,amt1,amt2]]
            nmatchingm2 = nmatching2[[uid1,amt1,amt2]]
            nmatching = pd.concat([nmatchingm1,nmatchingm2])
            nmatching.columns = ["omitted_items",amt1,amt2]
        #find difference of matching items
            amc1 = mdf[amt1]
            amc2 = mdf[amt2]

        mdf["difference"] =  abs(amc1 - amc2)
        self.Reconciledm = mdf
        self.reconcilednm = nmatching 
        return f"Total Variance {td}\n\n{mdf}\n\n{nmatching}"

    def export_to_excel(self,file_path, rec = False):
        with pd.ExcelWriter(file_path, engine = "openpyxl") as writer:

            if self.Tb is not None:
                self.Tb.to_excel(writer, sheet_name = "Trial Balance", index = True)

            if self.Income_statement is not None:
                self.Income_statement.to_excel(writer, sheet_name = "Income statement", index = True)

            if self.Balance_sheet is not None:
                self.Balance_sheet.to_excel(writer, sheet_name = "Balance sheet", index = True)


            if rec is not False:
                if self.Reconciledm is not None and self.reconcilednm is not None:
                    self.Reconciledm.to_excel(writer, sheet_name = "Reconciliation", startrow = 2, startcol = 1  , index = True)
                    self.Reconcilednm.to_excel(writer, sheet_name = "Reconcilation", startrow = 2, startcol = 6, index = True)



