# Import the necessary module
import pandas as pd
import numpy as np
import os
from PyQt5 import QtWidgets
from PyQt5 import QtGui
import chardet
import datetime
import shutil

class myApp(QtWidgets.QWizard):
    def __init__(self, parent=None):
        super(myApp, self).__init__(parent)

        """ Initialise Environment (If not yet initialised) """
        self.envInit()

        self.mainPage = MainPage()
        self.addPage(self.mainPage)

        self.setPixmap(QtWidgets.QWizard.BannerPixmap, QtGui.QPixmap(os.getcwd()+"/CASSCare.png"))

        self.setWindowTitle("Generate Summary Report")
        self.setWindowIcon(QtGui.QIcon(r'icon.ico'))
        self.setWizardStyle(1)

        self.setButtonText(self.FinishButton, "Run!")
        #self.button(QtWidgets.QWizard.FinishButton).clicked.connect(self.mainPage.saveFile)
        self.button(QtWidgets.QWizard.FinishButton).clicked.connect(self.runProgram)

        self.setOption(self.NoBackButtonOnStartPage, True)

        self.errorCode = 0
        #self.mismatchInfo = ""
        #self.missingIndex = ""
        self.dcwFilePath = ""
        self.payshtFilePath = ""
        self.outputFilePath = ""
        self.arcDirectory = ""

    """Predefined functions"""
    """envInit - initialise working directories"""
    def envInit(self):
        if not(os.path.isdir("DATA/current/input_files/")):
            os.makedirs("DATA/current/input_files",exist_ok=True)
        if not(os.path.isdir("DATA/current/output_file/")):
            os.makedirs("DATA/current/output_file",exist_ok=True)
        if not(os.path.isdir("DATA/archive/")):
            os.makedirs("DATA/archive",exist_ok=True)


    """errorMsg - Pop Up Window Box to show the error Message"""
    def errorMsg(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        if self.errorCode == 1:
            msg.setText("You have choose to run in Custom Mode, but you haven't choose any file or folders.")
        elif self.errorCode == 2:
            msg.setText("Invalid file name or file path has been selected.")
        elif self.errorCode == 3:
            msg.setText("The Selected Paysht INP File is Invalid.")
        elif self.errorCode == 4:
            msg.setText("The selected column is not pure text.")
        elif self.errorCode == 6:
            msg.setText("The input file(s) not exist. Please check the DATA\\current\\input_files folder")
        elif self.errorCode == 8:
            msg.setText("There are missing entries in PAYSHT File.")
            msg.setDetailedText(self.missingIndex)
        msg.setWindowTitle("Oppos...")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    """successMsg - Pop Up Window Box to show the File Saved Success Message"""
    def successMsg(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("Operation Complete. File Saved at: "+os.path.normpath(os.path.abspath(self.outputFilePath)))
        msg.setWindowTitle("Success!")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    """sucArcMsg - Pop Up Window Box to show the File Achived Success Message"""
    def sucArcMsg(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("File Archived. File Saved at: "+os.path.normpath(os.path.abspath(self.arcDirectory)))
        msg.setWindowTitle("Success!")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    """capColumn - Capitalise the String data in a specific column"""
    def capColumn(self, df, columnIdx):
        for idx in df.index:
            if not isinstance(df.iloc[idx][columnIdx],str):
                self.errorCode = 4
                self.errorMsg()
                sys.exit()
            content = df.iloc[idx][columnIdx]
            contentUpper = content.upper()
            df.set_value(idx,columnIdx,contentUpper)
        return df

    """emCodeExt - extract employee code from Attache import file"""
    def emCodeExt(self, df, colIdx):
        emCodes = []
        for idx in df.index:
            code = df.iloc[idx][colIdx]
            if not(code in emCodes):
                emCodes.append(code)
        return emCodes

    """casEmCodeExt - extract employee code of casual workers from DCW Rates file"""
    def casEmCodeExt(self, df):
        return [str(x) for x in df[df["Employment Status"] == "C"]["Employee Code"].tolist()]

    """emPayExt - extract employee salary from Attache import file"""
    def emPayExt(self, df, emCodes):
        emPays = {}
        pays = [0]*2
        for code in emCodes:
            tableSlice = df[(df[2] == code) & ((df[5] != "BACKPAY") & (df[5] != "MILEINT") & (df[5] != "MILEOUT") & (df[5] != "TRANSPORT") & (df[9] != 0.78))]
            highPay = max(tableSlice[9])
            lowPay = min(tableSlice[9])
            highPaySection = tableSlice[tableSlice[9]==highPay]
            lowPaySection = tableSlice[tableSlice[9]==lowPay]
            if (highPaySection.iloc[0][5] == "ES"):
                pays[0] = highPaySection.iloc[0][9] * 8
            elif (highPaySection.iloc[0][5] == "NS"):
                pays[0] = highPaySection.iloc[0][9] / 0.15
            elif (highPaySection.iloc[0][5] == "SAT"):
                pays[0] = highPaySection.iloc[0][9] * 2
            elif (highPaySection.iloc[0][5] == "PHLOAD"):
                pays[0] = highPaySection.iloc[0][9] / 1.5
            else:
                pays[0] = highPaySection.iloc[0][9]
            if (lowPaySection.iloc[0][5] == "ES"):
                pays[1] = lowPaySection.iloc[0][9] * 8
            elif (lowPaySection.iloc[0][5] == "NS"):
                pays[1] = lowPaySection.iloc[0][9] / 0.15
            elif (lowPaySection.iloc[0][5] == "SAT"):
                pays[1] = lowPaySection.iloc[0][9] * 2
            elif (lowPaySection.iloc[0][5] == "PHLOAD"):
                pays[1] = lowPaySection.iloc[0][9] / 1.5
            else:
                pays[1] = lowPaySection.iloc[0][9]
            emPays[code] = pays[:]
        return emPays

    """mileInCheck - Check the recorded milage data"""
    def mileInCheck(self, dfA, dfM, emCodes):
        hasMismatch = False
        self.mismatchInfo = "You have mismatched mileage record, the total mileage calculated from per visit is not the same as the one recorded in Attache import data file\n"
        for code in emCodes:
            dfA_slice = dfA[(dfA[2] == code) & (dfA[5]=="MILEINT")]
            ccA = dfA_slice[12].tolist()
            mileAcc = dfA_slice[6].tolist()
            mileAcc = [ '%.3f' % elem for elem in mileAcc ]
            mileAcc = map(float, mileAcc)
            mileageDict = dict(zip(ccA, mileAcc))
            dfM_slice = dfM[dfM["Employee Code"] == code]
            for cc in ccA:
                dfM_slice_fine = dfM_slice[dfM_slice["Client Cost Center"] == cc]
                totalMile = dfM_slice_fine["KMs"].sum()
                totalMile = round(totalMile,3)
                if~(totalMile == mileageDict[cc]):
                    diff = abs(totalMile - mileageDict[cc])
                    if diff > 0.005:
                        self.mismatchInfo = self.mismatchInfo + "Employee Code: {}, Cost Centre: {}. Attache Value: {}, Calculated Value: {}.\n".format(code, cc, mileageDict[cc], totalMile)
                        hasMismatch = True
        if hasMismatch:
            self.errorCode = 5
            self.errorMsg()
            sys.exit()


    """mileRM - Remove unrelated records"""
    def mileRM(self, dfM):
        visitDate = dfM["Visit Date"]
        theDate = visitDate.apply(lambda x: x.split(' ')[0])
        x = theDate[theDate.index[0]]

        # Added Additional Conditions to check the employee Code #
        theEmployee = dfM["Employee Code"]
        y = theEmployee[theEmployee.index[0]]
        # End of Additional Conditions #

        resultIndex = []
        for i in range(1,len(theDate.index)):
            l = [ord(a) ^ ord(b) for a,b in zip(x,theDate[theDate.index[i]])]
            m = [ord(c) ^ ord(d) for c,d in zip(y,theEmployee[theEmployee.index[i]])]
            x = theDate[theDate.index[i]]
            y = theEmployee[theEmployee.index[i]]
            if (sum(l) == 0) and (sum(m) == 0):
                resultIndex.append(theDate.index[i-1])
        resultIndex = list(set(resultIndex))
        dfM_FF = dfM.loc[resultIndex]
        return dfM_FF

    """mileBonus - give additional 10% of mileage amount to its actual travel"""
    def mileBonus(self, dfM):
        dfM_KMs = dfM["KMs"]
        dfM_KMs = dfM_KMs.apply(lambda x:x*1.10)
        dfM["KMs"] = dfM_KMs
        return dfM

    """mileInLoading - Calculate additional mileage loading"""
    def mileInLoading(self, dfA, dfM, emCodes):
        extraMileage = {}
        for code in emCodes:
            workerExtraMiles = {}
            dfA_slice = dfA[(dfA[2] == code) & (dfA[5]=="MILEINT")]
            ccA = dfA_slice[12].tolist()
            dfM_slice = dfM[dfM["Employee Code"] == code]
            for cc in ccA:
                extraMiles = []
                dfM_slice_fine = dfM_slice[dfM_slice["Client Cost Center"] == cc]
                visitTimes = dfM_slice_fine["Visit Date"]
                visitTimes = visitTimes.apply(lambda x: int(x[len(x)-5:len(x)-3]))
                ordShifts = dfM_slice_fine.loc[visitTimes>=6]
                ordShifts = ordShifts.loc[visitTimes<20]
                #ordShifts1 = dfM_slice_fine[visitTimes>=6]
                #ordShifts1 = ordShifts1[visitTimes<20]
                esShifts = dfM_slice_fine[visitTimes>=20]
                nsShifts = dfM_slice_fine[visitTimes<6]
                """ Extra Mileage for ORD work shifts """
                if ordShifts.empty:
                    ordextraMiles = []
                else:
                    ordmileages = ordShifts["KMs"]
                    ordmileages= round(ordmileages,2)
                    ordmileCandidates = ordmileages[ordmileages > 10.0]
                    ordextraMiles = (ordmileCandidates-10.0)/30.0
                    ordextraMiles = ordextraMiles.tolist()
                """ Extra Mileage for ES work shifts """
                if esShifts.empty:
                    esextraMiles = []
                else:
                    esmileages = esShifts["KMs"]
                    esmileages= round(esmileages,2)
                    esmileCandidates = esmileages[esmileages > 10.0]
                    esextraMiles = (esmileCandidates-10.0)/30.0
                    esextraMiles = esextraMiles.tolist()
                """ Extra Mileage for ES work shifts """
                if nsShifts.empty:
                    nsextraMiles = []
                else:
                    nsmileages = nsShifts["KMs"]
                    nsmileages= round(nsmileages,2)
                    nsmileCandidates = nsmileages[nsmileages > 10.0]
                    nsextraMiles = (nsmileCandidates-10.0)/30.0
                    nsextraMiles = nsextraMiles.tolist()
                extraMiles=[ordextraMiles,esextraMiles,nsextraMiles]
                workerExtraMiles[cc] = extraMiles
            extraMileage[code] = workerExtraMiles
        return extraMileage


    """addML - Add the additional mileage loading to existing Attache Data Records"""
    def addML(self, dfA, emCodes, emPays, extraMiles):
        newDF = pd.DataFrame(columns=range(0,23)) # New Data Frame for store the update records
        for code in emCodes:
            newDF = newDF.append(dfA[dfA[2] == code], ignore_index=True)
            for cc in extraMiles[code]:
                for shiftIdx in range(3):
                    shift = extraMiles[code][cc][shiftIdx]
                    listLen = len(shift)
                    if listLen > 0:
                        for miles in shift:
                            newRow = pd.Series(index=range(23))
                            newRow[0] = "T6"
                            newRow[2] = code
                            newRow[4] = "N"
                            newRow[5] = "ORD"
                            newRow[6] = round(miles,2)
                            if (shiftIdx == 0):
                                newRow[9] = emPays[code][0]
                            else:
                                newRow[9] = emPays[code][1]
                            newRow[12] = cc
                            newDF = newDF.append(newRow, ignore_index=True)
        return newDF

    """reGenData - Regenerate Attache Data Records"""
    def reGenData(self, dfA, dfM, emCodes, emPays, extraMiles):
        newDF = pd.DataFrame(columns=range(0,23)) # New Data Frame for store the update records
        for code in emCodes:
            dfM_slice = dfM[dfM["Employee Code"] == code]
            for cc in extraMiles[code]:
                dfM_slice_fine = dfM_slice[dfM_slice["Client Cost Center"] == cc]
                if round(sum(dfM_slice_fine["KMs"]),2) != 0.000:
                    newRow = pd.Series(index=range(23))
                    newRow[0] = "T6"
                    newRow[2] = code
                    newRow[4] = "A"
                    newRow[5] = "MILEINT"
                    newRow[6] = round(sum(dfM_slice_fine["KMs"]),2)
                    newRow[9] = 0.78
                    newRow[12] = cc
                    newDF = newDF.append(newRow, ignore_index=True)
            newDF = newDF.append(dfA[dfA[2] == code], ignore_index=True)
            for cc in extraMiles[code]:
                for shiftIdx in range(3):
                    shift = extraMiles[code][cc][shiftIdx]
                    listLen = len(shift)
                    if listLen > 0:
                        for miles in shift:
                            newRow = pd.Series(index=range(23))
                            newRow[0] = "T6"
                            newRow[2] = code
                            newRow[4] = "N"
                            newRow[5] = "ORD"
                            newRow[6] = round(miles,2)
                            if (shiftIdx == 0):
                                newRow[9] = emPays[code][0]
                            else:
                                newRow[9] = emPays[code][1]
                            newRow[12] = cc
                            newDF = newDF.append(newRow, ignore_index=True)
        return newDF


    """addLL - Add 'leaving load' (LL) to AL paytype"""
    def addLL(self, df):
        # If paytype is AL, add a new row with paytype set to LL,
        # and the 9th column is set to 17.5% leaving load of the current rate (the 9th column of the AL row)
        newDF = pd.DataFrame(columns=range(0,23)) # New Data Frame for store the update records
        for idx in df.index:
            newDF = newDF.append(df.iloc[idx], ignore_index=True)
            paytype = df.iloc[idx][5]
            if (paytype == "AL"):
                LL = df.iloc[idx].copy()
                LL[5] = LL[5].replace("AL", "LL")
                LL[9] = LL[9] * 0.175
                newDF = newDF.append(LL, ignore_index=True)
        return newDF

    """addInternetAllowance - Add INTERNET allowance to each staff"""
    def addInternetAllowance(self, df, emCodes):
        newDF = pd.DataFrame(columns=range(0,23)) # New Data Frame for store the update records
        for code in emCodes:
            df_slice = df[df[2] == code]
            newDF = newDF.append(df_slice, ignore_index=True)
            df_ord_slice = df_slice[df_slice[5] == "ORD"]
            if sum(df_ord_slice[6]) > 20:
                internet_rate = 2.000
            else:
                internet_rate = 1.000
            new_row = df_slice.iloc[0].copy()
            new_row[4] = 'A'
            new_row[5] = 'INTERNET'
            new_row[6] = internet_rate
            new_row[9] = 1.250
            newDF = newDF.append(new_row, ignore_index=True)
        return newDF


    """rpPayType - Replace OUTING paytype with MILEOUT
                   Replace TRANSPORT paytype with MILEIN
                   Replace SL paytype with PCL
    """
    def rpPayType(self, df):
        df = df.apply(lambda x: x.replace("OUTING","MILEOUT"))
        df = df.apply(lambda x: x.replace("MILEAGE(OUTING)","MILEOUT"))
        df = df.apply(lambda x: x.replace("TRANSPORT","MILEOUT"))
        df = df.apply(lambda x: x.replace("TRANSPORT SERVICES","MILEOUT"))
        df = df.apply(lambda x: x.replace("SL","PCL"))
        return df

    """updateSatCasual - Update Saturday Casual rate (multiplied by 0.2)"""
    def updateSatCasual(self, df, casEmCodes):
        # use "mask" to filter for Saturday casual rows
        mask = (df[2].isin(casEmCodes)) & (df[5] == 'SAT') & (df[4] == 'A')
        df.loc[mask, 5] = df.loc[mask, 5].map(lambda x: x.replace('SAT', 'SATCAS'))
        df.loc[mask, 4] = df.loc[mask, 4].map(lambda x: x.replace('A', 'N'))
        df.loc[mask, 9] = df.loc[mask, 9].map(lambda x: x * 0.2)
        return df

    """updateSunCasual - Update Sunday Casual rate (multiplied by 0.6)"""
    def updateSunCasual(self, df, casEmCodes):
        # use "mask" to filter for Sunday casual rows
        mask = (df[2].isin(casEmCodes)) & (df[5] == 'SUN') & (df[4] == 'A')
        df.loc[mask, 5] = df.loc[mask, 5].map(lambda x: x.replace('SUN', 'SUNCAS'))
        df.loc[mask, 4] = df.loc[mask, 4].map(lambda x: x.replace('A', 'N'))
        df.loc[mask, 9] = df.loc[mask, 9].map(lambda x: x * 0.6)
        return df

    """updatePhloadCasual - Update PHLOAD Casual rate (multiplied by 1.2)"""
    def updatePhloadCasual(self, df, casEmCodes):
        # use "mask" to filter for PHLOAD rows for casual workers
        mask = (df[2].isin(casEmCodes)) & (df[5] == 'PHLOAD')
        df.loc[mask, 5] = df.loc[mask, 5].map(lambda x: x.replace('PHLOAD', 'PHCAS'))
        df.loc[mask, 9] = df.loc[mask, 9].map(lambda x: x * 1.2)
        return df

    """changeCol5toN - Change col 5 to 'N' if the 6th col is SAT/SUN/PHLOAD/PHNW/PHCAS"""
    def changeCol5toN(self, df):
        mask = df[5].isin(['SAT', 'SUN', 'PHLOAD', 'PHNW', 'PHCAS'])
        df.loc[mask, 4] = 'N'
        return df

    """ genTimesheet - Generate the timesheet dataframe"""
    def genTimesheet(self, emCodes, df_main, df_dcw_rates):
        #headers = ['Employee', 'Name', 'Cost Center', 'Mileint', 'Mileout', 'KM>10', 'Ord', 'Training', 'Mkup', 'Total', 'Sat Loading', 'Sun Loading', 'SatCasual Loading', 'SunCasual Loading	ES	NS	PH	PH Loading	PHNW	AL	LL	PCL	Compassionate Leave	OT1x5	OT2	OT2x5	Internet]
        df = pd.DataFrame(column)
#### TODO: I WAS HERE

    """encodeStrWithAscii - Encode string columns with Ascii method"""
    def encodeStrWithAscii(self, df):
        df[0] = df[0].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        df[2] = df[2].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        df[4] = df[4].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        df[5] = df[5].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        df[12] = df[12].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        return df


    """ arcFile - Archive the source input files and generated Attache Output Data files to the default archieve folder """
    def archiveFile(self):
        cDate = datetime.datetime.now()
        cDateString = cDate.strftime("%Y%m%d")
        self.arcDirectory = "DATA/archive/"+cDateString
        if not(os.path.isdir(self.arcDirectory)):
            os.mkdir(self.arcDirectory)
        shutil.copyfile(self.attacheFilePath, self.arcDirectory+"/attacheExport_"+cDateString+".csv")
        shutil.copyfile(self.mileageFilePath, self.arcDirectory+"/mileageExport_"+cDateString+".csv")
        shutil.copyfile(self.outputFilePath, self.arcDirectory+"/PAYTSHT_"+cDateString+".INP")


    """Main part of the program"""
    def runProgram(self):
        """
        Data Frame (DF) variables used in this function:
        - df_main: df created from PAYSHT (ATTACHE) input file
        - df_dcw_rates: df created from DCW Rates files
        - df_ts: df created for the timesheet in the output file
        - df_sum: df created for the summary report in the output file
        """
        """ Extract Feild Entry """
        PayshtFileName = self.mainPage.field("PayshtFileName")
        DCWRatesFileName = self.mainPage.field("DCWRatesFileName")
        OutputFileName = self.mainPage.field("OutputFileName")

        """ Check User inputs"""
        if (self.mainPage.runMode == 2):
            # runMode = 2 means custom mode
            if (not PayshtFileName) and (not DCWRatesFileName) and (not OutputFileName): # Four Fields are all empty
                self.errorCode = 1
                self.errorMsg()
                sys.exit()
            elif (self.mainPage.inputFile1NameLineEdit == ".") or (DCWRatesFileName == ".") or (OutputFileName == "."): # One of the Field has a None string due to Cancel btn clicked
                self.errorCode = 2
                self.errorMsg()
                sys.exit()
            else:
                if PayshtFileName: # Paysht Custom File is selected
                    self.payshtFilePath = self.mainPage.inf1PH_C[0]
                else: # Paysht Custom File is not selected
                    self.payshtFilePath = self.mainPage.inf1PH

                if DCWRatesFileName: # DCW Rates Custom File is selected
                    self.DCWRatesFilePath = self.mainPage.inf2PH_C[0]
                else: # DCW Rates Custom File is not selected
                    self.DCWRatesFilePath = self.mainPage.inf2PH

                if OutputFileName: # Custom Output Directory is defined
                    self.outputFilePath = self.mainPage.outfPH_C + "/Timesheet & Summary Report .xlsx"
                else: # Custom Output Directory is not defined
                    self.outputFilePath = self.mainPage.outfPH
        else:
            # assume default mode was in use
            self.payshtFilePath = self.mainPage.inf1PH
            self.DCWRatesFilePath = self.mainPage.inf2PH
            self.outputFilePath = self.mainPage.outfPH

        """ Check if the existance of the input files """
        if not(os.path.isfile(self.payshtFilePath)) or not(os.path.isfile(self.DCWRatesFilePath)):
            self.errorCode = 6
            self.errorMsg()
            sys.exit()

        """ Detect the file encoding scheme """
        f = open(self.payshtFilePath,"rb")
        file = f.read()
        detectResult = chardet.detect(file)

        """ Load the PAYSHT file into a data frame """
        df_main = pd.read_csv(self.payshtFilePath,header=None,converters={2 : str, 4 : str, 5 : str, 12 : str})

        """ Check to see if a valid Attache Data File is provided """
        self.missingIndex = "You have missing entry. Please check: \n"
        misIdxs = []
        if detectResult["encoding"] == "ascii":
            firstEntry = df_main.iloc[0][0]
        else:
            firstEntry = df_main.iloc[1][0]
        if not(firstEntry == "T6"):
                self.errorCode = 3
                self.errorMsg()
                sys.exit()
        elif not(len(list(df_main.columns.values)) == 24):
            self.errorCode = 3
            self.errorMsg()
            sys.exit()
        #misIdxs = (np.isnan(df_main.iloc[:,2]).index.tolist()) + (np.isnan(df_main.iloc[:,4]).index.tolist()) + (np.isnan(df_main.iloc[:,5]).index.tolist()) + (np.isnan(df_main.iloc[:,6]).index.tolist()) + (np.isnan(df_main.iloc[:,9]).index.tolist())+ (np.isnan(df_main.iloc[:,12]).index.tolist())
        colValues = df_main.iloc[:,2]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        colValues = df_main.iloc[:,4]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        colValues = df_main.iloc[:,5]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        colValues = df_main.iloc[:,6]
        misIdxs = misIdxs + colValues[np.isnan(colValues)].index.tolist()
        colValues = df_main.iloc[:,9]
        misIdxs = misIdxs + colValues[np.isnan(colValues)].index.tolist()
        colValues = df_main.iloc[:,12]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        misIdxs = list(set(misIdxs))
        if len(misIdxs) > 0:
            for misIdx in misIdxs:
                self.missingIndex = self.missingIndex + "Line: {}\n".format(misIdx+1)
            self.errorCode = 8
            self.errorMsg()
            sys.exit()


        """ Generate employee code lists """
        emCodes = self.emCodeExt(df_main,2)

        """ Load the DCW Rates into a data frame """
        df_dcw_rates = pd.read_excel(self.DCWRatesFilePath)

        """ Create the timesheet dataframe """
        df_ts = self.genTimesheet(emCodes, df_main, df_dcw_rates)


        """ Extract Employee Salary """
        emPays = self.emPayExt(df_main, emCodes)


        """ Load the mileage data into a data frame """
        df_mile_org = pd.read_csv(self.mileageFilePath, converters={"Employee Code" : str}, sep="\t")

        """ Check to see if a valid Mileage Data File is provided """
        df_mile_header = list(df_mile_org.columns.values)
        if not(len(df_mile_header) == 6):
            self.errorCode = 7
            self.errorMsg()
            sys.exit()
        if df_mile_header != ["Worker Name", "Employee Code", "Client Name", "Client Cost Center", "Visit Date", "KMs"]:
            self.errorCode = 7
            self.errorMsg()
            sys.exit()

        # Filter the table to remove non recorded rows
        df_mile = df_mile_org[~(df_mile_org["KMs"].isnull())]

        """ Verify if there is any error with the recorded mileage """
        checkResult = self.mileInCheck(df_main, df_mile, emCodes)

        """ Remove unrelated Mileage Records """
        df_mile_filtered = self.mileRM(df_mile)

        """ Calculate the bonus 10% mileage """
        df_mile_filtered = self.mileBonus(df_mile_filtered)

        """ Determine the additional amount of Milein needed for each worker """
        extraMiles = self.mileInLoading(df_main, df_mile_filtered, emCodes)

        """ Generate the new dataframe with additional mileage loading """
        #newDF = self.addML(df_main_filtered,emCodes,emPays,extraMiles)
        #print("run to here")

        """ Generate the new dataframe with additional mileage loading """
        newDF = self.reGenData(df_main_filtered, df_mile_filtered, emCodes, emPays, extraMiles)

        """ Add additional leaving loading after AL paytype"""
        newDF = self.addLL(newDF)

        """ Add Internet allowance for each staff"""
        newDF = self.addInternetAllowance(newDF, emCodes)

        """ Replace the Pay Type """
        newDF = self.rpPayType(newDF)




        """ Generate casual employee code lists """
        casEmCodes = self.casEmCodeExt(df_dcw_rates)

        """ Update the SAT rates for casual workers """
        newDF = self.updateSatCasual(newDF, casEmCodes)

        """ Update the SUN rates for casual workers """
        newDF = self.updateSunCasual(newDF, casEmCodes)

        """ Update the PHLOAD rates for casual workers """
        newDF = self.updatePhloadCasual(newDF, casEmCodes)

        """ Change column 5 from A to N if the 6th col is SAT/SUN/PHLOAD/PHNW/PHCAS """
        newDF = self.changeCol5toN(newDF)

        """ Re-encode the string columns using ASCII method """
        newDF = self.encodeStrWithAscii(newDF)

        """ Save the updated records to a csv file """
        newDF.to_csv(self.outputFilePath,header=None,index=None,encoding="ascii",float_format='%.3f')
        self.successMsg()

        """ Archive the files if requested by user """
        if self.mainPage.archive:
            self.archiveFile()
            self.sucArcMsg()



class MainPage(QtWidgets.QWizardPage):

    def __init__(self, parent=None):
        super(MainPage, self).__init__(parent)

        self.inf1PH = "DATA/current/input_files/PAYTSHT.INP"
        self.inf2PH = "DATA/current/input_files/DCW_Rates.xlsx"
        self.outfPH = "DATA/current/output_file/Timesheet & Summary Report.xlsx"

        self.runMode = 0
        self.archive = True

        self.setTitle("Welcome!")
        self.setPixmap(QtWidgets.QWizard.WatermarkPixmap, QtGui.QPixmap("CASSCare.png"))


        label = QtWidgets.QLabel("Generate timesheet and payment summary report. \n")

        label.setWordWrap(True)

        inputFile1NameLabel = QtWidgets.QLabel("Choose Your PAYSHT INP File (.INP):")
        self.inputFile1NameLineEdit = QtWidgets.QLineEdit()
        inputFile1NameLabel.setBuddy(self.inputFile1NameLineEdit)

        inputFile1SelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.inf1PH_C = ""

        inputFile2NameLabel = QtWidgets.QLabel("Choose Your DCW Rates File (.xlsx):")
        self.inputFile2NameLineEdit = QtWidgets.QLineEdit()
        inputFile2NameLabel.setBuddy(self.inputFile2NameLineEdit)

        inputFile2SelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.inf2PH_C = ""

        outputFileNameLabel = QtWidgets.QLabel("Select the folder for the Output File:")
        self.outputFileNameLineEdit = QtWidgets.QLineEdit()
        outputFileNameLabel.setBuddy(self.outputFileNameLineEdit)

        outputFileSelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.outfPH_C = ""

        archiveCheckBox = QtWidgets.QCheckBox("Archive Files?")
        archiveCheckBox.setChecked(True)

        self.registerField("PayshtFileName",self.inputFile1NameLineEdit)
        self.registerField("DCWRatesFileName",self.inputFile2NameLineEdit)
        self.registerField("OutputFileName",self.outputFileNameLineEdit)
        self.registerField("fileArchiveFlag",archiveCheckBox)

        runModeBox = QtWidgets.QGroupBox("Running Mode")

        defaultModeRadioButton = QtWidgets.QRadioButton("Default Mode")
        customModeRadioButton = QtWidgets.QRadioButton("Custom Mode")

        defaultModeRadioButton.setChecked(True)

        if defaultModeRadioButton.isChecked():
            self.runMode = 1
            self.inputFile1NameLineEdit.setEnabled(False)
            self.inputFile2NameLineEdit.setEnabled(False)
            self.outputFileNameLineEdit.setEnabled(False)
            archiveCheckBox.setEnabled(False)
            inputFile1SelectionBtn.setEnabled(False)
            inputFile2SelectionBtn.setEnabled(False)
            outputFileSelectionBtn.setEnabled(False)
            self.inputFile1NameLineEdit.setText(os.path.normpath(self.inf1PH))
            self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf2PH))
            self.outputFileNameLineEdit.setText(os.path.normpath(self.outfPH))



        #defaultModeRadioButton.toggled.connect(ToggleFlag)

        #defaultModeRadioButton.toggled.connect(self.setButtonText(QtWidgets.QWizard.WizardButton(1), "run!"))

        #defaultModeRadioButton.toggled.connect(inputFile1NameLineEdit.setDisabled)
        #defaultModeRadioButton.toggled.connect(inputFile2NameLineEdit.setDisabled)
        #defaultModeRadioButton.toggled.connect(outputFileNameLineEdit.setDisabled)
        defaultModeRadioButton.toggled.connect(archiveCheckBox.setDisabled)
        defaultModeRadioButton.toggled.connect(archiveCheckBox.setChecked)
        defaultModeRadioButton.toggled.connect(inputFile1SelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(inputFile2SelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(outputFileSelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(self.alterMode)

        archiveCheckBox.stateChanged.connect(self.alterArchive)

        inputFile1SelectionBtn.clicked.connect(self.selectPayshtFile)
        inputFile2SelectionBtn.clicked.connect(self.selectDCWRatesFile)
        outputFileSelectionBtn.clicked.connect(self.selectOutputDirc)

        self.registerField("defaultMode", defaultModeRadioButton)
        self.registerField("customMode", customModeRadioButton)

        runModeBoxLayout = QtWidgets.QVBoxLayout()
        runModeBoxLayout.addWidget(defaultModeRadioButton)
        runModeBoxLayout.addWidget(customModeRadioButton)
        runModeBox.setLayout(runModeBoxLayout)

        layout = QtWidgets.QGridLayout()
        layout.addWidget(label,0,0)
        layout.addWidget(runModeBox,1,0)

        layout.addWidget(inputFile1NameLabel,2,0)
        layout.addWidget(self.inputFile1NameLineEdit,3,0,1,2)
        layout.addWidget(inputFile1SelectionBtn,3,3,1,1)

        layout.addWidget(inputFile2NameLabel,4,0)
        layout.addWidget(self.inputFile2NameLineEdit,5,0,1,2)
        layout.addWidget(inputFile2SelectionBtn,5,3,1,1)

        layout.addWidget(outputFileNameLabel,6,0)
        layout.addWidget(self.outputFileNameLineEdit,7,0,1,2)
        layout.addWidget(outputFileSelectionBtn,7,3,1,1)
        layout.addWidget(archiveCheckBox)
        self.setLayout(layout)

    def alterMode(self):
        if self.runMode == 1:
            self.runMode = 2
            self.inputFile1NameLineEdit.clear()
            self.inputFile2NameLineEdit.clear()
            self.outputFileNameLineEdit.clear()
            self.archive = False
        else:
            self.runMode = 1
            self.inputFile1NameLineEdit.setText(os.path.normpath(self.inf1PH))
            self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf2PH))
            self.outputFileNameLineEdit.setText(os.path.normpath(self.outfPH))
            self.archive = True

    def alterArchive(self):
        if self.archive:
            self.archive = False
        else:
            self.archive = True

    def selectPayshtFile(self):
        self.inf1PH_C = QtWidgets.QFileDialog.getOpenFileName(self, 'Choose Paysht INP File..','',"Comma Separated Values File (*.csv);;Attache Data File (*.INP);;All Files (*)")
        self.inputFile1NameLineEdit.setText(os.path.normpath(self.inf1PH_C[0]))

    def selectDCWRatesFile(self):
        self.inf2PH_C = QtWidgets.QFileDialog.getOpenFileName(self, 'Choose DCW Rates File..','',"Excel File (*.xlsx);;All Files (*)")
        self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf2PH_C[0]))

    def selectOutputDirc(self):
        self.outfPH_C = QtWidgets.QFileDialog.getExistingDirectory(self, 'Choose Saving Directory..','', QtWidgets.QFileDialog.ShowDirsOnly)
        self.outputFileNameLineEdit.setText(os.path.normpath(self.outfPH_C))
        #self.outputFileNameLineEdit.setText((self.outfPH_C))

    def saveFile(self):
        AttacheFileName = self.field("PayshtFileName")
        OutputFileName = self.field("OutputFileName")
        with open("Output.txt", "w") as text_file:
            text_file.write(OutputFileName)
            text_file.write("\n")
            text_file.write("The Operation Mode is: {}".format(self.runMode))
            text_file.write("\n")
            if (not(OutputFileName == ".")) and (OutputFileName):
                text_file.write("Field is not empty!")
            else:
                text_file.write("Field is empty!")



if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    wizard = myApp()
    wizard.show()
    sys.exit(app.exec_())
