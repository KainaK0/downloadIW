# %%
from sap4me import gototx
from sap4me import gettxdata
from sap4me import table2dataframe
import win32com.client
import subprocess
import logging
import os
import time
import pandas as pd
logger = logging.Logger('catch_all')

# %%
def saplogin():
    try:
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        #time.sleep(10)
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:   
            return
        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        #connection = application.OpenConnection("ConnectionName", True)
        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        ruteAvOt = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\02. Ordenes Avisos\01. raw2data"
        
        start=time.time()
        # Download tx iw29
        iw29tx = "iw29"
        iw29varian = "/BDIW29"
        iw29SAPuser = ""
        iw29layout = "/BDIW29"
        gototx(session,iw29tx,iw29varian,iw29SAPuser,iw29layout)
        iw29dataname = "iw29.txt"
        gettxdata(session,iw29tx,iw29dataname,ruteAvOt,"4110")
        end=time.time()
        print("get raw IW29 | time: {} min".format(round((end-start)/60,1)))
        
        # Download tx iw39
        start=time.time()
        iw39tx = "iw39"
        iw39varian = "/BDIW39"
        iw39SAPuser = ""
        iw39layout = "/BDIW39"
        gototx(session,iw39tx,iw39varian,iw39SAPuser,iw39layout)
        iw39dataname = "iw39.txt"
        gettxdata(session,iw39tx,iw39dataname,ruteAvOt,"4110")
        end=time.time()
        print("get raw IW39 | time:{} min".format(round((end-start)/60,1)))

        # Download tx iw37n
        start=time.time()
        iw37ntx = "iw37n"
        iw37nvarian = "/BDIW37N"
        iw37nSAPuser = ""
        iw37nlayout = "/BDIW37N"
        gototx(session,iw37ntx,iw37nvarian,iw37nSAPuser,iw37nlayout)
        iw37ndataname = "iw37n.txt"
        gettxdata(session,iw37ntx,iw37ndataname,ruteAvOt,"4110")   
        end=time.time()
        print("get raw IW37N | time:{} min".format(round((end-start)/60,1)))        

        ruteMB = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\04. Materiales\01. raw2data"
        
        # Download tx mb25
        start=time.time()
        mb25tx = "mb25"
        mb25varian = "/BDMB25"
        mb25SAPuser = ""
        mb25layout = "/BDMB25"
        gototx(session,mb25tx,mb25varian,mb25SAPuser,mb25layout)
        mb25dataname = "mb25.txt"
        gettxdata(session,mb25tx,mb25dataname,ruteMB,"4110")
        end=time.time()
        print("get raw MB25 | time:{} min".format(round((end-start)/60,1)))

        # Download tx mb52
        
        start=time.time()
        mb52tx = "mb52"
        mb52varian = "/BDMB52"
        mb52SAPuser = "JCALLOMAMANB"
        mb52layout = "/BDMB52_S"
        gototx(session,mb52tx,mb52varian,mb52SAPuser,mb52layout)
        mb52dataname = "mb52_s.txt"
        gettxdata(session,mb52tx,mb52dataname,ruteMB,"4110")
        end=time.time()
        print("get raw MB52 | time:{} min".format(round((end-start)/60,1)))

        ruteME = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\05. Solpeds y Pedidos\01. raw2data"
        
        # Download tx me5a

        start=time.time()
        me5atx = "me5a"
        me5avarian = "/BDME5A"
        me5aSAPuser = ""
        me5alayout = "/BDME5A"
        gototx(session,me5atx,me5avarian,me5aSAPuser,me5alayout)
        me5adataname = "me5a.txt"
        gettxdata(session,me5atx,me5adataname,ruteME,"4110")
        end=time.time()
        print("get raw ME5A | time:{} min".format(round((end-start)/60,1)))

        # Download tx me2n

        start=time.time()
        me2ntx = "me2n"
        me2nvarian = "/BDME2N"
        me2nSAPuser = "JCALLOMAMANB"
        me2nlayout = "/BDME2N"
        gototx(session,me2ntx,me2nvarian,me2nSAPuser,me2nlayout)
        me2ndataname = "me2n.txt"
        gettxdata(session,me2ntx,me2ndataname,ruteME,"4110")
        end=time.time()
        print("get raw ME2N | time:{} min".format(round((end-start)/60,1)))

        ruteKOKS = r'C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\01. Costos\01. raw2data'

        # # Download tx kob1
        # start=time.time()
        # kob1tx = "kob1"
        # kob1varian = "/BDKOB1"
        # kob1SAPuser = ""
        # kob1layout = "/BDKOB1"
        # gototx(session,kob1tx,kob1varian,kob1SAPuser,kob1layout)
        # kob1dataname = "kob1.txt"
        # gettxdata(session,kob1tx,kob1dataname,ruteKOKS,"4110")
        # end=time.time()ll
        # print("get raw KOB1 | time:{} min".format(round((end-start)/60,1)))

        # # Download tx ksb1
        # start=time.time()
        # ksb1tx = "ksb1"
        # ksb1varian = "/BDKSB1"
        # ksb1SAPuser = ""
        # ksb1layout = "/BDKSB1"
        # gototx(session,ksb1tx,ksb1varian,ksb1SAPuser,ksb1layout)
        # ksb1dataname = "ksb1.txt"
        # gettxdata(session,ksb1tx,ksb1dataname,ruteKOKS,"4110")
        # end=time.time()
        # print("get raw KSB1 | time:{} min".format(round((end-start)/60,1)))


    except Exception as e:
        logger.error(e, exc_info=True)
        pass

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None

saplogin()

# %%
ruteAvOt = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\02. Ordenes Avisos\01. raw2data"    
ruteMB = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\04. Materiales\01. raw2data"
ruteME = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\05. Solpeds y Pedidos\01. raw2data"
ruteKOKS = r'C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\01. Costos\01. raw2data'

start = time.time()
iw29dataname = "iw29.txt"
fileiw29 = os.path.join(ruteAvOt,iw29dataname)
rawiw29 = pd.read_table(fileiw29, header = None, sep= '\n', encoding="utf-8-sig")
dataiw29 = table2dataframe(rawiw29)
# dataiw29.to_excel(fileiw29[:-4] + ".xlsx", index=False, encoding="UTF-8")
dataiw29.to_csv(fileiw29[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw IW29 transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
iw39dataname = "iw39.txt"
fileiw39 = os.path.join(ruteAvOt,iw39dataname)
rawiw39 = pd.read_table(fileiw39, header = None, sep= '\n', encoding="utf-8-sig")
dataiw39 = table2dataframe(rawiw39)
# dataiw39.to_excel(fileiw39[:-4] + ".xlsx", index=False, encoding="UTF-8")
dataiw39.to_csv(fileiw39[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw IW39 transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
iw37ndataname = "iw37n.txt"
fileiw37n = os.path.join(ruteAvOt,iw37ndataname)
rawiw37n = pd.read_table(fileiw37n, header = None, sep= '\n', encoding="utf-8-sig")
dataiw37n = table2dataframe(rawiw37n)
# dataiw37n.to_excel(fileiw37n[:-4] + ".xlsx", index=False, encoding="UTF-8")
dataiw37n.to_csv(fileiw37n[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw IW37N transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
mb25dataname = "mb25.txt"
filemb25 = os.path.join(ruteMB,mb25dataname)
rawmb25 = pd.read_table(filemb25, header = None, sep= '\n', encoding="utf-8-sig")
datamb25 = table2dataframe(rawmb25)
# datamb25.to_excel(filemb25[:-4] + ".xlsx", index=False, encoding="UTF-8")
datamb25.to_csv(filemb25[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw MB25 transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
mb52dataname = "mb52_s.txt"
filemb52 = os.path.join(ruteMB,mb52dataname)
rawmb52 = pd.read_table(filemb52, header = None, sep= '\n', encoding="utf-8-sig")
datamb52 = table2dataframe(rawmb52)
# datamb52.to_excel(filemb52[:-4] + ".xlsx", index=False, encoding="UTF-8")
datamb52.to_csv(filemb52[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw MB52 transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
me5adataname = "me5a.txt"
fileme5a = os.path.join(ruteME,me5adataname)
rawme5a = pd.read_table(fileme5a, header = None, sep= '\n', encoding="utf-8-sig")
datame5a = table2dataframe(rawme5a)
# datame5a.to_excel(fileme5a[:-4] + ".xlsx", index=False, encoding="UTF-8")
datame5a.to_csv(fileme5a[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw ME5A transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
me2ndataname = "me2n.txt"
fileme2n = os.path.join(ruteME,me2ndataname)
rawme2n = pd.read_table(fileme2n, header = None, sep= '\n', encoding="utf-8-sig")
datame2n = table2dataframe(rawme2n)
# datame2n.to_excel(fileme2n[:-4] + ".xlsx", index=False, encoding="UTF-8")
datame2n.to_csv(fileme2n[:-4] + ".csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw ME2N transformed | time:{} min".format(round((end-start)/60,1)))

# start = time.time()
# kob1dataname = "kob1.txt"
# filekob1 = os.path.join(ruteKOKS,kob1dataname)
# rawkob1 = pd.read_table(filekob1, header = None, sep= '\n', encoding="UTF-8")
# datakob1 = table2dataframe(rawkob1)
# datakob1.to_excel(filekob1[:-4] + ".xlsx", index=False, encoding="UTF-8")
# datakob1.to_csv(filekob1[:-4] + ".csv", index=False, encoding="UTF-8")
# end = time.time()
# print("raw kob1 transformed | time:{} min".format(round((end-start)/60,1)))

# start = time.time()
# ksb1dataname = "ksb1.txt"
# fileksb1 = os.path.join(ruteKOKS,ksb1dataname)
# rawksb1 = pd.read_table(fileksb1, header = None, sep= '\n', encoding="UTF-8")
# dataksb1 = table2dataframe(rawksb1)
# dataksb1.to_excel(fileksb1[:-4] + ".xlsx", index=False, encoding="UTF-8")
# dataksb1.to_csv(fileksb1[:-4] + ".csv", index=False, encoding="UTF-8")
# end = time.time()
# print("raw ksb1 transformed | time:{} min".format(round((end-start)/60,1)))




# %%
ruteAvOt = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\02. Ordenes Avisos\01. raw2data"    
ruteMB = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\04. Materiales\01. raw2data"
ruteME = r"C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\05. Solpeds y Pedidos\01. raw2data"
ruteKOKS = r'C:\Users\jcallomamanib\OneDrive - Minsur S.A\2.0 AREAS\13. Data SAP\00. Data Base\01. DB-SAP\01. Costos\01. raw2data'

iw29model = "IW29M.XLSX"
iw39model = "IW39M.XLSX"
iw37nmodel = "IW37NM.XLSX"
filemodeliw29 = os.path.join(ruteAvOt,iw29model)
filemodeliw39 = os.path.join(ruteAvOt,iw39model)
filemodeliw37n = os.path.join(ruteAvOt,iw37nmodel)

mb25model = "MB25M.XLSX"
mb52model = "MB52_SM.XLSX"
filemodelmb25 = os.path.join(ruteMB,mb25model)
filemodelmb52 = os.path.join(ruteMB,mb52model)

me5amodel = "ME5A_RM.XLSX"
me2nmodel = "ME2N_RM.XLSX"
filemodelme5a = os.path.join(ruteME,me5amodel)
filemodelme2n = os.path.join(ruteME,me2nmodel)

# kob1model = "KOB1M.XLSX"
# ksb1model = "KSB1M.XLSX"
# filemodelkob1 = os.path.join(rawruteKOKS,kob1model)
# filemodelksb1 = os.path.join(rawruteKOKS,ksb1model)

# %%
start = time.time()
IW29M = pd.read_excel(filemodeliw29,sheet_name=0)
end = time.time()
print("raw IW29 transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
IW39M = pd.read_excel(filemodeliw39,sheet_name=0)
end = time.time()
print("raw IW39 transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
IW37NM = pd.read_excel(filemodeliw37n,sheet_name=0)
end = time.time()
print("raw IW37N transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
MB25M = pd.read_excel(filemodelmb25,sheet_name=0)
end = time.time()
print("raw MB25 transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
MB52M = pd.read_excel(filemodelmb52,sheet_name=0)
end = time.time()
print("raw MB52 transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
ME5AM = pd.read_excel(filemodelme5a,sheet_name=0)
end = time.time()
print("raw ME5A transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
ME2NM = pd.read_excel(filemodelme2n,sheet_name=0)
end = time.time()
print("raw ME2N transformed | time:{} min".format(round((end-start)/60,1)))

# start = time.time()
# KOB1M = pd.read_excel(filemodelkob1,sheet_name=0, encoding="UTF-8")
# end = time.time()
# print("raw KOB1 transformed | time:{} min".format(round((end-start)/60,1)))
# start = time.time()
# KSB1M = pd.read_excel(filemodelksb1,sheet_name=0, encoding="UTF-8")
# end = time.time()
# print("raw KSB1 transformed | time:{} min".format(round((end-start)/60,1)))

# %%
dataiw29.columns = IW29M.columns

dataiw39.columns = IW39M.columns

dataiw37n.columns = IW37NM.columns

datamb25.columns = MB25M.columns

datamb52.columns = MB52M.columns

datame5a.columns = ME5AM.columns

datame2n.columns = ME2NM.columns

# datakob1.columns = KOB1M.columns

# dataksb1.columns = KSB1M.columns

# %%
start = time.time()
dataiw29.to_csv(filemodeliw29[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw IW29 transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
dataiw39.to_csv(filemodeliw39[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw IW39 transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
dataiw37n.to_csv(filemodeliw37n[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw IW37N transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
datamb25.to_csv(filemodelmb25[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw MB25 transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
datamb52.to_csv(filemodelmb52[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw MB52 transformed | time:{} min".format(round((end-start)/60,1)))

start = time.time()
datame5a.to_csv(filemodelme5a[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw ME5A transformed | time:{} min".format(round((end-start)/60,1)))
start = time.time()
datame2n.to_csv(filemodelme2n[:-6] + "_.csv", index=False, encoding="utf-8-sig")
end = time.time()
print("raw ME2N transformed | time:{} min".format(round((end-start)/60,1)))

# start = time.time()
# datakob1.to_excel(os.path.join(filemodelkob1[:-10],r"01. raw2data") + filemodelkob1[-10:-5] + "_.xlsx", index=False, encoding="UTF-8")
# end = time.time()
# print("raw KOB1 transformed | time:{} min".format(round((end-start)/60,1)))
# start = time.time()
# dataksb1.to_excel(os.path.join(filemodelksb1[:-10],r"01. raw2data") + filemodelksb1[-10:-5] + "_.xlsx", index=False, encoding="UTF-8")
# end = time.time()
# print("raw KSB1 transformed | time:{} min".format(round((end-start)/60,1)))


