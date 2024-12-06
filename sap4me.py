# %% CREAR OT MANTENIMIENTO
# Importing the Libraries
import pandas as pd
#----------------------------------------------------------------------
# Importing the Libraries
import numpy as np
import pandas as pd
import pyperclip


# %%
def rawData2listClean(db, condition=0):
    db_hasHyphen = (db.iloc[:,0].str.contains(r"[a-z]|[0-9]",case=False,regex=True))
    a = False
    list01 = [0]*len(db)
    list02 = [0]*len(db)
    for e in range(0,len(db),1):
        if e==0:
            list01[e] = (False and (db_hasHyphen[e]) and not(db_hasHyphen[e+1]))
            if list01[e] == True:
                a = True
            list02[e] = a
        elif e==len(db)-1:
            list01[e] = (not(db_hasHyphen[e-1]) and (db_hasHyphen[e]) and False)
            if list01[e] == True:
                a = True
            list02[e] = a
        else:
            list01[e] = (not(db_hasHyphen[e-1]) and (db_hasHyphen[e]) and not(db_hasHyphen[e+1]))
            if list01[e] == True:
                a = True
            list02[e] = a

    if condition == 0:  # Get column name row
        return db.iloc[list01,0].values[0]
    elif condition == 1: # Get data rows
        return db.iloc[[x and y and z for x,y,z in zip(list02,db_hasHyphen.to_list(), [not elem for elem in (list01)])],0].to_list()
    elif condition == 2: # Get bool for rows are full of "-"
        return db_hasHyphen.to_list()
    elif condition == 3: # Get bool for columns names
        return list01
    elif condition == 4: # Get bool for columns names
        return list02
    else:
        print('Use a integer 0 to 1 plz')

# %%
def getMatrixVerticalBar(rawColum):
    havesep = [a for a in range(len(rawColum)) if rawColum[a]=='|']
    matrixsep = [[havesep[b],havesep[b+1]] for b in range(len(havesep)-1)]
    return matrixsep
    
def splitStringByMatrix(rawData,matrixhavesep):
    
    dataList = [[rowRawData[matrixhavesep[c][0]+1:matrixhavesep[c][1]].strip() for c in range(len(matrixhavesep))] for rowRawData in rawData]
    return dataList

def table2dataframe(db):
    rawColumn = rawData2listClean(db,0)
    rawData = rawData2listClean(db,1)
    matrixhavesep = getMatrixVerticalBar(rawColumn)

    dfColumName = [rawColumn[matrixhavesep[c][0]+1:matrixhavesep[c][1]].strip() for c in range(len(matrixhavesep))]
    dfData = splitStringByMatrix(rawData,matrixhavesep)

    df = pd.DataFrame(columns=dfColumName,data=dfData)

    return df #bd_df


#-----------------------------------------------------------

def createOT(claseOrden,ubicacionTecnica,equipo,centroPlanificacion,division,ordenModelo,session):
    session.findById(r"wnd[0]/tbar[0]/okcd").text = "/niw31"
    session.findById(r"wnd[0]/tbar[0]/btn[0]").press()
    # Setting the Ot class and eq to input
    session.findById(r"wnd[0]/usr/ctxtAUFPAR-PM_AUFART").text = claseOrden                                             # input clase de ot
    session.findById(r"wnd[0]/usr/subOBJECT:SAPLCOIH:7100/ctxtCAUFVD-TPLNR").text = ubicacionTecnica    # input a ubication to crate the ot
    session.findById(r"wnd[0]/usr/subOBJECT:SAPLCOIH:7100/ctxtCAUFVD-EQUNR").text = equipo                      # input a Equitment to create the ot
    session.findById(r"wnd[0]/usr/ctxtCAUFVD-IWERK").text = centroPlanificacion
    session.findById(r"wnd[0]/usr/ctxtCAUFVD-GSBER").text = division  
    session.findById(r"wnd[0]/usr/ctxtRC62C-REFNR").text = ordenModelo                                                    #"6041720" input a OT as a model
    session.findById(r"wnd[0]").sendVKey(0)
    session.findById(r"wnd[0]").sendVKey(0)
    #session.findById(r"wnd[1]/tbar[0]/btn[0]").press()

def completeOT(textoBreveOT,ptoTrabajoResponsable,claseActividadOrden,estadoInstalacion,fechaIniExtrema,fechaFinExtrema,prioridadOrden,revisionOrden,session):
# Completing the ot
    # General
        # input the title of the OT
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-KTEXT").text = textoBreveOT
    # Datos Cabecer
        # input Pto trabjo responsable
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-VAPLZ").text = ptoTrabajoResponsable
        # input clase de actividad de OT |PM1: Control/Inspección/Lubricación|
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").text = claseActividadOrden
        # input estado de instalación |0: Fuera de servicio|1: En funcionamiento|2: Parada de planta Parcial|3: Parada de Planta Total|
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ANLZU").text = estadoInstalacion
        # input prioridad  |1:Emergencia|2:Alta|3:Media|4:Baja|5:Muy baja|
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/cmbCAUFVD-PRIOK").key = prioridadOrden
        # input revisión
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-REVNR").text = revisionOrden
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GSTRP").text = fechaIniExtrema
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP").text = fechaFinExtrema

    #Previous to move to other tab
    session.findById(r"wnd[0]").sendVKey(0)
    session.findById(r"wnd[0]").sendVKey(0)
    #session.findById(r"wnd[0]").sendVKey(0)
    # if revisionOrden == "":
    #     session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
    #     session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
    # else:
    #     session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
    #     session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
    #     session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
       
def addOperationOT(ptoTrabajo,textoBreveOperacion,cantidad,duracion,destinatario,puestoDescarga,session,posicion=0):
    # Operaciones
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-ARBPL[2,{}]".format(posicion)).text = ptoTrabajo
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7,{}]".format(posicion)).text = textoBreveOperacion
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ANZZL[12,{}]".format(posicion)).text = cantidad
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-DAUNO[13,{}]".format(posicion)).text = duracion
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-WEMPF[17,{}]".format(posicion)).text = destinatario
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-ABLAD[18,{}]".format(posicion)).text = puestoDescarga

    session.findById(r"wnd[0]").sendVKey(0)
    session.findById(r"wnd[0]").sendVKey(0)
    session.findById(r"wnd[0]").sendVKey(0)

def addTextoExtendidoOP(textoBreveOperacion,textoExtendidoOP,session,posicion=0):
    # Get into texto textendido windows
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8,{}]".format(posicion)).setFocus()
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8,{}]".format(posicion)).press()

    # need a code to copy the textoExtendido and clear it
    if textoExtendidoOP == "":
        pyperclip.copy(textoExtendidoOP)
    else:
        pyperclip.copy(textoBreveOperacion+'\n'+ textoExtendidoOP)

    # Paste the text
    session.findById(r"wnd[0]/tbar[1]/btn[9]").press()
    # Return to Operation tab of OT
    session.findById(r"wnd[0]/tbar[0]/btn[3]").press()

def addMaterialesOP(material, cantidad, unidad, tipoPosicion, almacen, centro, posOrden, destinatario, puestoDescarga, reservaSolped, session, position=0):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1,{}]".format(position)).text = material
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-MENGE[4,{}]".format(position)).text = cantidad
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-EINHEIT[5,{}]".format(position)).text = unidad
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-POSTP[6,{}]".format(position)).text = tipoPosicion
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-LGORT[8,{}]".format(position)).text = almacen
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9,{}]".format(position)).text = centro
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10,{}]".format(position)).text = posOrden
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-WEMPF[12,{}]".format(position)).text = destinatario
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-ABLAD[13,{}]".format(position)).text = puestoDescarga
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/cmbRESBD-AUDISP[17,{}]".format(position)).key = reservaSolped

def pressEnter(session, number=1):
    for i in range(number):
        session.findById(r"wnd[0]").sendVKey(0)
    
def saveTx(session):
    session.findById(r"wnd[0]/tbar[0]/btn[11]").press()
def goToCabeceraTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ").select()
def goToOperationTabFromCabeceraTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select()
def goToOperationTabFromComponetesTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE").select()
def goToComponentesTabFromCabeceraTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB").select()
def goToObjetosTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpIOLU").select()
def goToDatosAdicionalesTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpIHKD").select()
def goToEmplazamientoTab(session):
    session.findById(r"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpILOA").select()


# This function will Login to SAP from the SAP Logon window
def createOTs(activities,resources,sapPath = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"):
    [p, q] = activities.shape
    activities = activities.replace(np.nan,'')
    resources = resources.replace(np.nan,'')
    resources = resources.drop(columns= {"correlativo","tag","revision","date plan","ot description","puesto responsable","task","specialist","craft","time","work"})
    activities.loc[:,"haveResources"] = activities.loc[:,"puesto descarga"].isin(resources.loc[:,"puesto_descarga"])
    
    try:
        path = sapPath
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

        for i in range(0,p):

            # if data.loc[i,'correlativo'] == data.loc[i+1,'correlativo']:
            if activities.loc[i,'position']  == 1:
                # print(data.loc[i,:])    
                # Variables before the creation of OT
                claseOrden = "ZM02"
                ubicacionTecnica = activities.loc[i,'ubt']
                equipo = activities.loc[i,'equipo']
                centroPlanificacion = "jp11"
                division = "jp11"
                ordenModelo = ""
                createOT(claseOrden,ubicacionTecnica,equipo,centroPlanificacion,division,ordenModelo,session)
                # Variable General information OT
                textoBreveOT = activities.loc[i,'ot description']
                # Variable Cabecera OT
                ptoTrabajoResponsable = activities.loc[i,'puesto responsable']
                claseActividadOrden = activities.loc[i,'activity']
                estadoInstalacion = activities.loc[i,'status eq']
                fechaIniExtremo = activities.loc[i,'date plan'].strftime("%d.%m.%Y")
                fechaFinExtremo = activities.loc[i,'date plan'].strftime("%d.%m.%Y")
                prioridadOrden = activities.loc[i,'prioridad']
                revisionOrden = activities.loc[i,'revision']
                completeOT(textoBreveOT,ptoTrabajoResponsable,claseActividadOrden,estadoInstalacion,fechaIniExtremo,fechaFinExtremo,prioridadOrden,revisionOrden,session)
                #Go to operation tab
                goToOperationTabFromCabeceraTab(session)

            ptoTrabajo = activities.loc[i,'specialist']
            textoBreveOperacion = activities.loc[i,'task']
            cantidadOp = str(int(activities.loc[i,'craft']))
            duracionOp = str(activities.loc[i,'time'])
            trabajoOp = str(activities.loc[i,'work'])
            position = activities.loc[i,'position']-1
            destinatario = ""
            puestoDescarga = str(activities.loc[i,'puesto descarga'])

            addOperationOT(ptoTrabajo,textoBreveOperacion,cantidadOp,duracionOp,destinatario,puestoDescarga,session,position)
            textoExtendidoOP = activities.loc[i,'detailed_task']
            textoBreveOperacion = activities.loc[i,'task']

            if textoExtendidoOP == "":
                pass
            else:
                addTextoExtendidoOP(textoBreveOperacion,textoExtendidoOP,session,position)

            if activities.loc[i,"haveResources"] == True:
                goToComponentesTabFromCabeceraTab(session)
                resourcesActivitie = resources.loc[resources.loc[:,"puesto_descarga"] == activities.loc[i,"puesto descarga"],:].reset_index()
                [m, n] = resourcesActivitie.shape
                for j in range(m):
                    material = resourcesActivitie.loc[j,"material"]
                    cantidadMaterial = resourcesActivitie.loc[j,"quantity"]
                    UnidadMaterial = resourcesActivitie.loc[j,"und"]
                    tipoPosicion = "L"
                    almacenMaterial = ""
                    centroMaterial = ""
                    posOrdenMaterial = resourcesActivitie.loc[j,"pos"]
                    destinatario = ""
                    puestoDescargaMaterial = resourcesActivitie.loc[j,"puesto_descarga"]
                    reservaSolpedMaterial = "3"
                    positionMaterial = resourcesActivitie.loc[j,"position"]-1
                    addMaterialesOP(
                        material, cantidadMaterial, UnidadMaterial, 
                        tipoPosicion, almacenMaterial, centroMaterial, posOrdenMaterial, 
                        destinatario, puestoDescargaMaterial, reservaSolpedMaterial, session, positionMaterial)
                pressEnter(session,m+1)       
                goToOperationTabFromComponetesTab(session)
            
            if p-1>i:
                if activities.loc[i+1,'position'] == 1:
                    saveTx(session)
            else:
                saveTx(session)

    except Exception as e:
        #print(sys.exc_info())
        logger.error(e, exc_info=True)
        print( activities.columns+ "\n" + activities.loc[i,:])
        pass
    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None

# GET DATA FROM SAP

# GET DATA OF MAINTENANCE TX

def gototx(session,tx,varian,SAPuser,layout,initialDate="01.01.2021",finalDate="31.12.2025"):
    
    session.findById(r"wnd[0]/tbar[0]/okcd").text = "/n" + tx # "IW29"
    session.findById(r"wnd[0]/tbar[0]/btn[0]").press()

    if tx == "iw29" or tx == "iw39" or tx == "iw37n":
        #session.findById(r"wnd[0]/tbar[0]/btn[0]").press()
        session.findById(r"wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById(r"wnd[1]/usr/txtV-LOW").text = varian  #"/BDIW29"
        session.findById(r"wnd[1]/usr/txtENAME-LOW").text = SAPuser #"JCALLOMAMANB"
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()
        #session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
        session.findById(r"wnd[0]").sendVKey(0)
    elif tx == "iwbk":
        #session.findById(r"wnd[0]/tbar[0]/btn[0]").press()
        session.findById(r"wnd[0]/usr/ctxtP_VARI").text = layout #layout
        # Adicionar UBT list
        session.findById(r"wnd[0]/usr/btn%_TPLNR_%_APP_%-VALU_PUSH").press()
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "JP11-SUL*"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "JP11-OXI*"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "JP11-INF*"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "JP11-MET*"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "JP11-ENE*"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "JP13-PRT*"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "JP11-MI1-TALL*"
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        # 
        session.findById(r"wnd[0]/usr/chkOFFEN").selected = True # OT Pendientes
        session.findById(r"wnd[0]/usr/chkINARB").selected = True # OT En tratamiento
        session.findById(r"wnd[0]/usr/chkABGES").selected = True # OT Concluido
        session.findById(r"wnd[0]/usr/chkHISTO").selected = True # OT Historico
        session.findById(r"wnd[0]/usr/ctxtAUFNR-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtMATNR-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtWERKS-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtLGORT-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtFTRMS-LOW").text = ""  
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()
        session.findById(r"wnd[1]/tbar[0]/btn[0]").press()

    elif tx == "kob1":
        # session.findById(r"wnd[0]/tbar[0]/okcd").text = "/n" + tx
        # session.findById(r"wnd[0]/tbar[0]/btn[0]").press()
        # Ingresa a la tx kob1
        try:

            session.findById(r"wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "mc01"
            session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        session.findById(r"wnd[0]/usr/ctxtP_DISVAR").text = layout  #"/BD2993"
        session.findById(r"wnd[0]/usr/btnBUT1").press()
        session.findById(r"wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999"
        session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
        # ingresar Periodo de busqueda
        session.findById(r"wnd[0]/usr/ctxtR_BUDAT-LOW").text = initialDate  #"01.01.2021"
        session.findById(r"wnd[0]/usr/ctxtR_BUDAT-HIGH").text = finalDate #"31.12.2022"
        # ingresar OTs  input
        session.findById(r"wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()
        session.findById(r"wnd[1]/tbar[0]/btn[6]").press()
        session.findById(r"wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB008").select()
        # Se ingresa a la tx iw39 y se parametriza OTs input
        session.findById(r"wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById(r"wnd[1]/usr/txtV-LOW").text = "/BDIW39"
        session.findById(r"wnd[1]/usr/txtENAME-LOW").text = "" #"JCALLOMAMANB" usuario sap
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        session.findById(r"wnd[0]/usr/ctxtVARIANT").text = "" # limpiar variante para rapidez
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()
        # Se selecciona todas las OTs como inputs
        session.findById(r"wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
        session.findById(r"wnd[0]/tbar[1]/btn[47]").press()
        # Aceptar la lista de OTs como input
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        # Se ejecuta la tx kob1
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press() 

    elif tx == "ksb1":
        # Ingresa a la tx ksb1
        try:
            session.findById(r"wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "mc01"
            session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass
        session.findById(r"wnd[0]/usr/ctxtP_DISVAR").text = layout #"/BD3611"
        # define 999999 lines
        session.findById(r"wnd[0]/usr/btnBUT1").press()
        session.findById(r"wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999"
        session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
        # ingresar Periodo de busqueda
        session.findById(r"wnd[0]/usr/ctxtR_BUDAT-LOW").text = initialDate #"01.01.2021"
        session.findById(r"wnd[0]/usr/ctxtR_BUDAT-HIGH").text = finalDate #"31.12.2022"
        # ingresar Grupo de CeCos
        session.findById(r"wnd[0]/usr/ctxtKSTGR").text = "2"

        # Se ejecuta la tx ksb1
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()

    elif tx == "mb25":
        #session.findById(r"wnd[0]//mbar/menu[2]/menu[0]/menu[0]").press()
        session.findById(r"wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press()
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "JP11"
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "" # "JP14"
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        session.findById(r"wnd[0]/usr/ctxtMATNR-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtBDTER-LOW").text = "01.01.2021"
        session.findById(r"wnd[0]/usr/ctxtBDTER-HIGH").text = "31.12.2025"
        #session.findById(r"wnd[0]/usr/txtUSNAM-LOW").text = ""
        # namelist = ["JCALLOMAMANB","JRODRIGUEZR","DMIRANDAF","HMACEDOD","NMORALESE","LALARCONC","JESPINOZAV","JFLORESV","JPARIONAH","TQUISPEC","VQUISPEC","APOLARG","AQUISPES","JVALENTINS","JQUISPET","CBARDALESC","NCARRASCOH"]
        namelist = ['*']
        pyperclip.copy('\r\n'.join(namelist))
        session.findById(r"wnd[0]/usr/btn%_USNAM_%_APP_%-VALU_PUSH").press()
        session.findById(r"wnd[1]/tbar[0]/btn[24]").press()
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()

        session.findById(r"wnd[0]/usr/txtWEMPF-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtKOSTL-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtAUFNR-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtPOSID-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtNPLNR-LOW").text = ""
        session.findById(r"wnd[0]/usr/txtP_VORNR").text = ""
        session.findById(r"wnd[0]/usr/ctxtANLN1-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtANLN2-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtUMWRK-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtUMLGO-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtKDAUF-LOW").text = ""
        session.findById(r"wnd[0]/usr/txtKDPOS-LOW").text = ""
        session.findById(r"wnd[0]/usr/txtKDEIN-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtALV_DEF").text = layout
        session.findById(r"wnd[0]/tbar[1]/btn[16]").press()
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").expandNode("        203")
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode("        209")
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "        203"
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode("        209")
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "          1"
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").text = "261"
        session.findById(r"wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
        session.findById(r"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "262"
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        # Se ejecuta la tx mb25
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()
    
    elif tx == "me5a":

        session.findById(r"wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById(r"wnd[1]/usr/txtV-LOW").text = "/BDME5A"
        session.findById(r"wnd[1]/usr/txtENAME-LOW").text = "JCALLOMAMANB"
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        # Se ejecuta la tx kob1
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press() 

        session.findById(r"wnd[0]/tbar[1]/btn[33]").press()
        session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    
    elif tx == "me2n":
        session.findById(r"wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById(r"wnd[1]/usr/txtV-LOW").text = varian
        session.findById(r"wnd[1]/usr/txtENAME-LOW").text = SAPuser
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        # Se ejecuta la tx kob1
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()
        session.findById(r"wnd[0]/mbar/menu[3]/menu[0]/menu[1]").select() # Repartos
        # #session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select() # Imputación
        # #session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[0]").select() # Lista basica
        
        # Se activa el layout
        session.findById(r"wnd[0]/tbar[1]/btn[33]").press()
        session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        #session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
        session.findById(r"wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    elif tx == "mb52":
        
        session.findById(r"wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById(r"wnd[1]/usr/txtV-LOW").text = varian
        session.findById(r"wnd[1]/usr/txtENAME-LOW").text = SAPuser
        session.findById(r"wnd[1]/tbar[0]/btn[8]").press()
        session.findById(r"wnd[0]/usr/ctxtMATNR-LOW").text = ""
        # session.findById(r"wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press() # Abrir material masivo
        # session.findById(r"wnd[1]/tbar[0]/btn[24]").press() # Pegar lista de materiales
        # session.findById(r"wnd[1]/tbar[0]/btn[8]").press() # Aceptar lista masiva de materiales
        session.findById(r"wnd[0]/usr/ctxtLGORT-LOW").text = ""
        session.findById(r"wnd[0]/usr/ctxtP_VARI").text = layout
        
        session.findById(r"wnd[0]/tbar[1]/btn[8]").press()
    else:
        print("The transaction {} , has no support by now").format(tx)

def gettxdata(session,tx,dataname,rute,code):
    if tx == "iw29" or tx == "iw39":
        # get local file IW29
        session.findById(r"wnd[0]/mbar/menu[0]/menu[11]/menu[2]").select()
    elif  tx == "iw37n":
        # get local file IW37n
        session.findById(r"wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select()
    elif tx == "kob1" or tx == "ksb1" or tx == "iwbk" or tx=="me5a" or tx=="me2n":
        session.findById(r"wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    elif tx == "mb25" or tx=="mb52":
        session.findById(r"wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()

    session.findById(r"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
    session.findById(r"wnd[1]/usr/ctxtDY_PATH").text = rute
    session.findById(r"wnd[1]/usr/ctxtDY_FILENAME").text = dataname
    session.findById(r"wnd[1]/usr/ctxtDY_FILE_ENCODING").text = code
    # Crear file
    #session.findById(r"wnd[1]/tbar[0]/btn[0]").press()
    # Reemplazar file
    session.findById(r"wnd[1]/tbar[0]/btn[11]").press()
    # Ampliar file
    #session.findById(r"wnd[1]/tbar[0]/btn[7]").press()

# GET DATA OF MAINTENANCE TX

