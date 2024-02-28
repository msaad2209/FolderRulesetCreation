import shutil
import os
import openpyxl
import sys
import ctypes

# using getlogin() returning username
UserName = os.getlogin()
UserEmail = UserName + '@sermo.com'

#Take input for project details
ProjectNumber = input(f"Enter Project Pcode(4 digit for SaaS, 6 digit for CF): ")
if ProjectNumber == "":
    ctypes.windll.user32.MessageBoxW(0, "There should be value for project number", "Project Number", 16)
    print("There should be value for project number")
    sys.exit()

ClientName = input(f"Enter client name: ")
ProjectName = input(f"Enter Project Name: ")
ProjectCode = input(f"Enter Project Confirmit Code: ")
templateid =input(f"Enter Project template id: ")

#Take input for project details
CFSaas = input(f"Is it the RTP project, Enter y if SaaS and n if CF (y/n): ")
if CFSaas == "":
    ctypes.windll.user32.MessageBoxW(0, "You should specify either SaaS or CF", "SaaS or CF", 16)
    print("You should specify either SaaS or CF")
    sys.exit()


#Folder Access
if 'swati' in UserName:
    BeforeUserName = 'C:/Users/'
    AfterUserName = '/OneDrive - Sermo/Projects/'
    Tempfolder = 'Template - RTP projects'
    Tempfolder2 = 'Template - RTP projects2'
    Mainfolderpath = BeforeUserName + UserName + AfterUserName
else:
    BeforeUserName = 'C:/Users/'
    AfterUserName = '/Sermo/Operations - Data Processing/Projects/'
    Tempfolder = 'Template - RTP projects'
    Tempfolder2 = 'Template - RTP projects2'
    Mainfolderpath = BeforeUserName + UserName + AfterUserName

new_folder_name = ProjectNumber + " - " + ClientName+ " - " + ProjectName

# set the source and destination folder paths
source_folder = BeforeUserName + UserName + AfterUserName + Tempfolder

destination_folder = BeforeUserName + UserName + AfterUserName + new_folder_name
print(destination_folder)
# check if the destination folder already exists
if os.path.exists(destination_folder):
    ctypes.windll.user32.MessageBoxW(0, "This folder is already exist, delete and then recreate it / or try with other name", "Folder already exist", 16)
    print("This folder is already exist, delete and then recreate it / or try with other name")
    sys.exit()
    # if it does, delete it
    #shutil.rmtree(destination_folder)
else:
    # make a copy of the source folder in the destination folder
    shutil.copytree(source_folder, destination_folder)


# get the full path of the file you want to update
file_path = Mainfolderpath + new_folder_name + '/B. DP/b. scripts/'
fileName = '!Delivery_manager.xlsm'

# Load the Excel file
workbook = openpyxl.load_workbook(file_path+fileName , data_only=False, keep_vba=True)
sh = workbook["Parameters"]
# Select the cell to fill with text
cellB2 = sh['B2'] 
cellB3 = sh['B3']
cellB4 = sh['B4']
cellB7 = sh['B7']
cellB8 = sh['B8']
cellB9 = sh['B9']
cellB16 = sh['B16']
cellB17 = sh['B17']

ncode = int(ProjectNumber)
cellB2.value = ProjectNumber
cellB3.value = ProjectName
cellB4.value = ClientName

if CFSaas == 'n' or CFSaas == 'N':
    cellB7.value = '<ProjectCode>_SurveyData.xlsx'
    cellB8.value = '<ProjectCode>_SurveyData.sav'
    cellB9.value = '<ProjectCode>_SurveyData_responseid.asc'
    cellB16.value = 'ccode=Country'
    cellB17.value = 'responseid, respid, pin, pass'
    sh2 = workbook["Operations"]
    # Select the cell to fill with text
    sh2['C2'] = 'TRUE'
    sh2['C3'] = 'FALSE'
    sh2['D2'] = 'TRUE'
    sh2['D3'] = 'FALSE'
    sh2['E2'] = 'TRUE'
    sh2['E3'] = 'FALSE'
    sh2['F2'] = 'TRUE'
    sh2['F3'] = 'FALSE'

#Save the Excel macro files
#UpdatedExcel = file_path+fileName
#print(UpdatedExcel)
workbook.save(file_path+fileName)

# get the new name of the file
new_name = file_path + '!' + ProjectNumber +' Delivery_manager.xlsm'

# rename the file
os.rename(file_path+fileName, new_name)


# ------- Rule Creator---------------

if ProjectCode == "" or templateid == "":
    ctypes.windll.user32.MessageBoxW(0, "Foldr created, rules not created as all the required information not provided", "Successful", 0)
    print("Foldr created, Rules not created as all the required information not provided")
    sys.exit()
else:
    import xml.etree.ElementTree as ET
    from datetime import datetime


    # create the file structure
    def create_rule(pcode, project_code, template_id, frmt, my_email, my_part, forsta):
        panelRule = ET.Element('PanelRule')
        panelRule.set('xmlns:xsd', "http://www.w3.org/2001/XMLSchema")
        panelRule.set('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance")

        panelId = ET.SubElement(panelRule, 'PanelId')
        panelId.text = "-1"
        created = ET.SubElement(panelRule, 'Created')
        created.text = "2016-10-21T09:36:13.343"
        lastUpdated = ET.SubElement(panelRule, 'LastUpdated')
        lastUpdated.text = "2016-10-21T09:42:51.837"
        createdBy = ET.SubElement(panelRule, 'CreatedBy')
        createdBy.text = "SERMO"
        lastUpdated = ET.SubElement(panelRule, 'LastUpdated')
        lastUpdated.text = "SERMO"
        propertyValues = ET.SubElement(panelRule, 'PropertyValues')
        ownerUserName = ET.SubElement(panelRule, 'OwnerUserName')
        ruleId = ET.SubElement(panelRule, 'RuleId')
        ruleId.text = "111111"
        ruleName = ET.SubElement(panelRule, 'RuleName')
        ruleName.text = "{}_{}".format(pcode,frmt)
        isTemporary = ET.SubElement(panelRule, 'IsTemporary')
        isTemporary.text='false'
        lastExecutedBy = ET.SubElement(panelRule, 'LastExecutedBy')
        lastExecuted = ET.SubElement(panelRule, 'LastExecuted')
        lastExecuted.text = '0001-01-01T00:00:00'
        companyId = ET.SubElement(panelRule, 'CompanyId')
        companyId.text = "2"
        status = ET.SubElement(panelRule, 'Status')
        status.text = 'Enabled'

        conditionExpression = ET.SubElement(panelRule, 'ConditionExpression')
        if forsta==True:
            conditionExpression.text = 'isTest = "0" AND NOT ISNULL(rdout) AND NOT IN(respondentStatus, \"DUPLICATE\", \"RESET\", \"REMOVE\", \"RECO\")'
        elif forsta==False:
            conditionExpression.text = 'xtest = "0" AND (VRFD = "1" OR ISNULL(VRFD)) AND NOT ISNULL(rdout) AND NOT IN(respondentStatus, \"DUPLICATE\", \"RESET\", \"REMOVE\", \"RECO\")'

        variables = ET.SubElement(panelRule, 'Variables')
        action = ET.SubElement(panelRule, 'Action')
        loopActions = ET.SubElement(panelRule, 'LoopActions')
        LoopIds = ET.SubElement(panelRule, "LoopIds")
        globalScript = ET.SubElement(panelRule, "GlobalScript")
        globalProperties = ET.SubElement(panelRule, "GlobalProperties")
        postScript = ET.SubElement(panelRule, "PostScript")
        comment = ET.SubElement(panelRule, "Comment")
        #comment.text = frmt
        if forsta==True:
            comment.text = "{}_{}_SurveyData{}_sFTP".format(pcode,frmt,my_part)

        if forsta==False:
            comment.text = "{}_{}_SurveyData{}".format(pcode,frmt,my_part)

        selectedCount = ET.SubElement(panelRule, "SelectedCount")
        selectedCount.text = "0"
        qualifiedCount = ET.SubElement(panelRule, "QualifiedCount")
        qualifiedCount.text = "0"
        updatedCount = ET.SubElement(panelRule, "UpdatedCount")
        updatedCount.text = "0"
        currentRuleTaskId = ET.SubElement(panelRule, "CurrentRuleTaskId")
        currentRuleTaskId.text = "-1"
        isAdHoc = ET.SubElement(panelRule, "IsAdHoc")
        isAdHoc.text = "false"
        
        targetSettings = ET.SubElement(panelRule, "TargetSettings")
        if frmt == 'Excel':
            targetSettings.set('xsi:type', "ExcelTargetSettings")
        elif frmt == 'SPSS':
            targetSettings.set('xsi:type', "SavTargetSettings")
        elif frmt == 'ASCII':
            targetSettings.set('xsi:type', "SssDataTargetSettings")
            
        mappedFields = ET.SubElement(targetSettings, "MappedFields")
        fileName = ET.SubElement(targetSettings, "FileName")
        #fileName.text = "{}_SurveyData{}".format(pcode,my_part)

        #2022-11-22: Igoris added
        if forsta==True:
            if frmt == 'Excel':
                fileName.text = "{}_SurveyData{}_excel".format(pcode,my_part)
            elif frmt == 'SPSS':
                fileName.text = "{}_SurveyData{}_spss".format(pcode,my_part)
            elif frmt == 'ASCII':
                fileName.text = "{}_SurveyData{}_ascii".format(pcode,my_part)
        if forsta==False:
            fileName.text = "{}_SurveyData{}".format(pcode,my_part)

        #2022-11-22: Igoris added
        if forsta==True:
            Uncompressed = ET.SubElement(targetSettings, "Uncompressed")
            Uncompressed.text = "true"

        emailOptions = ET.SubElement(targetSettings, "EmailOptions")
        replyTo = ET.SubElement(emailOptions, "ReplyTo")
        subject = ET.SubElement(emailOptions, "Subject")
        encodingCodePage = ET.SubElement(emailOptions, "EncodingCodePage")
        encodingCodePage.set('xsi:nil',"true")
        HtmlBody = ET.SubElement(emailOptions, 'HtmlBody')
        PlainTextBody = ET.SubElement(emailOptions, 'PlainTextBody')

        emailAddress = ET.SubElement(targetSettings, "EmailAddress")
        emailAddress.text = my_email
        overrideFileName = ET.SubElement(targetSettings, "OverrideFileName")
        overrideFileName.text = "true"
        fileTransferType = ET.SubElement(targetSettings, "FileTransferType")
        if forsta==True:
            #fileTransferType.text = "Email"
            fileTransferType.text = "ExternalFtpServer"
        elif forsta==False:
            fileTransferType.text = "FtpServer"
        encryptFile = ET.SubElement(targetSettings, "EncryptFile")
        encryptFile.text = "false"
        useInternally = ET.SubElement(targetSettings, "UseInternally")
        useInternally.text = "false"
        openTextWidth = ET.SubElement(targetSettings, "OpenTextWidth")
        if frmt=="Excel" or frmt=='ASCII':
            openTextWidth.text = "-1"
        elif frmt=='SPSS':
            openTextWidth.text = "200"
        loopHandling = ET.SubElement(targetSettings, "LoopHandling")
        loopHandling.text = "SeparateFiles"
        loopPosition = ET.SubElement(targetSettings, "LoopPosition")
        loopPosition.text = "AsQuestionnaire"
        if forsta==True:
            externalFtpTargetParameters = ET.SubElement(targetSettings, "ExternalFtpTargetParameters")
            externalFtpTargetParameters.set("FolderName", "Download")
            externalFtpTargetParameters.set("HostName", "share.sermo.com")
            externalFtpTargetParameters.set("UserName", "w7o3MYw9ECcoD7ZWAUUbadApgYQvH2NR/lHshG5lrdE=")
            externalFtpTargetParameters.set("Password", "5uojQn2LKas9OJQhBmw1hw==")
            externalFtpTargetParameters.set("UseSsh", "true")
            externalFtpTargetParameters.set("HostFingerprint", "c7:e1:28:09:2a:b3:e5:df:23:70:dc:2e:aa:bd:c9:74")
            externalFtpTargetParameters.set("AlwaysTrustThisHost", "false")
        elif forsta==False:
            externalFtpTargetParameters = ET.SubElement(targetSettings, "ExternalFtpTargetParameters")
            externalFtpTargetParameters.set("AlwaysTrustThisHost", "false")
            externalFtpTargetParameters.set("HostFingerprint", "")
            externalFtpTargetParameters.set("UseSsh", "false")
            externalFtpTargetParameters.set("Password", "CCGYoC3KgCySSEmApXHqOA==")
            externalFtpTargetParameters.set("UserName", "4BxDU5hon8xwv2+CdvRLAQ==")
            externalFtpTargetParameters.set("HostName", "")
            externalFtpTargetParameters.set("FolderName", "")
        
        
        #recodeMultis = ET.SubElement(targetSettings, "RecodeMultis")
        #recodeMultis.text = "false" !!not in excel xml

        if frmt=='Excel':
            excelVersion = ET.SubElement(targetSettings, "ExcelVersion")
            excelVersion.text = "MsExcel2007"
        if frmt=='ASCII':
            SssTemplateId = ET.SubElement(targetSettings, "SssTemplateId")
            SssTemplateId.text = template_id
            IncludeSchema = ET.SubElement(targetSettings, "IncludeSchema")
            IncludeSchema.text = "true"
            IncludeData = ET.SubElement(targetSettings, "IncludeData")
            IncludeData.text = "true"
            CodePage = ET.SubElement(targetSettings, "CodePage")
            CodePage.text = "28591"
            CodePageForSchema = ET.SubElement(targetSettings, "CodePageForSchema")
            CodePageForSchema.text = "65001"
            IncludeConfirmitMetaTags = ET.SubElement(targetSettings, "IncludeConfirmitMetaTags")
            IncludeConfirmitMetaTags.text = "false"
            
        sourceSettings = ET.SubElement(panelRule, "SourceSettings")
        sourceSettings.set("xsi:type", "SurveyDataSourceSettings")
        projectIds = ET.SubElement(sourceSettings, "ProjectIds")
        string_projectcode = ET.SubElement(projectIds, "string")
        string_projectcode.text = project_code

        databaseType = ET.SubElement(sourceSettings, "DatabaseType")
        databaseType.text = "Production"
        ruleDateFilterType = ET.SubElement(sourceSettings, "RuleDateFilterType")
        ruleDateFilterType.text = "None"
        isIncrementalUpdate = ET.SubElement(sourceSettings, "IsIncrementalUpdate")
        isIncrementalUpdate.text = "false"
        lastChangeTrackingVersion = ET.SubElement(sourceSettings, "LastChangeTrackingVersion")
        lastChangeTrackingVersion.text = "0"
        keywordFilter = ET.SubElement(sourceSettings, "KeywordFilter")
        rowLimit = ET.SubElement(sourceSettings, "RowLimit")
        rowLimit.text = "0"
        responseFilter = ET.SubElement(sourceSettings, "ResponseFilter")
        responseStatus = ET.SubElement(responseFilter, "ResponseStatus")       
        responseStatus.text = "Complete"
        filterTemplateId = ET.SubElement(sourceSettings, "FilterTemplateId")
        exportFieldLabelSourceType = ET.SubElement(sourceSettings, "ExportFieldLabelSourceType")
        
        if frmt=='Excel':
            filterTemplateId.text = template_id
            exportFieldLabelSourceType.text = "Project"
        elif frmt=='SPSS':
            filterTemplateId.text = template_id
            exportFieldLabelSourceType.text = "Template"
        elif frmt=='ASCII':
            filterTemplateId.text = "0"
            exportFieldLabelSourceType.text = "Project"

        hideTemplate = ET.SubElement(sourceSettings, "HideTemplate")
        hideTemplate.text = "false"
        allowFilterVarsNotExistInDb = ET.SubElement(sourceSettings, "AllowFilterVarsNotExistInDb")
        allowFilterVarsNotExistInDb.text = "false"
        showVarNotExistWarning = ET.SubElement(sourceSettings, "ShowVarNotExistWarning")
        showVarNotExistWarning.text = "false"
        allowFilterAnswersNotExistInDb = ET.SubElement(sourceSettings, "AllowFilterAnswersNotExistInDb")
        allowFilterAnswersNotExistInDb.text = "false"
        showAnswerNotExistWarning = ET.SubElement(sourceSettings, "ShowAnswerNotExistWarning")
        showAnswerNotExistWarning.text = "false"
        AddLabels = ET.SubElement(sourceSettings, "AddLabels")
        AddLabels.text = "false"
        labelLanguage = ET.SubElement(sourceSettings, "LabelLanguage")
        labelLanguage.text = "9"
        labelType = ET.SubElement(sourceSettings, "LabelType")
        labelType.text = "QuestionId"
        questionElementDescriptionType = ET.SubElement(sourceSettings, "QuestionElementDescriptionType")
        questionElementDescriptionType.text = "AnswerQuestionLabel"
        openEndHandling = ET.SubElement(sourceSettings, "OpenEndHandling")
        openEndHandling.text = "IncludeOpenEnds"
        
        source = ET.SubElement(panelRule, "Source")
        source_typeId = ET.SubElement(source, "TypeId")
        source_typeId.text = "SurveyData"
        source_text = ET.SubElement(source, "Text")
        source_text.text = "Survey Database"

        
        target = ET.SubElement(panelRule, "Target")
        target_typeId = ET.SubElement(target, "TypeId")
        target_text = ET.SubElement(target, "Text")
        if frmt=='Excel':
            target_typeId.text = "Excel"
            target_text.text = "Excel File"
        elif frmt=='SPSS':
            target_typeId.text = "SpssSav"
            target_text.text = "SPSS File (sav)"
        elif frmt=='ASCII':
            target_typeId.text = "SssDataFile"
            target_text.text = "Triple-S XML (Standard)"
            
        ruleType = ET.SubElement(panelRule, "RuleType")
        ruleType.text = "Normal"
        
        sTVS = ET.SubElement(panelRule, "SourceAndTargetValidationSettings")
        allowVarInSourceNotInTarget = ET.SubElement(sTVS, "AllowVarInSourceNotInTarget")
        allowVarInSourceNotInTarget.text = "false"
        showVarInSourceNotInTargetWarning = ET.SubElement(sTVS, "ShowVarInSourceNotInTargetWarning")
        showVarInSourceNotInTargetWarning.text = "false"
        allowVarInTargetNotInSource = ET.SubElement(sTVS, "AllowVarInTargetNotInSource")
        allowVarInTargetNotInSource.text = "false"
        showVarInTargetNotInSourceWarning = ET.SubElement(sTVS, "ShowVarInTargetNotInSourceWarning")
        showVarInTargetNotInSourceWarning.text = "false"
        allowAnswerInSourceNotInTarget = ET.SubElement(sTVS, "AllowAnswerInSourceNotInTarget")
        allowAnswerInSourceNotInTarget.text = "false"
        showAnswerInSourceNotInTargetWarning = ET.SubElement(sTVS, "ShowAnswerInSourceNotInTargetWarning")
        showAnswerInSourceNotInTargetWarning.text = "false"
        allowAnswerInTargetNotInSource = ET.SubElement(sTVS, "AllowAnswerInTargetNotInSource")
        allowAnswerInTargetNotInSource.text = "false"
        showAnswerInTargetNotInSourceWarning = ET.SubElement(sTVS, "ShowAnswerInTargetNotInSourceWarning")
        showAnswerInTargetNotInSourceWarning.text = "false"

        dCPD = ET.SubElement(panelRule, "DataCentralProjectDescription")
        dCPD.set('DesktopVersionOnly', "false")
        dCPD.set('Version', "1")
        dCPD.set('Id', "00000000-0000-0000-0000-000000000000")

        inputSurveys = ET.SubElement(dCPD, "InputSurveys")
        outputSurveys = ET.SubElement(dCPD, "OutputSurveys")
        messages_surveys = ET.SubElement(dCPD, "Messages")
        logs_surveys = ET.SubElement(dCPD, "Logs")
        reports_surveys = ET.SubElement(dCPD, "Reports")

        dataCentralProjectId = ET.SubElement(panelRule, "DataCentralProjectId")
        dataCentralProjectId.text = "0"
        dataCentralProjectLastUpdated = ET.SubElement(panelRule, "DataCentralProjectLastUpdated")
        dataCentralProjectLastUpdated.text = "0001-01-01T00:00:00"

        # create a new XML file with the results
        mydata = ET.tostring(panelRule)
        if forsta == True:
            myfile = open("{}_{}_SurveyData{}_sFTP.xml".format(pcode,frmt,my_part), "wb")
        if forsta == False:
            myfile = open("{}_{}_SurveyData{}.xml".format(pcode,frmt,my_part), "wb")
        myfile.write(mydata)


    #### !!!!!!!!!!!!!!!
    #username, password in ExternalFtpTargetParameters to be either deleted or updated individually




    import os
    import sys
    os.chdir(Mainfolderpath+new_folder_name+'/B. DP/b. scripts/DataMapLayout/Rules')

    #from xml_structure_script import create_rule

    if CFSaas == 'n' or CFSaas == 'N':
        forsta = False
    else:
        forsta = True
    project_number = ProjectNumber 
    confirmit_pcode = ProjectCode
    part_of_study = ""            #Specify Sufix for multiple sets (for exaple if you have HCP and PATS, write _HCP or _PATS - one run at a time)
    template_id = templateid
    ASCII = True
    my_email = UserEmail

    for f in ['Excel', 'SPSS']:
        create_rule(pcode = project_number, project_code = confirmit_pcode, template_id = template_id, frmt=f, my_email=my_email, my_part = part_of_study, forsta=forsta)
        if ASCII==True:
            create_rule(pcode = project_number, project_code = confirmit_pcode, template_id = template_id, frmt='ASCII', my_email=my_email, my_part = part_of_study, forsta=forsta)

    print("Foldr and Rules are created")
    ctypes.windll.user32.MessageBoxW(0, "Foldr and Rules are created", "Successful", 0)
    sys.exit()
    








