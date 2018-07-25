import gspread
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import os
from apiclient.http import MediaFileUpload
#from apiclient import errors
import unicodedata
import sys
from lxml import etree as et
import requests
import time



#function to update names of jpeg and dng files
def change_names(currToken):
    results = service.files().list(pageSize=1000, pageToken = currToken, fields="nextPageToken, files(id, name,mimeType)").execute()
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        print('Filename (File ID)')
        for item in items:
            if item['name'][-4:] == '.jpg' and '.' in item['name'][:-4]:
                new_name = item['name'][:-4]
                new_name = new_name.replace('.','_')
                new_name += '.jpg'
                id = item['id']
                body = {'name' : new_name}
                file = service.files().get(fileId=id).execute()
                service.files().update(fileId = id, body = body).execute()
            print('{0} ({1})'.format(item['name'].encode('utf-8'), item['id']))
        print('Total=', len(items))
        currToken = results.get('nextPageToken', None)
        if currToken != None:
            listfiles(currToken)

# change_names('')

def create_xml(sheet,row):
    time.sleep(10)
    # prefixes
    mods = "http://www.loc.gov/mods/v3"
    xsi = "http://www.w3.org/2001/XMLSchema-instance"
    NS_map = {"mods": mods, "xsi": xsi}

    # register namespace prefix
    et.register_namespace('mods', mods)
    et.register_namespace('xsi', xsi)

    rootName = et.QName(mods, 'mods')
    root = et.Element(rootName, nsmap=NS_map)
    root.set('ID', 'enterID')
    root.set('{' + xsi + '}schemaLocation',
             "http://www.loc.gov/mods/v3 http://www.loc.gov/mods/v3/mods-3-7.xsd")
    tree = et.ElementTree(root)

    #file name
    cell = sheet.cell(row,1)
    file_name = cell.value.split('.dng')
    file_name = file_name[0].split('.jpg')
    file_name = file_name[0].replace('.','_')

    #title
    titleInfo = et.SubElement(root, et.QName(mods, "titleInfo"))
    title = et.SubElement(titleInfo, et.QName(mods, "title"))
    title.text = sheet.cell(row,3).value

    #creator
    name_author = et.SubElement(root, et.QName(mods, "name"), type="personal")
    namePart_author = et.SubElement(name_author, et.QName(mods, "namePart"))
    namePart_author.text = "Doumani, Beshara"
    role_author = et.SubElement(name_author, et.QName(mods, "role"))
    roleTerm_author = et.SubElement(role_author, et.QName(mods, "roleTerm"), type="text", authority="marcrelator")
    roleTerm_author.text = "creator"

    #type of resource
    type = et.SubElement(root, et.QName(mods, "typeOfResource"))
    type.text = "text"

    # genre
    genre = et.SubElement(root, et.QName(mods, "genre"), authority="aat")
    genre.text = "court records"

    #date created
    origin = et.SubElement(root, et.QName(mods, "originInfo"))
    date1 = et.SubElement(origin, et.QName(mods, "dateCreated"))
    span = sheet.cell(2,4).value
    date1.text = sheet.cell(2,4).value
    start = span.split('-')[0]
    end = span.split('-')[1]
    date2 = et.SubElement(origin, et.QName(mods, "dateCreated"), keyDate = "yes", encoding="w3cdtf", point="start")
    date2.text = start
    date2 = et.SubElement(origin, et.QName(mods, "dateCreated"), encoding="w3cdtf", point="end")
    date2.text = end

    #abstract
    abstract = et.SubElement(root, et.QName(mods, "abstract"))
    cell = sheet.cell(2,5)
    cell.value = unicodedata.normalize('NFKD', cell.value).encode('ascii','ignore')
    abstract.text = cell.value

    #register
    part_register = et.SubElement(root, et.QName(mods, "part"))
    detail = et.SubElement(part_register, et.QName(mods, "detail"), type = "register")
    number = et.SubElement(detail, et.QName(mods, "number"))
    number.text = sheet.cell(row, 6).value

    #key words
    subject1 = et.SubElement(root, et.QName(mods, "subject"), authority = "lcsh")
    word1 = et.SubElement(subject1, et.QName(mods, "geographic"))
    word1.text = "Tripoli (Libya)"

    subject2 = et.SubElement(root, et.QName(mods, "subject"), authority = "local")
    word2 = et.SubElement(subject2, et.QName(mods, "geographic"))
    word2.text = "Ottoman Palestine"

    subject3 = et.SubElement(root, et.QName(mods, "subject"), authority = "lcsh")
    word3 = et.SubElement(subject3, et.QName(mods, "geographic"))
    word3.text = "Syria"

    subject4 = et.SubElement(root, et.QName(mods, "subject"), authority = "lcsh")
    word4 = et.SubElement(subject4, et.QName(mods, "topic"))
    word4.text = "Court records"

    subject5 = et.SubElement(root, et.QName(mods, "subject"), authority = "local")
    word5 = et.SubElement(subject5, et.QName(mods, "topic"))
    word5.text = "Shar'iyya"

    subject6 = et.SubElement(root, et.QName(mods, "subject"), authority = "local")
    word6 = et.SubElement(subject6, et.QName(mods, "topic"))
    word6.text = "Sijills"

    #license
    xlink = "http://www.w3.org/1999/xlink"
    NS_map = {"xlink": xlink}

    # register namespace prefix
    et.register_namespace('xlink', xlink)
    #et.register_namespace('xsi', xsi)

    access1 = et.SubElement(root, et.QName(mods, 'AccessCondition'), nsmap = NS_map, type = "use and reproduction")
    access1.set('{' + xlink + '}href', 'https://creativecommons.org/licenses/by-nc/4.0')
    access1.text = 'This work is licensed under a Creative Commons Attribution-NonCommercial 4.0 International license'

    access2 = et.SubElement(root, et.QName(mods, 'AccessCondition'), nsmap = NS_map, type = "logo")
    access2.set('{' + xlink + '}href', "https://licensebuttons.net/l/by-nc/4.0/88x31.png")


    xml = tree.write(file_name + '.xml', encoding="utf-8", pretty_print=True)
    return file_name +'.xml'

def createRemoteFolder(folderName, parentID):
        # Create a folder on Drive, returns the newely created folders ID
        body = {
          'name': folderName,
          'mimeType': "application/vnd.google-apps.folder"
        }
        if parentID:
            body['parents'] = [parentID]
        root_folder = service.files().create(body = body).execute()
        return root_folder['id']

def get_foldernames():
    results = service.files().list(q = "parents in '1KTBfAzQ-XtuoRFCQbFQ7RmUC4OdgRjvZ'", fields="nextPageToken, files(id, name,mimeType)").execute()
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        print('Filename (File ID)')
        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                print('{0} ({1})'.format(item['name'].encode('utf-8'), item['id']))
    print('Total=', len(items))


def callback(request_id, response, exception):
    if exception:
        # Handle error
        print exception
    else:
        print "Permission Id: %s" % response.get('id')

def give_permission():
    batch = service.new_batch_http_request(callback=callback)
    user_permission = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': 'jt0898developer@gmail.com'
    }
    batch.add(service.permissions().create(
            fileId=file_id,
            body=user_permission,
            fields='id',
    ))
    batch.execute()

def uploadFilesToDrive(credentials):
    stop = False
    gc = gspread.authorize(credentials)
    wks = gc.open('Metadata Islamic Court Records of Tripoli Volumes 1 to 32').worksheets()
    for sheet in wks:
        if not stop:
            title = str(sheet.title)
            folder_id = createRemoteFolder(title, '1KTBfAzQ-XtuoRFCQbFQ7RmUC4OdgRjvZ')
            num_rows = (len(sheet.get_all_values()))
            for row in range(2,num_rows):
                if not stop:
                    time.sleep(1)
                    file_name = create_xml(sheet,row)
                    file_metadata = {'name': file_name, 'parents': [folder_id]}
                    media = MediaFileUpload(file_name,mimetype='text/xml')
                    file = service.files().create(body=file_metadata,media_body=media,fields='id').execute()
                    print(file_name)
                    os.remove(file_name)
                    if row == num_rows:
                        stop = True

scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
credentials = ServiceAccountCredentials.from_json_keyfile_name('Doumani-1bbad2f463ef.json',scope)

service = build('drive', 'v3', http = credentials.authorize(Http()))
uploadFilesToDrive(credentials)
