import os
import os.path
import pickle
import shutil
from string import Template

import pandas as pd
from PIL import Image
from cairosvg import svg2png
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


def gsheet_api_check(SCOPES):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def pull_sheet_data(SCOPES, SPREADSHEET_ID, RANGE_NAME):
    creds = gsheet_api_check(SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                  range=RANGE_NAME).execute()
        data = rows.get('values')
        print("COMPLETE: Data copied")
        return data


def load_data():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SPREADSHEET_ID = '1qzCZfRZUJ8iXuFJ2Ka74mxH0eIeJu-lPCC-jfvYvxBE'
    RANGE_NAME = 'A1:C16'
    data = pull_sheet_data(SCOPES, SPREADSHEET_ID, RANGE_NAME)
    return pd.DataFrame(data[1:], columns=data[0])


def getSVG(filemane):
    file = open(filemane, 'r')
    data = file.readlines()
    file.close()
    data = ''.join(data)
    return Template(data)


def parseName(name):
    res = ''
    for letter in name:
        res += f'&#{ord(letter)};'
    return res


def makeBackground():
    head = Image.open('pictures\\samples\\head.png')
    center = Image.open('pictures\\samples\\center.png')
    bottom = Image.open('pictures\\samples\\bottom.png')

    head_w, head_h = head.size
    bg_w, bg_h = head_w, head_h

    center_w, center_h = center.size
    for i in range(16):
        bg_h += center_h

    bottom_w, bottom_h = bottom.size
    bg_h += bottom_h

    bg = Image.new('RGBA', (bg_w, bg_h), (255, 255, 255, 255))

    bg.paste(head, (0, 0))
    height = head_h
    for i in range(15):
        new_center = Image.open(f'pictures\\temp\\centers\\center{i + 1}.png')
        bg.paste(new_center, (0, height))
        height += center_h
    bg.paste(center, (0, height))
    height += center_h
    bg.paste(bottom, (0, height))

    bg.save('pictures\\output\\result.png')


def makeCenters(data):
    svgName = getSVG('svgName.xml')
    svgNumber = getSVG('svgNumber.xml')
    center = Image.open('pictures\\samples\\center.png')
    center_w, center_h = center.size
    colors = {1: '#E9CC59', 2: '#E8E8E8', 3: '#C39D4F', 4: '#07CCF7'}
    rank = 1
    dName = {'width': str(center_w), 'height': str(center_h)}
    dNumber = {'width': str(center_w), 'height': str(center_h)}

    for ind in range(15):
        new_center = center.copy()
        if rank < 4 and data['Количество баллов'].iloc[ind] < data['Количество баллов'].iloc[ind - 1]:
            rank += 1

        name = makeName(svgName, data['Участник'].iloc[ind], colors[rank], ind, dName)
        number = makeNumber(svgNumber, str(data['Количество баллов'].iloc[ind]) + ' ', colors[rank], ind, dNumber)
        new_center.paste(name, (365, 0), name)
        new_center.paste(number, (1037, 0), number)
        new_center.save(f"pictures\\temp\\centers\\center{ind + 1}.png")


def makeName(svgName, name, color, index, template):
    template['color'] = color
    template['name'] = parseName(name)
    svg2png(bytestring=svgName.safe_substitute(template), write_to=f"pictures\\temp\\names\\name{index + 1}.png")
    return Image.open(f"pictures\\temp\\names\\name{index + 1}.png")


def makeNumber(svgNumber, number, color, index, template):
    template['color'] = color
    template['number'] = parseName(number)
    svg2png(bytestring=svgNumber.safe_substitute(template), write_to=f"pictures\\temp\\numbers\\number{index + 1}.png")
    return Image.open(f"pictures\\temp\\numbers\\number{index + 1}.png")


def clearDirectory(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


if __name__ == "__main__":
    df = load_data()
    df['Количество баллов'] = df['Количество баллов'].astype('int32')

    makeCenters(df)
    makeBackground()
    clearDirectory('pictures\\temp\\centers')
    clearDirectory('pictures\\temp\\names')
    clearDirectory('pictures\\temp\\numbers')
