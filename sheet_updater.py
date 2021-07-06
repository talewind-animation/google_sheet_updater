import os
import json
import argparse
import gspread
import gspread_formatting as gsf
from oauth2client.service_account import ServiceAccountCredentials
from termcolor import colored

root = os.path.dirname(os.path.abspath(__file__))
client_secret_file = os.getenv("UPDATER_CLIENT_SECRET", os.path.join(root, 'client_secret.json'))
updater_config_file = os.getenv("UPDATER_CONFIG", "")

if not os.path.exists(client_secret_file):
    raise ValueError("Please provide client secret file. Or set UPDATER_CLIENT_SECRET environment variable pointing to it.")

if not os.path.exists(updater_config_file):
    print(colored('Configs not found. Using defaults.', 'yellow'))
    configs = {
        "s": "Season 3 TZ",
        "m": "wip",
        "t": "",
        "cw": "Compose",
        "scn": "Scenes",
    }
else:
    with open(updater_config_file, "r") as f:
        configs = json.load(f)

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(client_secret_file, scope)
client = gspread.authorize(creds)

TZ_ROOT = os.getenv('TZ_ROOT')
TZ_COMP = os.path.join(TZ_ROOT, 'Compositing')
SHEET_NAMES = [
    'Season 3 TZ',
    'Season 4 TZ',
    ]

def gsf_color_hex(hex):
    hex = hex.replace('#', '')
    rgb = tuple(float(int(hex[i:i+2], 16))/256 for i in (0, 2, 4))
    return gsf.color(rgb[0], rgb[1], rgb[2])

modes = {
    'wip': {'bg_color': gsf_color_hex('#fff2cc'), 'comment': 'wip'},
    'rendered': {'bg_color': gsf_color_hex('#c6efce'), 'comment': 'published'},
    'bad': {'bg_color': gsf_color_hex('#f4cccc'), 'comment': 'bad'},
    'nuke': {'bg_color': gsf_color_hex('#eba850'), 'comment': 'nuke'},
}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', type=str, default=configs["s"], help='Season name. Google sheets document name(Season 3 by default)')
    parser.add_argument('-ep', type=str, help='Episode name. Google sheets page name')
    parser.add_argument('-sc', type=str, help='Scene name.')
    parser.add_argument('-m', type=str, default=configs["m"], help='Edit mode (wip, rendered, nuke, bad).')
    parser.add_argument('-t', type=str, default=configs["t"], help='Optional text to write in cell')
    parser.add_argument('-cw', type=str, default=configs["cw"], help='Column name to be edited(Compose by default)')
    parser.add_argument('-scn', type=str, default=configs["scn"], help='Scenes column name to be edited(Scenes by default)')

    args = parser.parse_args()

    if not args.ep or not args.sc:
        print(colored('Not enough arguments!', 'red'))
        return

    update_cell(args)

def get_digits(text):
    return str(int(''.join(filter(str.isdigit, text))))

def get_sheet_name(sname):
    for sh in SHEET_NAMES:
        if str(sname) in sh:
            return sh

def get_seasons():
    return [os.path.join(TZ_COMP, f) for f in os.listdir(TZ_COMP) if f.startswith('season_')]

def get_epname(name):
    seasons = get_seasons()
    for season in seasons:
        for ep in os.listdir(season):
            if name in ep:
                return ep
    return None

def get_sheet(sname, epname):
    if not epname or not sname:
        print(colored('Wrong Episode or Season name.', 'red'))
        return
    try:
        return client.open(get_sheet_name(sname)).worksheet(epname)
    except gspread.exceptions.WorksheetNotFound:
        print(colored('Sheet: {} not found in google sheet: {}'.format(epname, sname), 'red'))
        return None

def format_cell(sheet, address, mode='wip', text='default'):
    fmt = gsf.cellFormat(
        backgroundColor=modes[mode]['bg_color'],
        )
    if text == 'default':
        text = mode
    if text.isdigit():
        text = int(text)
    if text != '':
        sheet.update(address, text)

    gsf.format_cell_range(sheet, '{0}:{0}'.format(address), fmt)


def find_cell(sheet, sc_name, comp_row_name, scene_row_name):
    rows = sheet.row_values(1)
    if not comp_row_name in rows or not scene_row_name in rows:
        print(colored("Provide correct column names for 'scene' and 'compose'.", 'red'))
        return

    comp_clmn_index = rows.index(comp_row_name)+1
    scene_clmn_index = rows.index(scene_row_name)+1

    scenes = [s.split('_')[0].split('-')[0] for s in sheet.col_values(scene_clmn_index)]
    if not sc_name in scenes:
        print(colored("Can't find the scene {}!".format(sc_name), 'red'))
        return

    scene_row_index = scenes.index(sc_name)+1

    return sheet.cell(scene_row_index, comp_clmn_index)

def update_cell(args):
    sheet = get_sheet(args.s, get_epname(args.ep))
    if not sheet:
        return
    cell = find_cell(sheet, get_digits(args.sc), args.cw, args.scn)
    if not cell:
        return
    format_cell(sheet, cell.address, args.m, args.t)
    print(colored("Updated scene: {} with status '{}' in google sheets".format(
        args.sc,
        args.m,
        ), 'blue'))

if __name__ == '__main__':
    main()
