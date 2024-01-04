#!/usr/bin/env python

from simple_term_menu import TerminalMenu
from pyzotero import zotero
from pathlib import Path
import platform
import re
import os

dict_loffice_path = {
    'Darwin': '/Applications/LibreOffice.app/Contents/MacOS/soffice'
}

dict_filedialog_cmd = {
    'Darwin': "osascript -e 'tell application (path to frontmost application as text)\nset myFile to choose file\nPOSIX path of myFile\nend'"
}

def get_loffice_convert(loffice_custom_path = None):
    loffice_path = dict_loffice_path.get(platform.system())
    if loffice_path is None:
        raise Exception('Platform is not supported.')
    if not os.path.isfile(loffice_path) and loffice_custom_path is None:
        raise Exception('No default LibreOffice install, and alternative path not provided.')
    if not os.path.isfile(loffice_path) and loffice_custom_path is not None:
        loffice_path = loffice_custom_path
        if not os.path.isfile(loffice_path):
            raise Exception('Provided custom LibreOffice path is not valid.')
    def loffice_convert(path: str, outfmt: str):
        os.system(f'{loffice_path} --convert-to {outfmt} {path}')
    return loffice_convert

def get_filedialog_func():
    filedialog_cmd = dict_filedialog_cmd.get(platform.system())
    if filedialog_cmd is None:
        raise Exception('Platform is not supported.')
    filedialog_func = lambda: os.popen(filedialog_cmd).read().strip()
    return filedialog_func

def obtain_user_credentials():
    user_id = os.getenv('ZOTERO_UID')
    api_key = os.getenv('ZOTERO_API_KEY')
    if user_id is None:
        print('Please enter Zotero user ID key:')
        user_id = input()
    if api_key is None:
        print('Please enter Zotero API key:')
        api_key = input()
    return user_id, api_key

def zotero_user_login(user_id, api_key):
    return zotero.Zotero(user_id, "user", api_key)

def get_collection_names(collections):
    return [
        collection['data']['name']
        for collection in collections
    ]

def query_user_collection(client):
    collections = client.collections()
    collection_names = get_collection_names(collections)
    menu = TerminalMenu(
        collection_names,
        title="Zotero Collection"
    )
    choice_idx = menu.show()
    chosen_collection = collections[choice_idx]
    print('Collection:', chosen_collection['data']['name'])
    return chosen_collection

def get_collection_items(client, collection):
    collection_items = client.collection_items(collection['key'])
    print('Number of items:', len(collection_items))
    return collection_items

def update_doc_html(filename_html, dict_items, user_id):
    paper_html = None
    with open(filename_html) as handle:
        paper_html = handle.read()
    for key, item_code in dict_items.items():
        paper_html = paper_html.replace(
            '@'+key+'}', f'zu:{user_id}:{item_code}'r'}'
        )
    filename_out = Path(filename_html).stem + '_FMT.html'
    with open(filename_out, 'w') as handle:
        handle.write(paper_html)
    return filename_out

def main():
    user_id, api_key = obtain_user_credentials()
    client = zotero_user_login(user_id, api_key)
    collection = query_user_collection(client)
    collection_items = get_collection_items(client, collection)
    dict_items = {
        re.search(r'RN[0-9]+', item['data']['extra']).group(0): item['key']
        for item in collection_items
    }
    loffice_convert = get_loffice_convert()
    filedialog = get_filedialog_func()
    filename = filedialog()
    filename_ext = Path(filename).suffix
    os.chdir(Path(filename).parent)
    if filename_ext not in ['.docx', '.doc']:
        raise Exception('Unsupported input format. Must be .doc or .docx')
    filename_odt = filename.replace(filename_ext, '.odt')
    filename_html = filename.replace(filename_ext, '.html')
    loffice_convert(filename, 'odt')
    loffice_convert(filename_odt, 'html')
    if not os.path.isfile(filename_html):
        raise Exception('Something went wrong during conversion, no HTML file to edit.')
    filename_out = update_doc_html(filename_html, dict_items, user_id)
    loffice_convert(filename_out, 'odt')
    filename_out_odt = filename_out.replace('.html', '.odt')
    if not os.path.isfile(filename_out_odt):
        raise Exception('Something went wrong during conversion, ODT output file does not exist.')
    print()
    print('------')
    print('Please run the ODF scanner Zotero plugin on the followinf file:', filename_out_odt)
    print('Open the output file based on the name you defined in LibreOffice, click refresh citations, and save.')
    print('At that point, all citations should be converted to Zotero.')
    print('------')

if __name__ == '__main__':
    main()
