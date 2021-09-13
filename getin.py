"""Python module to replace the password in VBA project"""
import argparse
import os
import re
import shutil
import sys
import traceback
import zipfile
from typing import List, Optional


class this:
    # Known password: 0
    known_pw = b'0F0DA38BE78F04AC04ACFB5405ACB7F3C2613696189B7A52D119BCA91EC8FDBE9E59AEA82B9A46'
    msaccess_ext = ['.mdb', '.accdb']
    inplace = False


def update_xl_vba_project(zip_name, new_zip_name, data):
    """Creates new zip archive and replacing the contents of xl/vbaProject.bit with 'data'

    Params
    ------
    zip_name: zip archive to copy from
    new_zip_name: zip archive to create
    data: data to replace
    """

    bin_name = r'xl/vbaProject.bin'

    with zipfile.ZipFile(zip_name, 'r') as zin:
        with zipfile.ZipFile(new_zip_name, 'w') as zout:
            zout.comment = zin.comment  # preserve the comment
            for item in zin.infolist():
                if item.filename.find(bin_name) == -1:
                    zout.writestr(item, zin.read(item.filename))

    with zipfile.ZipFile(new_zip_name, mode='a', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(bin_name, data)


def read_vba_project(file_name):
    """Reads the contents of xl/vbaProject.bin from the zip archive

    Params
    ------
    file_name: zip archive to read from
    """

    _, ext = os.path.splitext(file_name)
    if ext in this.msaccess_ext:
        with open(file_name, mode='r+b') as fd:
            return fd.read()
    else:
        with zipfile.ZipFile(file_name, mode='r') as zf:
            return zf.read('xl/vbaProject.bin')


def unlock_vba(file_name: str) -> bool:
    """Unlocks VBA project

    Params
    ------
    file_name: path to Excel file in new macro enabled format (xlsm, xlam)
    """
    try:
        vba_project = read_vba_project(file_name)
    except Exception:
        traceback.print_exc()
        return False

    try:
        vba_project = find_and_replace(this.known_pw, vba_project)
        unlock_file_name = get_unlock_filename(file_name)

        _, ext = os.path.splitext(file_name)
        if ext in this.msaccess_ext:
            shutil.copyfile(file_name, unlock_file_name)

            with open(unlock_file_name, mode='w+b') as fd:
                fd.write(vba_project)
        else:
            update_xl_vba_project(file_name, unlock_file_name, vba_project)

        if this.inplace:
            os.remove(file_name)
            os.rename(unlock_file_name, file_name)
    except Exception:
        traceback.print_exc()
        return False
    return True


def find_and_replace(pw: str, vba_project: bytes) -> bytes:
    start = vba_project.find(b'\x44\x50\x42\x3D\x22') + 5 # find DPB="
    end = start + vba_project[start:].find(b'\x22')       # find next "
    password = pw
    if end - start > len(pw):
        password = pw + b'0' * (end - start - len(pw))
    vba_project = vba_project.replace(vba_project[start : end], password)
    return vba_project

def get_unlock_filename(file_name: str) -> str:
    path = os.path.dirname(file_name)
    base, ext = os.path.splitext(os.path.basename(file_name))
    unlock = '_unlocked'
    fname = ''.join([base, unlock, ext])

    i = 0
    while os.path.isfile(os.path.join(path, fname)):
        suffix = ' ({})'.format(str(i))
        fname = ''.join([base, unlock, suffix, ext])
        i += 1

    return os.path.join(path, fname)


def get_filelist(args: argparse.Namespace) -> Optional[List[str]]:
    path = os.getcwd()
    file_list: List[str] = None
    if args.files:
        file_list = args.files
    elif args.extensions:
        # adding . ot each value, if missing
        extensions: List[str] = []
        for ext in args.extensions:
            if not str(ext).startswith('.'):
                ext = '.' + ext
            extensions.append(ext)
        file_list = [f for f in os.listdir(path)
                     if os.path.splitext(f)[1] in extensions]
    elif args.regex:
        pattern = re.compile(args.regex)
        file_list = [f for f in os.listdir(path) if re.match(pattern, f)]
    return file_list


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Python module to replace the password in VBA project")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-f', '--files',
                       help='file list <file1 file2 ...>', 
                       nargs='+')
    group.add_argument('-e', '--extensions',
                       help='extensions list <xlsm  .mdb accdb ...> (with or without .)', 
                       nargs='+')
    group.add_argument('-r', '--regex',
                       help='regex pattern')
    parser.add_argument('-i', '--inplace',
                        help='Optional: update(s) files in place',
                        action='store_true')
    args = parser.parse_args()
    return args


def main():
    args = parse_args()
    file_list = get_filelist(args)
    if not file_list:
        raise FileNotFoundError('Could not find file(s) that match the citeria.')
    
    if args.inplace:
        this.inplace = True
    failed: List[str] = []
    succeeded: List[str] = []

    for file_name in file_list:
        if unlock_vba(file_name):
            lst = succeeded
        else:
            lst = failed
        lst.append(file_name)

    def printf(msg, lst):
        print(f'--- {msg}:\n-\t', end='')
        print('\n-\t'.join(lst))

    if succeeded:
        printf('Unlocked', succeeded)
    if failed:
        printf('Failed', failed)


if __name__ == '__main__':
    main()