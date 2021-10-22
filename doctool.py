# -*- coding: utf-8 -*-
"""
Created on Fri Oct 22 14:19:13 2021
License: GPLv3 https://www.gnu.org/licenses/gpl-3.0.en.html

@author: Adrien DEMAREZ
"""

from zipfile import ZipFile,ZIP_DEFLATED
import os,re
import tempfile
import argparse

def updateZip(zipname, filename, data, destfile=None):
    """Update sub-file filename with new data within zipfile. Since updating is not supported, a new archive must be created"""
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)

    with ZipFile(zipname, 'r') as zin, ZipFile(tmpname, 'w') as zout:
        zout.comment = zin.comment # preserve the comment
        for item in zin.infolist():
            if item.filename != filename:
                zout.writestr(item, zin.read(item.filename))
            else:
                zout.writestr(filename, data)

    if destfile:
        os.rename(tmpname, destfile)
    else:
        os.remove(zipname)
        os.rename(tmpname, zipname)

def docx_remove_protection(docxfile):
    """Remove protection (e.g. restrictions on formatting, etc) from docx file"""
    xmldata = ZipFile(docxfile).open("word/settings.xml").read().decode()
    xmldata = re.sub("<w:documentProtection .*/>", "", xmldata)
    updateZip(docxfile, "word/settings.xml", xmldata)

def docx_change_authors(docxfile, authorstable, splitdates=False):
    """Change authors of track changes. authorstable is a dict with entries {'oldauthor1': 'newauthor1', 'oldauthor2': 'newauthor2', ...}"""
    xmldata = ZipFile(docxfile).open("word/document.xml").read().decode()
    for old,new in authorstable.items():
        print(f'Replacing {old} -> {new}')
        xmldata=xmldata.replace(f'w:author="{old}"', f'w:author="{new}"')
    updateZip(docxfile, "word/document.xml", xmldata)

def docx_list_authors(docxfile, splitdates=False):
    """List authors in track changes. If splitdates==True, the date is appended to author names"""
    xmldata = ZipFile(docxfile).open("word/document.xml").read().decode()
    # = re.compile("w:author=\"(.*?)\"")
    if splitdates==False:
        p = re.compile('w:author="(.*?)"')
        author_list = list(set(p.findall(xmldata)))
    else:
        p = re.compile('w:author="(.*?)" w:date="(.*?)T')
        res = p.findall(xmldata)
        author_list = list(set([b.replace('-', '') + '_' + a for a,b in res]))
    return author_list

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("docfile", help="Docx path")
    subparsers = parser.add_subparsers(dest="subcommand", required=True)
    parser_removeprot = subparsers.add_parser('remove_protection', help="Remove protection")
    parser_listauth = subparsers.add_parser('list_authors', help="List authors")
    parser_chauth = subparsers.add_parser('change_authors', help='Change authors: "old1" "new1" "old2" "new2"...')
    parser_chauth.add_argument('authors', nargs='*')

    args = parser.parse_args()
    if args.subcommand=='remove_protection':
        docx_remove_protection(args.docfile)
    elif args.subcommand=='list_authors':
        res=docx_list_authors(args.docfile)
        print('\n'.join(res))
    elif args.subcommand=='change_authors':
        a=iter(args.authors)
        authlist={k:v for k,v in zip(a,a)}
        docx_change_authors(args.docfile, authlist)
