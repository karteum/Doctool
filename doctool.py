#!/usr/bin/python
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
from PIL import Image
#import base64
#def docx_b64decode(b64bstring):
#    return base64.b64decode(b64bstring.replace(b'#xA;',b'\n') + b'==')

def zip_update(zipname, newfiledata, destfile=None, deleted=[]):
    """Update subfiles with new data within zipfile (newfiledata is a dict with {filename: data, ...})"""
    # Unfortunately this can only be done by re-creating the whole zipfile
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)
    with ZipFile(zipname, 'r') as zin, ZipFile(tmpname, 'w', compression=ZIP_DEFLATED, compresslevel=5) as zout:
        zout.comment = zin.comment # preserve the comment
        for item in zin.infolist():
            if item.filename in deleted: continue
            if item.filename in newfiledata.keys():
                zout.writestr(item, newfiledata[item.filename])
                del newfiledata[item.filename]
            else:
                zout.writestr(item, zin.read(item.filename))
        for remaining in newfiledata.keys():
            zout.writestr(remaining, newfiledata[remaining])
    os.replace(tmpname, zipname if destfile==None else destfile)

def docx_remove_protection(docxfile):
    """Remove protection (e.g. restrictions on formatting, etc) from docx file"""
    xmldata = ZipFile(docxfile).open("word/settings.xml").read().decode()
    xmldata = re.sub("<w:documentProtection .*/>", "", xmldata)
    zip_update(docxfile, {"word/settings.xml": xmldata})

def docx_change_authors(docxfile, authorstable, outfile=None):
    """Change authors of track changes. authorstable is a dict with entries {'oldauthor1': 'newauthor1', 'oldauthor2': 'newauthor2', ...}"""
    # FIXME: check what happens if merging author A->C and author B->C with conflicting/overlapping track changes
    with ZipFile(docxfile, 'r') as zin:
        xmldata_document = zin.open("word/document.xml").read().decode() if "word/document.xml" in zin.namelist() else ""
        xmldata_comment = zin.open("word/comments.xml").read().decode() if "word/comments.xml" in zin.namelist() else ""
        xmldata_people = zin.open("word/people.xml").read().decode() if "word/people.xml" in zin.namelist() else ""
    for old,new in authorstable.items():
        print(f'Replacing {old} -> {new}')
        xmldata_document=xmldata_document.replace(f'w:author="{old}"', f'w:author="{new}"')
        xmldata_comment=xmldata_comment.replace(f'w:author="{old}"', f'w:author="{new}"')
        xmldata_people=xmldata_people.replace(f'w15:author="{old}"', f'w15:author="{new}"')
    zip_update(docxfile, {"word/document.xml": xmldata_document,
                          "word/comments.xml": xmldata_comment,
                          "word/people.xml": xmldata_people},
               destfile=outfile)

def docx_list_authors(docxfile, splitdates=False):
    """List authors in track changes. If splitdates==True, the date is appended to author names"""
    xmldata = ZipFile(docxfile).open("word/document.xml").read().decode()
    if splitdates==False:
        p = re.compile('w:author="(.*?)"')
        author_list = list(set(p.findall(xmldata)))
    else:
        p = re.compile('w:author="(.*?)" w:date="(.*?)T')
        res = p.findall(xmldata)
        author_list = list(set([b.replace('-', '') + '_' + a for a,b in res]))
    return author_list

def png2jpg(fin,fout):
    # Converts png to jpg using PIL and performing RGBA to RGB conversion with transparent color -> white
    im = Image.open(fin)
    try:
        if im.mode=="RGBA":
            im2 = Image.new('RGB', im.size, (255, 255, 255))
            im2.paste(im, mask=im.split()[3]) # 3 is the alpha channel
            im2.save(fout)
        elif im.mode == "P":
            im.convert('RGBA').convert('RGB').save(fout)
        else: im.save(fout)
        return "ok"
    except: # malformed .png files. FIXME: use im.verify()
        return None

def docx_slimfast(docxfile, outfile=None): # FIXME: avoid use of os.system(), which probably could be exploited with a malicious .docx
    pwd = os.getcwd() ; emf2svg_conv = f"LD_LIBRARY_PATH={pwd} {pwd}/emf2svg-conv" # https://github.com/kakwa/libemf2svg
    deleted=[]
    newfiledata = {}
    maxsize = 600000
    with ZipFile(docxfile, 'r') as zin, tempfile.TemporaryDirectory() as extract_dir:
        xmldata_rels = zin.open("word/_rels/document.xml.rels").read().decode()
        for afile in zin.infolist():
            path, ext = os.path.splitext(afile.filename)
            bname = os.path.basename(path)
            fin = zin.extract(afile.filename, path=extract_dir)
            if ext.lower() == ".emf":
                svgfile = f"{extract_dir}/{path}.svg"
                os.system(f"{emf2svg_conv} --input {fin} --output {svgfile}") # FIXME: replace this quick hack with something not using os.system()
                if os.stat(svgfile).st_size > maxsize and afile.file_size > maxsize:
                    # convert to jpg. Lossy but often smaller. Conversion starts from svg since emf2svg_conv works well and other emf rasterizers are often low quality
                    os.system(f"convert {svgfile} {extract_dir}/{path}.jpg") # FIXME: replace this quick hack with something not using os.system()
                    deleted.append(afile.filename)
                    xmldata_rels = xmldata_rels.replace(bname+ext, bname+".jpg")
                    newfiledata[path+".jpg"] = open(f"{extract_dir}/{path}.jpg", "rb").read()
                elif os.stat(svgfile).st_size < afile.file_size: # If the emf is smaller than the svg, of course we keep the emf !
                    deleted.append(afile.filename)
                    xmldata_rels = xmldata_rels.replace(bname+ext, bname+".svg")
                    newfiledata[path+".svg"] = open(f"{extract_dir}/{path}.svg", "rb").read()
            elif ext.lower() == ".png":
                #os.system(f"zopflipng -y --lossy_8bit {fin} {fin}")
                #os.system(f"pngcrush -ow {fin}")
                ret = png2jpg(fin, f"{extract_dir}/{path}.jpg")
                deleted.append(afile.filename)
                if ret != None:
                    xmldata_rels = xmldata_rels.replace(bname+ext, bname+".jpg")
                    newfiledata[path+".jpg"] = open(f"{extract_dir}/{path}.jpg", "rb").read()
        with open(extract_dir + "/word/_rels/document.xml.rels", "w") as fh:
            fh.write(xmldata_rels)
        newfiledata["word/_rels/document.xml.rels"] = xmldata_rels
    zip_update(docxfile, newfiledata, destfile=outfile, deleted=deleted)
    return deleted, newfiledata

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("docfile", help="Docx path")
    subparsers = parser.add_subparsers(dest="subcommand", required=True)
    parser_removeprot = subparsers.add_parser('remove_protection', help="Remove protection")
    parser_listauth = subparsers.add_parser('list_authors', help="List authors")
    parser_chauth = subparsers.add_parser('change_authors', help='Change authors: "old1" "new1" "old2" "new2"...')
    parser_chauth.add_argument('-o', '--outputfile', help='Output file name', default=None)
    parser_chauth.add_argument('authors', nargs='*')
    parser_slimfast = subparsers.add_parser('slimfast', help="make the docx more lightweight (lossy compression on pictures)")
    parser_slimfast.add_argument('-o', '--outputfile', help='Output file name', default=None)

    args = parser.parse_args()
    if args.subcommand=='remove_protection':
        docx_remove_protection(args.docfile)
    elif args.subcommand=='list_authors':
        res=docx_list_authors(args.docfile)
        print('\n'.join(res))
    elif args.subcommand=='change_authors':
        a=iter(args.authors)
        authlist={k:v for k,v in zip(a,a)}
        docx_change_authors(args.docfile, authlist,outfile=args.outputfile)
    elif args.subcommand=='slimfast':
        docx_slimfast(args.docfile, outfile=args.outputfile)
