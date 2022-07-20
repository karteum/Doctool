#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 22 14:19:13 2021
License: GPLv3 https://www.gnu.org/licenses/gpl-3.0.en.html

@author: Adrien DEMAREZ
"""

from zipfile import ZipFile,ZIP_DEFLATED,BadZipFile
import os,re
import tempfile
import argparse
from PIL import Image
import sys
from glob import glob
import shutil
from time import time
from flask import Flask, request, send_file, after_this_request
app=Flask(__name__)

#import base64
#def docx_b64decode(b64bstring):
#    return base64.b64decode(b64bstring.replace(b'#xA;',b'\n') + b'==')

def zip_update(zipname, newfiledata, destfile=None, deleted=[]):
    print(zipname)
    """Update subfiles with new data within zipfile (newfiledata is a dict with {filename: data, ...})"""
    # Unfortunately this can only be done by re-creating the whole zipfile
    tmpfd, tmpname = tempfile.mkstemp() # dir=os.path.dirname(zipname)
    os.close(tmpfd)
    with ZipFile(zipname, 'r') as zin, ZipFile(tmpname, 'w', compression=ZIP_DEFLATED, compresslevel=5) as zout:
        zout.comment = zin.comment # preserve the comment
        for item in zin.infolist():
            if item.filename in deleted: continue
            if item.filename in newfiledata.keys():
                zout.writestr(item, newfiledata[item.filename])
                del newfiledata[item.filename]
            else:
                try: zout.writestr(item, zin.read(item.filename))
                except BadZipFile:
                    print(f"\n__ Error on file {item.filename}")
                    zout.writestr(item, "_error_") # FIXME: force extract the partial/corrupted file
                    #zout.writestr(item, open("image104.tiff", "rb").read()) # FIXME: force extract the partial/corrupted file
        for remaining in newfiledata.keys():
            zout.writestr(remaining, newfiledata[remaining])
    #os.replace(tmpname, zipname if destfile==None else destfile)
    if destfile:
        shutil.move(tmpname, destfile)
        return destfile
    return tmpname

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
        if im.mode == "P": im = im.convert('RGBA') # Convert to RGBA first, then will be handled properly
        if im.mode=="RGBA":
            im2 = Image.new('RGB', im.size, (255, 255, 255))
            im2.paste(im, mask=im.split()[3]) # 3 is the alpha channel
            im2.save(fout)
        else: im.save(fout)
        return True
    except: # malformed .png files. FIXME: use im.verify()
        print(" error png2jpg")
        return False

def docx_slimfast(docxfile, outfile=None, do_png=True, do_emf=True, do_charts=False): # FIXME: avoid use of os.system(), which probably could be exploited with a malicious .docx
    """Reduces the size of the .docx file by converting embedded images : PNG over 30kB are converted to JPG, and EMF are converted first to SVG (using libemf2svg, which seems to produce good quality results). Resulting SVG may be already significantly more lightweight than the original EMF in some case, and if it is still above 600kB the script will rasterize the SVG to JPG. Of course all of this is a lossy compression => use it at your own risk and check the result !"""
    # TODO: handle charts
    pwd = os.getcwd() ; emf2svg_conv = f"LD_LIBRARY_PATH={pwd} {pwd}/emf2svg-conv" # https://github.com/kakwa/libemf2svg
    deleted=[]
    newfiledata = {}
    with ZipFile(docxfile, 'r') as zin, tempfile.TemporaryDirectory() as extract_dir:
        xmldata_rels = zin.open("word/_rels/document.xml.rels").read().decode()
        numfiles = len(zin.namelist()) ; k=0
        for afile in zin.infolist():
            path, ext = os.path.splitext(afile.filename)
            bname = os.path.basename(path)
            if ext.lower() == ".emf" and do_emf==True:
                fin = zin.extract(afile.filename, path=extract_dir)
                # Conversion starts from svg since emf2svg_conv works well and other emf rasterizers are often low quality
                svgfile = f"{extract_dir}/{path}.svg" ; os.system(f"{emf2svg_conv} --input {fin} --output {svgfile}") # FIXME: replace this quick hack with something not using os.system()
                pngfile = f"{extract_dir}/{path}.png" ; os.system(f"inkscape {svgfile} --export-type=png 2>/dev/null")
                #os.system(f"convert {svgfile} {pngfile}") # FIXME: replace this quick hack with something not using os.system()
                jpgfile = f"{extract_dir}/{path}.jpg" ; png2jpg(pngfile, jpgfile)
                if 2 * min(os.stat(jpgfile).st_size, os.stat(pngfile).st_size) < afile.file_size: # Only do it if it makes sense from a size perspective
                    deleted.append(afile.filename)
                    if os.stat(jpgfile).st_size * 1.5 < os.stat(pngfile).st_size:
                        xmldata_rels = xmldata_rels.replace(bname+ext, bname+".jpg")
                        newfiledata[path+".jpg"] = open(f"{jpgfile}", "rb").read()
                    else:
                        xmldata_rels = xmldata_rels.replace(bname+ext, bname+".png")
                        newfiledata[path+".png"] = open(f"{pngfile}", "rb").read()
            elif ext.lower() == ".png" and do_png==True: # and afile.file_size > 30000
                fin = zin.extract(afile.filename, path=extract_dir)
                #os.system(f"zopflipng -y --lossy_8bit {fin} {fin}")
                #os.system(f"pngcrush -ow {fin}")
                jpgfile = f"{extract_dir}/{path}.jpg"
                if png2jpg(fin, jpgfile) and os.stat(jpgfile).st_size *1.5 < afile.file_size: # if the jpg is more than 66% of the original png size, then keep the png
                    xmldata_rels = xmldata_rels.replace(bname+ext, bname+".jpg")
                    newfiledata[path+".jpg"] = open(f"{extract_dir}/{path}.jpg", "rb").read()
                    deleted.append(afile.filename)
            elif path.startswith('word/charts') and do_charts==True:
                fin = zin.extract(afile.filename, path=extract_dir)
                #relfile = f'{os.path.dirname(path)}/_rels/{bname}.xml.rels'
                #rel = zin.extract(relfile, path=extract_dir)
                chart_img = render_chart(fin, afile.filename)
            #elif path.startswith('word/charts/_rels'):
            #    continue
            k+=1
            sys.stderr.write(f'\r\033[2K{100*k//numfiles} % : {afile.filename}') ; sys.stderr.flush()
        xmldata_rels = xmldata_rels.replace('/>', '/>\n')
        newfiledata["word/_rels/document.xml.rels"] = xmldata_rels
        content_types = zin.open("[Content_Types].xml").read().decode().replace('/>', '/>\n')
        insert_index = content_types.index("<Default Extension=")
        if not '<Default Extension="jpg"' in content_types:
            content_types = content_types[:insert_index] + '<Default Extension="jpg" ContentType="image/jpeg"/>\n' + content_types[insert_index:]
        if not '<Default Extension="png"' in content_types:
            content_types = content_types[:insert_index] + '<Default Extension="png" ContentType="image/png"/>\n' + content_types[insert_index:]
        newfiledata["[Content_Types].xml"] = content_types
    return zip_update(docxfile, newfiledata, destfile=outfile, deleted=deleted)

def render_chart(chartfile, chartname):
    docxparts = {chartname: open(chartfile).read()} # , relname: open(relfile).read()
    docstring = """<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5731510" cy="3978910"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Objet1"/>
<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
</a:graphicData></a:graphic></wp:inline></w:drawing>"""
    contents_types = f'<Override PartName="/{chartname}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    relstring = f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{chartname[5:]}"/>'
    with ZipFile("blank.docx", 'r') as zin:
        docxparts["word/document.xml"] = zin.open("word/document.xml").read().decode().replace("Test", docstring)
        docxparts["word/_rels/document.xml.rels"] = zin.open("word/_rels/document.xml.rels").read().decode().replace("</Relationships>", relstring+"\n</Relationships>")
        docxparts["[Content_Types].xml"] = zin.open("[Content_Types].xml").read().decode().replace("</Types>", contents_types + "\n</Types>")

    #with tempfile.NamedTemporaryFile() as tmpdocx, tempfile.TemporaryDirectory() as extract_dir, tempfile.NamedTemporaryFile() as outimg:
    tmpdocx = "tmp.docx" ; extract_dir = "foo" ; outimg = "outimg"
    zip_update("blank.docx", docxparts, destfile=tmpdocx)
    os.system(f'libreoffice --convert-to "html:HTML" {tmpdocx} --outdir {extract_dir}') # FIXME: replace this quick hack with something not using os.system()
    gifname = glob(f"{extract_dir}/*.gif")
    im = Image.open(gifname).save(f"{outimg}.png")
    return open(f"{outimg}.png").read()


############## Web UI

@app.route('/', methods = ['GET', 'POST'])
def ui_root():    
    if request.method == 'GET':
        return """<html><head><title>Docx diet</title>
<style>body { font-family:sans-serif; background-color: #DFDBE5; }</style></head>
<body>

<h1>Docx diet</h1>
<img src="https://secretnews.fr/wp-content/uploads/2018/03/michel-ange-david-obese-gros.jpg" align="right" width="30%" style="transform: scaleX(-1);">
<i>Work in progress, non-multithreaded test server i.e. one single connection at a time, slow implementation (especially when processing EMF): wait up to one minute after clicking "send"...</i>
<form action="https://ksufi.karteum.ovh/docxdiet" method="POST" enctype="multipart/form-data" target="_blank"><br>
<label for="docx_file">Original docx file : </label><input type="file" name="docx_file" id="docx" /><br>
<input type="checkbox" id="do_png" name="do_png" checked><label for="do_png">PNG->JPG (slightly lossy)</label><br>
<input type="checkbox" id="do_emf" name="do_emf"><label for="do_emf">EMF->SVG->PNG (possibly lossy if bugs in rasterizer. Need double-check on result)</label><br>
<input type="submit" name="action" value="Send" />
</form>

</body></html>"""

    if request.files['docx_file']:
        docx_file = request.files['docx_file']
        do_png = True if "do_png" in request.form and request.form["do_png"]=="on" else False
        do_emf = True if "do_emf" in request.form and request.form["do_emf"]=="on" else False
        docx_newfile = docx_slimfast(docx_file, do_png=do_png, do_emf=do_emf)

        @after_this_request
        def remove_file(response):
            os.remove(docx_newfile)
            return response

        docx_mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        return send_file(docx_newfile, docx_mime, as_attachment=True, attachment_filename=f"docx_slimfast_{int(time())}.docx")

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
    parser_web = subparsers.add_parser('web', help="Web interface")

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
    elif args.subcommand=='web':
        #os.environ['FLASK_ENV']="development"
        #app.testing = True
        #app.debug = True
        print('Launch on port 5000')
        app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)
        sys.exit(0)
