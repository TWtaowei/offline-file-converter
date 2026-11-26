"""Offline conversion routes. Uses only the libraries shipped next to this file."""
# 本项目自有代码及其修改版仅供非商业使用，禁止转售及企业内部业务使用。见根目录 LICENSE。
import csv
from contextlib import closing
import io
import itertools
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import uuid

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "vendor"))
ROOT = HERE.parent
if sys.platform == "darwin":
    OFFICE_CLI = Path(os.environ.get("LOCAL_CONVERTER_OFFICE", "/Applications/LibreOffice.app")) / "Contents/MacOS/soffice"
else:
    OFFICE_CLI = ROOT / "LibreOffice" / "program" / "soffice.com"
APP = ROOT / "本地文件转换.exe"
from PIL import Image, ImageOps

WRITER = {"doc", "docx", "odt", "rtf", "txt"}
CALC = {"xls", "xlsx", "ods", "csv", "tsv"}
SLIDES = {"ppt", "pptx", "odp"}
IMAGES = {"png", "jpg", "jpeg", "bmp", "webp"}
CORE = {**dict.fromkeys(WRITER, "docx"), **dict.fromkeys(CALC, "xlsx"),
        **dict.fromkeys(SLIDES, "pptx")}
OUTPUT = {"docx-layout": "docx", "pptx-edit": "pptx"}
FILTERS = {"doc": "MS Word 97", "docx": "Office Open XML Text", "odt": "writer8",
           "rtf": "Rich Text Format", "xls": "MS Excel 97", "xlsx": "Calc MS Excel 2007 XML",
           "ods": "calc8", "ppt": "MS PowerPoint 97", "pptx": "Impress MS PowerPoint 2007 XML",
           "odp": "impress8"}

def ext(path):
    return path.suffix[1:].lower()

def work_folder(parent):
    # Inherit the current job's Windows permissions, including its AppContainer SID.
    folder = parent / uuid.uuid4().hex[:12]
    folder.mkdir()
    return folder

def load_text(path):
    raw = path.read_bytes()
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode("utf-16")
    for codec in ("utf-8-sig", "gb18030", "cp1252"):
        try:
            return raw.decode(codec)
        except UnicodeDecodeError:
            pass
    raise ValueError("无法识别文本编码，请将文件保存为 UTF-8。")

def profile_at(tmp):
    from xml.etree import ElementTree as ET
    profile = tmp / "profile"
    user = profile / "user"
    user.mkdir(parents=True, exist_ok=True)
    ns = "{http://openoffice.org/2001/registry}"
    root = ET.Element(ns+"items")
    for path, key, val in [
        ("/org.openoffice.Office.Common/Security/Scripting","MacroSecurityLevel","3"),
        ("/org.openoffice.Office.Common/Security/Scripting","DisableMacrosExecution","true"),
        ("/org.openoffice.Office.Jobs/Jobs/['UpdateCheck']/Arguments","AutoCheckEnabled","false")]:
        item = ET.SubElement(root, "item", {ns+"path": path})
        ET.SubElement(ET.SubElement(item,"prop",{ns+"name":key, ns+"op":"fuse"}),"value").text = val
    ET.ElementTree(root).write(user/"registrymodifications.xcu",encoding="utf-8")
    return profile

def office_export(source, fmt, tmp):
    folder = work_folder(tmp)
    target = fmt if fmt == "pdf" else fmt+":"+FILTERS[fmt]
    cmd = [str(OFFICE_CLI), "-env:UserInstallation="+profile_at(tmp).as_uri(),
           "--headless", "--norestore","--nodefault","--convert-to", target,
           "--outdir", str(folder), str(source)]
    result = subprocess.run(cmd, capture_output=True, timeout=300,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    file = folder / (source.stem+"."+fmt)
    if result.returncode or not file.exists() or not file.stat().st_size:
        raise ValueError("本地引擎未能导出文件。请检查文件是否有打开密码或已经损坏。")
    return file

def picture(source):
    with Image.open(source) as im:
        im.load()
        return ImageOps.exif_transpose(im).convert("RGBA")

def white(im):
    rgba = im.convert("RGBA")
    rgb = Image.new("RGB", rgba.size,"white")
    rgb.paste(rgba, mask=rgba.getchannel("A"))
    return rgb

def png_bytes(im):
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()

def image_block(raw):
    with Image.open(io.BytesIO(raw)) as im:
        return ("image", png_bytes(ImageOps.exif_transpose(im).convert("RGBA")))

def sheet_name(name):
    return re.sub(r"[\x00-\x1f\\/*?:\[\]]", "_", str(name)).strip("'")[:31] or "Sheet"

def clean_text(value):
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", "" if value is None else str(value))

def read_word(source):
    from docx import Document
    from docx.table import Table
    doc = Document(source)
    parts = []
    for item in doc.iter_inner_content():
        if isinstance(item, Table):
            parts.append(("table", [[cell.text for cell in row.cells] for row in item.rows]))
            continue
        if item.text.strip():
            parts.append(("text", item.text))
        for blip in item._p.xpath(".//a:blip"):
            rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if rid:
                try: parts.append(image_block(doc.part.related_parts[rid].blob))
                except (OSError, ValueError): pass  # Text extraction also supports documents with vector illustrations.
    return [(source.stem, parts)]

def read_book(source):
    from openpyxl import load_workbook
    values = load_workbook(source, read_only=True, data_only=True)
    formulas = load_workbook(source, read_only=True, data_only=False)
    result = []
    try:
        for sheet in values:
            rows = []
            for row, original in zip(sheet.values, formulas[sheet.title].values):
                cells = [old if val is None and isinstance(old,str) and old.startswith("=") else val
                         for val,old in zip(row,original)]
                while cells and cells[-1] is None:
                    cells.pop()
                rows.append(cells)
            while rows and not rows[-1]:
                rows.pop()
            result.append((sheet.title, [("table", rows)] if rows else []))
    finally:
        values.close()
        formulas.close()
    return result

def read_slides(source):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    def blocks(shapes):
        for shape in sorted(shapes, key=lambda s:(s.top,s.left)):
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from blocks(shape.shapes)
            elif shape.has_table:
                yield "table", [[cell.text for cell in row.cells] for row in shape.table.rows]
            elif shape.has_text_frame and shape.text.strip():
                yield "text", shape.text
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try: yield image_block(shape.image.blob)
                except (OSError, ValueError): pass
    return [(f"幻灯片 {i}",list(blocks(slide.shapes))) for i,slide in enumerate(Presentation(source).slides,1)]

def render(page, long_edge=2200):
    bitmap = page.render(scale=min(2.2, long_edge/max(page.get_size())), rev_byteorder=True)
    try:
        return bitmap.to_pil().convert("RGB")
    finally:
        bitmap.close()

def read_pdf(source):
    import pdfplumber
    import pypdfium2 as pdfium
    parts = []
    with pdfplumber.open(source) as pdf, closing(pdfium.PdfDocument(source)) as rendered:
        for i, page in enumerate(pdf.pages):
            items = []
            view = rendered[i]
            try:
                if not page.chars:
                    items.append(("image",png_bytes(render(view))))
                else:
                    tables = page.find_tables()
                    boxes = [t.bbox for t in tables]
                    def outside(obj):
                        x = (obj["x0"]+obj["x1"])/2
                        y = (obj["top"]+obj["bottom"])/2
                        return not any(l <= x <= r and t <= y <= b for l,t,r,b in boxes)
                    ordered = [(t.bbox[1],t.bbox[0],"table",t.extract()) for t in tables]
                    for line in page.filter(outside).extract_text_lines(return_chars=False):
                        ordered.append((line["top"], line["x0"], "text", line["text"]))
                    if page.images:
                        snapshot = render(view)
                        for image in page.images:
                            box = tuple(int(v) for v in (max(0,image["x0"]-page.bbox[0])*snapshot.width/page.width,
                                max(0,image["top"]-page.bbox[1])*snapshot.height/page.height,
                                min(page.width,image["x1"]-page.bbox[0])*snapshot.width/page.width,
                                min(page.height,image["bottom"]-page.bbox[1])*snapshot.height/page.height))
                            if box[2]>box[0] and box[3]>box[1]:
                                ordered.append((image["top"],image["x0"],"image",png_bytes(snapshot.crop(box))))
                    items = [(kind,data) for _,__,kind,data in sorted(ordered,key=lambda v:(v[0],v[1]))]
            finally:
                view.close()
                page.close()
            parts.append((f"第 {i+1} 页",items))
    return parts

def sections(source,tmp):
    kind = ext(source)
    if kind == "pdf": return read_pdf(source)
    if kind in IMAGES: return [(source.stem,[("image",png_bytes(picture(source)))])]
    if kind == "txt": return [(source.stem,[("text",line) for line in load_text(source).splitlines()])]
    if kind in {"csv","tsv"}:
        lines = io.StringIO(load_text(source), newline="")
        return [(source.stem,[("table",list(csv.reader(lines,delimiter="," if kind=="csv" else "\t")))])]
    native = CORE[kind]
    return {"docx":read_word,"xlsx":read_book,"pptx":read_slides}[native](
        source if kind == native else office_export(source,native,tmp))

def block_rows(blocks):
    for kind,data in blocks:
        if kind == "table": yield from data
        elif kind == "text": yield from ([line] for line in data.splitlines() or [""])

def write_word(parts,path):
    from docx import Document
    from docx.shared import Inches
    doc = Document()
    for page,(name,blocks) in enumerate(parts):
        if page: doc.add_page_break()
        if len(parts)>1: doc.add_heading(clean_text(name),2)
        for kind,data in blocks:
            if kind=="text": doc.add_paragraph(clean_text(data))
            elif kind=="image":
                with Image.open(io.BytesIO(data)) as im:
                    ratio = min(6.3/im.width,8.0/im.height,1/96)
                    doc.add_picture(io.BytesIO(data),width=Inches(im.width*ratio))
            elif data:
                cols=max(1,max(map(len,data)))
                # Wide worksheets are split into consecutive column groups so Word tables stay readable.
                for start in range(0,cols,12):
                    table=doc.add_table(rows=0,cols=min(12,cols-start))
                    table.style="Table Grid"
                    for row in data:
                        cells=table.add_row().cells
                        for cell,val in zip(cells,row[start:start+12]): cell.text=clean_text(val)
                    doc.add_paragraph()
    doc.save(path)

def write_book(parts,path):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as SheetImage
    from openpyxl.utils import get_column_letter
    wb=Workbook()
    wb.remove(wb.active)
    for name,blocks in parts:
        ws=wb.create_sheet(sheet_name(name))
        pos=1
        for kind,data in blocks:
            if kind=="image":
                img=SheetImage(io.BytesIO(data))
                scale=min(1,850/img.width,1050/img.height)
                img.width,img.height=int(img.width*scale),int(img.height*scale)
                ws.add_image(img,"A"+str(pos))
                pos+=(img.height+18)//19+1
            else:
                for row in block_rows([(kind,data)]):
                    for j,val in enumerate(row,1):
                        if val is None: continue
                        if isinstance(val,str): val=clean_text(val)
                        cell=ws.cell(pos,j,val)
                        if isinstance(val,str): cell.data_type="s"
                        dim=ws.column_dimensions[get_column_letter(j)]
                        dim.width=min(65,max(dim.width,2+len(str(val))))
                    pos+=1
    if not wb.worksheets: wb.create_sheet()
    wb.save(path)

def make_pdf(source,tmp):
    kind=ext(source)
    if kind=="pdf": return source
    out=work_folder(tmp)
    if kind in IMAGES:
        file=out/"image.pdf"
        white(picture(source)).save(file,"PDF",resolution=144)
        return file
    if kind in {"txt","csv","tsv"}:
        native=CORE[kind]
        canonical=tmp/("text-input."+native)
        (write_word if native=="docx" else write_book)(sections(source,tmp),canonical)
        source=canonical
    if sys.platform == "darwin":
        return office_export(source, "pdf", tmp)
    proc=subprocess.run([str(APP),"--convert","pdf",str(source),str(out)],capture_output=True,creationflags=subprocess.CREATE_NO_WINDOW)
    files=list(out.glob("*.pdf"))
    if proc.returncode or len(files)!=1: raise ValueError("无法排版本地文档，请检查源文件。")
    return files[0]

def page_images(source,tmp):
    if ext(source) in IMAGES:
        yield picture(source)
        return
    import pypdfium2 as pdfium
    with closing(pdfium.PdfDocument(make_pdf(source,tmp))) as pdf:
        if not len(pdf): raise ValueError("源文件没有可导出的页面。")
        for index in range(len(pdf)):
            page=pdf[index]
            try: yield render(page)
            finally: page.close()

def save_pages(source,mode,path,tmp):
    from docx import Document
    from docx.shared import Inches
    from docx.enum.section import WD_SECTION
    from docx.shared import Pt
    from pptx import Presentation
    from pptx.util import Inches as SlideInches
    doc = Document() if mode=="docx-layout" else None
    slides = Presentation() if mode=="pptx" else None
    folder=path.parent/"multiple"
    first=None
    for i,im in enumerate(page_images(source,tmp),1):
        ratio=im.width/im.height
        if slides is not None:
            if i==1:
                slides.slide_width=SlideInches(max(1,min(55,10*ratio)))
                slides.slide_height=SlideInches(10)
            slide=slides.slides.add_slide(slides.slide_layouts[6])
            scale=min(slides.slide_width/im.width,slides.slide_height/im.height)
            w,h=round(im.width*scale),round(im.height*scale)
            slide.shapes.add_picture(io.BytesIO(png_bytes(im)),(slides.slide_width-w)//2,
                                     (slides.slide_height-h)//2,w,h)
        elif doc is not None:
            sec=doc.sections[0] if i==1 else doc.add_section(WD_SECTION.NEW_PAGE)
            # Fit the page image inside Word's page size limits, preserving its aspect ratio.
            h=min(11,20/ratio)
            sec.page_width,sec.page_height=Inches(max(.5,h*ratio)),Inches(h)
            sec.left_margin=sec.right_margin=sec.top_margin=sec.bottom_margin=Inches(.05)
            p=doc.add_paragraph()
            p.paragraph_format.space_before=p.paragraph_format.space_after=Pt(0)
            p.add_run().add_picture(io.BytesIO(png_bytes(im)),height=Inches(max(.1,min(h-.33,(h*ratio-.1)/ratio))))
        else:
            if i==2:
                folder.mkdir()
                path.rename(folder/("001."+mode))
            target=path if i==1 else folder/(f"{i:03d}."+mode)
            (white(im) if mode in ("jpg","bmp") else im).save(target,quality=92)
        first=i
    if first is None: raise ValueError("源文件没有页面。")
    if slides is not None: slides.save(path)
    if doc is not None: doc.save(path)

def editable_slides(parts,path):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    import textwrap
    ppt=Presentation()
    ppt.slide_width,ppt.slide_height=Inches(13.33),Inches(7.5)
    def new_slide(name):
        slide=ppt.slides.add_slide(ppt.slide_layouts[6])
        title=slide.shapes.add_textbox(Inches(.5),Inches(.2),Inches(12.3),Inches(.65))
        title.text_frame.paragraphs[0].text=clean_text(name)
        title.text_frame.paragraphs[0].font.size=Pt(26)
        return slide
    def text_page(name,lines):
        slide=new_slide(name)
        box=slide.shapes.add_textbox(Inches(.55),Inches(1.1),Inches(12.2),Inches(5.9))
        box.text_frame.word_wrap=True
        box.text_frame.text="\n".join(lines)
        for p in box.text_frame.paragraphs: p.font.size=Pt(19)
    for name,blocks in parts:
        lines=[]
        for kind,data in blocks:
            if kind=="text":
                for line in data.splitlines():
                    lines.extend(textwrap.wrap(line,38,replace_whitespace=False) or [""])
                while len(lines)>=11:
                    text_page(name,lines[:11]); lines=lines[11:]
                continue
            if lines: text_page(name,lines); lines=[]
            if kind=="table" and data and any(data):
                columns=max(map(len,data))
                for col in range(0,columns,6):
                    for start in range(0,len(data),9):
                        rows=data[start:start+9]
                        slide=new_slide(name)
                        table=slide.shapes.add_table(len(rows),min(6,columns-col),Inches(.5),Inches(1.1),
                            Inches(12.3),Inches(min(5.8,.58*len(rows)))).table
                        for j,row in enumerate(rows):
                            for k,value in enumerate(row[col:col+6]):
                                cell=table.cell(j,k); cell.text=clean_text(value)
                                for p in cell.text_frame.paragraphs: p.font.size=Pt(14)
            elif kind=="image":
                slide=new_slide(name)
                with Image.open(io.BytesIO(data)) as im:
                    scale=min(Inches(12.2)/im.width,Inches(5.9)/im.height)
                    w,h=int(im.width*scale),int(im.height*scale)
                    slide.shapes.add_picture(io.BytesIO(data),(ppt.slide_width-w)//2,Inches(1.1),w,h)
        if lines: text_page(name,lines)
        elif not blocks: new_slide(name)
    if not len(ppt.slides): new_slide("")
    ppt.save(path)

def write_delimited(parts,path,fmt):
    folder=path.parent/"multiple"
    if len(parts)>1: folder.mkdir()
    for i,(name,blocks) in enumerate(parts,1):
        safe=re.sub(r'[<>:"/\\|?*\x00-\x1f]',"_",name)[:65].rstrip(". ")
        dest=path if len(parts)==1 else folder/f"{i:03d}-{safe}.{fmt}"
        rows=list(block_rows(blocks))
        if not rows and any(k=="image" for k,_ in blocks):
            raise ValueError("此页只有图像，提取为文本或数据需要 OCR。可选 Excel 或 Word 保留图像。")
        with dest.open("w",encoding="utf-8-sig",newline="") as stream:
            csv.writer(stream,delimiter="," if fmt=="csv" else "\t").writerows(rows)

def convert(source,mode,path,tmp):
    kind=ext(source)
    fmt=OUTPUT.get(mode,mode)
    if mode in IMAGES or mode=="docx-layout":
        return save_pages(source,mode,path,tmp)
    if mode=="pdf":
        import shutil
        shutil.copyfile(make_pdf(source,tmp),path)
        return
    if fmt in FILTERS and CORE.get(kind)==CORE.get(fmt) and kind not in {"csv","tsv","txt"} and mode!= "pptx-edit":
        import shutil
        shutil.copyfile(office_export(source,fmt,tmp),path)
        return
    native=CORE.get(fmt)
    if native=="pptx" and mode!="pptx-edit":
        canonical=path if fmt=="pptx" else tmp/"pages.pptx"
        save_pages(source,"pptx",canonical,tmp)
    else:
        parts=sections(source,tmp)
        if mode in {"csv","tsv"}: return write_delimited(parts,path,mode)
        if mode=="txt":
            if not any(any(any(str(v or "").strip() for v in row) for row in block_rows(b)) for _,b in parts):
                if any(k=="image" for _,b in parts for k,_ in b): raise ValueError("源文件仅含图片，可选按页转换，识别文字需要 OCR。")
            path.write_text("\n\n".join("\n".join("\t".join(clean_text(v) for v in row) for row in block_rows(b)) for _,b in parts),encoding="utf-8-sig")
            return
        canonical=path if fmt==native else tmp/("editable."+native)
        {"docx":write_word,"xlsx":write_book,"pptx":editable_slides}[native](parts,canonical)
    if canonical!=path:
        import shutil
        shutil.copyfile(office_export(canonical,fmt,tmp),path)

if __name__=="__main__":
    source,mode,path,error=sys.argv[1:]
    try:
        # Keep nested Office profile paths short and inherit the account's temp ACL.
        scratch=work_folder(Path(tempfile.gettempdir()))
        try:
            convert(Path(source),mode,Path(path),scratch)
        finally:
            shutil.rmtree(scratch)
    except Exception as exc:
        Path(error).write_text(str(exc) or type(exc).__name__,encoding="utf-8")
        sys.exit(1)
