# -*- coding: utf-8 -*-
"""
Smart Rename — 磁盘文件智能重命名工具

功能:
  1. 扫描指定根目录下所有 PDF/DOCX/PPTX/XLSX/XLS 文档 和 图片
  2. 解析文档内容提取标题 / 读取图片 EXIF 拍摄时间
  3. 自动重命名混乱文件名 (纯数字/hash/设备前缀/截图类/发票类)
  4. 已命名良好的文件不动
  5. 同目录重名自动加 _1 _2 后缀, 不覆盖任何文件

安全保障:
  - 仅重命名, 不删除任何文件
  - 同目录重名自动加后缀, 零覆盖
  - 完整日志留存

用法:
  python smart_rename.py -r D:\\Work D:\\Downloads        # 预览, 不改
  python smart_rename.py -r D:\\Work --apply              # 执行重命名
"""
import os, sys, re, csv, datetime, argparse
from pathlib import Path

SKIP_DIRS = {'.git','node_modules','.wechat','Tencent Files','__pycache__'}
DOC_EXTS = {'.pdf','.docx','.pptx','.xlsx','.xls'}
IMG_EXTS = {'.jpg','.jpeg','.png','.cr2','.cr3','.nef','.arw','.dng','.heic','.heif','.bmp','.gif','.webp'}

def ensure_libs():
    missing = []
    for mod in ['exifread','PIL','PyPDF2','docx','pptx','openpyxl']:
        try: __import__(mod.replace('PIL','PIL').replace('docx','docx').replace('pptx','pptx').replace('openpyxl','openpyxl'))
        except: missing.append(mod)
    if not missing: return
    print(f"首次运行, 安装依赖: {missing}")
    import subprocess
    subprocess.run([sys.executable,'-m','pip','install','--quiet'] + missing, capture_output=True, timeout=180)
    print("依赖安装完成")

def clean(name, maxlen=80):
    if not name: return ''
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', ' ', str(name))
    name = re.sub(r'\s+', ' ', name).strip()
    name = re.sub(r'^[\._\-\s]+|[\._\-\s]+$', '', name)
    noise = [r'^Microsoft\s+(PowerPoint|Word|Excel|Office)\s*[-–—]?\s*',
             r'^PowerPoint Presentation\s*[-–—]?\s*',
             r'^Title Slide Layout.*?(?:–|—|-)\s*']
    for n in noise: name = re.sub(n, '', name, flags=re.IGNORECASE)
    name = re.sub(r'(电\s*子\s*发\s*票|发\s*票\s*号\s*码)', lambda m: re.sub(r'\s+','',m.group(0)), name)
    name = re.sub(r'\.(docx|pptx|xlsx?|xls?)\.(docx|pptx|xlsx?|xls?|pdf)$', r'.\2', name, flags=re.IGNORECASE)
    name = re.sub(r'\.(docx|pptx|xlsx?|xls?)\.(pdf)$', r'.pdf', name, flags=re.IGNORECASE)
    return name[:maxlen] if len(name) > maxlen else name

BAD = {'PowerPoint Presentation','Word Document','Excel Workbook','Title Slide Layout',
       'Sheet0','Sheet1','Sheet2','Sheet3','新建 Microsoft Excel 工作表','PowerPoint 演示文稿'}
def is_bad(t):
    if not t or len(t)<4: return True
    for b in BAD:
        if t.strip().lower() == b.lower(): return True
    return False

def is_invoice(t): return bool(t and re.search(r'电子发票|增值税发票|普通发票|发票号码', t))

def is_poor_doc(name):
    n = Path(name).stem
    if len(n)<4: return True
    if re.match(r'^\d{6,}$', n): return True
    if re.match(r'^[a-f0-9]{24,}$', n, re.IGNORECASE): return True
    if re.search(r'(id|doi|ssrn|arxiv|ieee)\s*[-_]?\d+', n, re.IGNORECASE): return True
    if len(n)<10 and not re.search(r'[\u4e00-\u9fff]', n): return True
    if re.match(r'^(final|draft|temp|tmp|new|old|copy|backup|report|doc|paper|test|data)\s*[-_vV]?\d*$', n, re.IGNORECASE): return True
    return False

def is_poor_img(name):
    n = Path(name).stem
    if re.match(r'^\d{6,}$', n): return True
    if re.match(r'^\d{14,}$', n): return True
    if re.match(r'^(IMG|VID|DSC|PANO|CR2|DCIM|FJ|NX|RAW|100_)?[-_]?\d{4,}$', n, re.IGNORECASE): return True
    if re.match(r'^[a-f0-9]{16,}$', n, re.IGNORECASE): return True
    if re.match(r'^(screenshot|capture|captura|shot|snap|img|image|photo|pic|picture|wechat|weixin|mmexport|wx|dingtalk|qq|sogou|bilibili)[-_]?\d{5,}$', n, re.IGNORECASE): return True
    if re.match(r'^(final|draft|temp|tmp|new|old|copy|backup)_?\d*$', n, re.IGNORECASE): return True
    if len(n)<8 and not re.search(r'[\u4e00-\u9fff]', n) and re.search(r'\d', n): return True
    if len(n)<6: return True
    return False

def extract_pdf(path):
    try:
        from PyPDF2 import PdfReader
        r = PdfReader(path)
        meta = r.metadata
        if meta and getattr(meta,'title',None):
            t = clean(str(meta.title).strip(), 120)
            if t and len(t)>5 and not is_bad(t): return t, 'pdf_meta'
        if len(r.pages)>0:
            txt = r.pages[0].extract_text() or ''
            lines = [l.strip() for l in txt.split('\n') if l.strip()]
            for l in lines[:20]:
                if 5<len(l)<100 and re.search(r'[\u4e00-\u9fff]', l):
                    c = clean(l)
                    if c and not is_bad(c): return c, 'pdf_pg1_cn'
            for l in lines[:20]:
                if 5<len(l)<100:
                    c = clean(l)
                    if c and not is_bad(c): return c, 'pdf_pg1'
    except: pass
    return None, None

def extract_docx(path):
    try:
        from docx import Document
        d = Document(path)
        try:
            t = d.core_properties.title
            if t and t.strip() and len(t.strip())>3:
                c = clean(t.strip())
                if c and not is_bad(c): return c, 'docx_title'
        except: pass
        for para in d.paragraphs[:20]:
            text = para.text.strip()
            if not text or len(text)<4: continue
            style = para.style.name if para.style else ''
            bold = para.runs and any(r.bold and r.text.strip() for r in para.runs)
            if style.startswith('Heading') or bold:
                c = clean(text)
                if c and not is_bad(c): return c, 'docx_heading'
        for para in d.paragraphs[:15]:
            text = para.text.strip()
            if 6<len(text)<120:
                c = clean(text[:80])
                if c and not is_bad(c): return c, 'docx_p1'
    except: pass
    return None, None

def extract_pptx(path):
    try:
        from pptx import Presentation
        p = Presentation(path)
        for slide in list(p.slides)[:3]:
            for shape in slide.shapes:
                if not hasattr(shape,'text'): continue
                txt = shape.text.strip()
                if 5<len(txt)<120:
                    c = clean(txt)
                    if c and not is_bad(c): return c, 'pptx'
    except: pass
    return None, None

def extract_xlsx(path):
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        for n in wb.sheetnames:
            c = clean(n.strip())
            if c and len(c)>2 and not is_bad(c): return c, 'xlsx_sheet'
        if len(wb.sheetnames)>0:
            ws = wb[wb.sheetnames[0]]
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True)):
                cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
                if len(cells) >= 2 and any(len(cc)>3 for cc in cells):
                    for c in cells:
                        if 4<len(c)<80 and not is_bad(clean(c)):
                            return clean(c), 'xlsx_content'
                if len(cells)==1 and 4<len(cells[0])<80:
                    c = clean(cells[0])
                    if c and not is_bad(c): return c, 'xlsx_content'
    except: pass
    return None, None

def get_img_dt(path):
    try:
        import exifread
        with open(path, 'rb') as f:
            tags = exifread.process_file(f, details=False, stop_tag='EXIF DateTimeOriginal')
            if 'EXIF DateTimeOriginal' in tags:
                dt_str = str(tags['EXIF DateTimeOriginal']).replace(':','-',2)
                return datetime.datetime.strptime(dt_str[:19], '%Y-%m-%d %H:%M:%S'), 'exif'
            if 'Image DateTime' in tags:
                dt_str = str(tags['Image DateTime']).replace(':','-',2)
                return datetime.datetime.strptime(dt_str[:19], '%Y-%m-%d %H:%M:%S'), 'exif_img'
    except: pass
    try:
        from PIL import Image
        img = Image.open(path)
        exif = img._getexif() if hasattr(img, '_getexif') else None
        if exif:
            for tag_id in (36867, 36868, 306):
                if tag_id in exif:
                    dt_str = str(exif[tag_id]).replace(':','-',2)
                    return datetime.datetime.strptime(dt_str[:19], '%Y-%m-%d %H:%M:%S'), 'pillow'
    except: pass
    return datetime.datetime.fromtimestamp(os.path.getmtime(path)), 'mtime'

def scan(roots, apply=False):
    print("="*60)
    print("Smart Rename 磁盘智能重命名")
    print(f"根目录: {roots}")
    print("="*60)
    results = []
    total = 0
    for root in roots:
        root = os.path.abspath(os.path.expanduser(root))
        if not os.path.exists(root):
            print(f"\n[跳过] 不存在: {root}")
            continue
        print(f"\n扫描 {root} ...")
        for dirpath, dirs, files in os.walk(root):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            for fn in files:
                full = os.path.join(dirpath, fn)
                ext = Path(fn).suffix.lower()
                total += 1
                if total % 3000 == 0: print(f"  已处理 {total}...")
                if ext in DOC_EXTS:
                    if not is_poor_doc(fn):
                        results.append({'path':full,'orig':fn,'suggested':fn,'status':'SKIP','method':'already_named'})
                        continue
                    ex = {'.pdf':extract_pdf,'.docx':extract_docx,'.pptx':extract_pptx,'.xlsx':extract_xlsx,'.xls':extract_xlsx}[ext]
                    extracted, method = ex(full)
                    if extracted:
                        if is_invoice(extracted):
                            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(full))
                            new = f"电子发票_{mtime.strftime('%Y%m%d')}{ext}"
                            method = 'invoice_date'
                        else:
                            new = extracted + ext
                        if new.lower() != fn.lower():
                            results.append({'path':full,'orig':fn,'suggested':new,'status':'RENAME','method':method})
                        else:
                            results.append({'path':full,'orig':fn,'suggested':fn,'status':'SKIP','method':method})
                    else:
                        results.append({'path':full,'orig':fn,'suggested':fn,'status':'FAILED','method':'no_content'})
                elif ext in IMG_EXTS:
                    if not is_poor_img(fn):
                        results.append({'path':full,'orig':fn,'suggested':fn,'status':'SKIP','method':'already_named'})
                        continue
                    dt, src = get_img_dt(full)
                    new = f"{dt.strftime('%Y%m%d_%H%M%S')}{ext}"
                    if new.lower() != fn.lower():
                        results.append({'path':full,'orig':fn,'suggested':new,'status':'RENAME','method':f'img_{src}'})
                    else:
                        results.append({'path':full,'orig':fn,'suggested':fn,'status':'SKIP','method':'img'})
    renames = [r for r in results if r['status']=='RENAME']
    skips = [r for r in results if r['status']=='SKIP']
    fails = [r for r in results if r['status']=='FAILED']
    img_r = [r for r in renames if r['method'].startswith('img_')]
    doc_r = [r for r in renames if not r['method'].startswith('img_')]
    print(f"\n{'='*60}")
    print(f"扫描完成! 共 {total} 个文件")
    print(f"  待重命名: {len(renames)}  (图片 {len(img_r)}, 文档 {len(doc_r)})")
    print(f"  已命名良好, 跳过: {len(skips)}")
    print(f"  提取失败, 保留原名: {len(fails)}")
    print(f"\n{'='*60}")
    print("重命名预览 (前 20):")
    for r in renames[:20]:
        print(f"  {r['orig']}")
        print(f"    -> {r['suggested']}  [{r['method']}]")
    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
    csv_path = os.path.join(script_dir, 'rename_plan.csv')
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=['path','orig','suggested','status','method'])
        w.writeheader(); w.writerows(results)
    print(f"\n完整计划已保存: {csv_path}")
    if not apply:
        print("\n当前为预览模式, 未执行任何重命名。")
        print("确认无误后, 加 --apply 参数执行")
        return len(renames)
    print(f"\n{'='*60}")
    print("执行重命名中...")
    print(f"{'='*60}")
    ok = fail = skip = collision = 0
    fail_list = []
    for r in renames:
        src = r['path']
        sug = r['suggested']
        if not os.path.exists(src): skip += 1; continue
        src_dir = os.path.dirname(src)
        target = os.path.join(src_dir, sug)
        counter = 1
        while os.path.exists(target) and target != src:
            stem, ext = os.path.splitext(sug)
            target = os.path.join(src_dir, f"{stem}_{counter}{ext}")
            counter += 1
            collision += 1
        if target == src: skip += 1; continue
        try:
            os.rename(src, target)
            ok += 1
        except Exception as ex:
            fail += 1
            if len(fail_list) < 20: fail_list.append((src, target, str(ex)))
        if ok % 500 == 0 and ok > 0:
            print(f"  进度: OK={ok}  FAIL={fail}")
    print(f"\n{'='*60}")
    print(f"执行完成!")
    print(f"  成功: {ok}")
    print(f"  失败: {fail}")
    print(f"  跳过: {skip}")
    print(f"  自动重名后缀: {collision}")
    print(f"{'='*60}")
    log_path = os.path.join(script_dir, 'rename_log.txt')
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"\n=== {datetime.datetime.now()} ===\n")
        f.write(f"ROOTS={roots}\n")
        f.write(f"OK={ok} FAIL={fail} SKIP={skip} COLLISIONS={collision}\n")
        for s,t,e in fail_list: f.write(f"FAIL|{s}|{t}|{e}\n")
    if fail_list:
        print("\n失败文件 (前 10):")
        for s,t,e in fail_list[:10]: print(f"  {os.path.basename(s)}")
    print(f"\n日志: {log_path}")
    return ok

if __name__ == '__main__':
    p = argparse.ArgumentParser(description='Smart Rename 磁盘文件智能重命名工具')
    p.add_argument('-r','--roots', nargs='+', required=True, help='要扫描的根目录, 可多个')
    p.add_argument('--apply', action='store_true', help='执行重命名 (默认仅预览)')
    args = p.parse_args()
    ensure_libs()
    scan(roots=args.roots, apply=args.apply)
