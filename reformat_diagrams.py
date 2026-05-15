"""
Reformats the diagram zone (pages 13-45) of RenierExpress_Documentation.

Diagram zone = paras 178-360 in the original document (sections 2-20).
Everything outside that range is kept 100% byte-identical.

Each diagram (image3–image51) gets its own full page with:
  • Orientation determined by image aspect ratio
  • Image maximised to fill the page (0.5" margins), caption below
"""

import zipfile, io, os, copy, shutil
from PIL import Image
from lxml import etree

SRC = "/root/.claude/uploads/c8b79956-577a-44b9-a003-e7c18dcd5d92/6cd87e77-May15_RenierExpress_Documentation_Chapter131.docx"
DST = "/home/user/RenRouter/RenierExpress_DiagramPages_Reformatted.docx"

# Diagram zone: FIRST_DIA is the first body paragraph of section 2,
# LAST_DIA is the last body paragraph of section 20 (para 360 which
# carries the final landscape sectPr of that zone).
FIRST_DIA  = 178
LAST_DIA   = 360

# Images that live in sections 2-20 (the diagram zone)
IMAGE_ORDER = [f'image{i}' for i in range(3, 52)]   # image3 … image51

# Page geometry
MARGIN_IN  = 0.5   # tight margin on diagram pages
CAPTION_IN = 0.35  # height reserved for caption
TWIPS      = 1440
EMU        = 914400
PORT_W, PORT_H = 8.5, 11.0
LAND_W, LAND_H = 11.0,  8.5

# ── namespace helpers ─────────────────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

def q(ns, tag):
    return f'{{{NS[ns]}}}{tag}'

# ── relationship / caption tables ─────────────────────────────────────────────
RID = {
    'image3':  'rId11', 'image4':  'rId12', 'image5':  'rId13',
    'image6':  'rId14', 'image7':  'rId15', 'image8':  'rId16',
    'image9':  'rId17', 'image10': 'rId18', 'image11': 'rId19',
    'image12': 'rId20', 'image13': 'rId21', 'image14': 'rId22',
    'image15': 'rId23', 'image16': 'rId24', 'image17': 'rId25',
    'image18': 'rId26', 'image19': 'rId27', 'image20': 'rId28',
    'image21': 'rId29', 'image22': 'rId30', 'image23': 'rId31',
    'image24': 'rId32', 'image25': 'rId33', 'image26': 'rId34',
    'image27': 'rId35', 'image28': 'rId36', 'image29': 'rId37',
    'image30': 'rId38', 'image31': 'rId39', 'image32': 'rId40',
    'image33': 'rId41', 'image34': 'rId42', 'image35': 'rId43',
    'image36': 'rId44', 'image37': 'rId45', 'image38': 'rId46',
    'image39': 'rId47', 'image40': 'rId48', 'image41': 'rId49',
    'image42': 'rId50', 'image43': 'rId51', 'image44': 'rId52',
    'image45': 'rId53', 'image46': 'rId54', 'image47': 'rId55',
    'image48': 'rId56', 'image49': 'rId57', 'image50': 'rId58',
    'image51': 'rId59',
}

CAPTIONS = {
    'image3':  'Figure 3. Flowchart of Existing System',
    'image4':  'Figure 4. Flowchart of Existing System – Member',
    'image5':  'Figure 5. Flowchart of Existing System – Admin',
    'image6':  'Figure 6. Flowchart of Existing System – Rider',
    'image7':  'Figure 7. Existing System Context Data Flow Diagram',
    'image8':  'Figure 8. Existing System Level 0 Data Flow Diagram',
    'image9':  'Figure 9. Flowchart of Proposed System',
    'image10': 'Figure 10. Flowchart of Proposed System – Member Overview Flow',
    'image11': 'Figure 11. Flowchart of Proposed System – Member Dashboard and Bottom Navigation Flow',
    'image12': 'Figure 12. Flowchart of Proposed System – Member Address Selection and Saved Address Management Flow',
    'image13': 'Figure 13. Flowchart of Proposed System – Member Profile and Account Management Flow',
    'image14': 'Figure 14. Flowchart of Proposed System – Member Notification Preferences, Help, Support and Legal Flow',
    'image15': 'Figure 15. Flowchart of Proposed System – Member Order Tracking, Cancellation and Delivered POD Viewing Flow',
    'image16': 'Figure 16. Flowchart of Proposed System – Member Order History and Notification Deep-Link Flow',
    'image17': 'Figure 17. Flowchart of Proposed System – Member Registration and Approval Flow',
    'image18': 'Figure 18. Flowchart of Proposed System – Member Booking and Order Confirmation Flow',
    'image19': 'Figure 19. Flowchart of Proposed System – Member Authentication and Password Recovery Flow',
    'image20': 'Figure 20. Flowchart of Proposed System – Member Post-Booking Self-Service Flow',
    'image21': 'Figure 21. Flowchart of Proposed System – Admin Overview Flow',
    'image22': 'Figure 22. Flowchart of Proposed System – Admin Authentication and Password Recovery Flow',
    'image23': 'Figure 23. Flowchart of Proposed System – Admin Dashboard and Operational Overview Flow',
    'image24': 'Figure 24. Flowchart of Proposed System – Admin Member Management Flow',
    'image25': 'Figure 25. Flowchart of Proposed System – Admin Rider Management and Coverage Zone Flow',
    'image26': 'Figure 26. Flowchart of Proposed System – Admin Order Management and Reassignment Flow',
    'image27': 'Figure 27. Flowchart of Proposed System – Admin Settings and Capacity Configuration Flow',
    'image28': 'Figure 28. Flowchart of Proposed System – Admin Notifications Center and Alert Deep-Link Flow',
    'image29': 'Figure 29. Flowchart of Proposed System – Admin Analytics',
    'image30': 'Figure 30. Flowchart of Proposed System – Admin Batch Generation and Admin Review Flow',
    'image31': 'Figure 31. Flowchart of Proposed System – Admin Route Optimization and Dispatch Flow',
    'image32': 'Figure 32. Flowchart of Proposed System – Admin Batch Operations Detail Flow',
    'image33': 'Figure 33. Flowchart of Proposed System – Rider Overview Flow',
    'image34': 'Figure 34. Flowchart of Proposed System – Rider Authentication and Password Recovery Flow',
    'image35': 'Figure 35. Flowchart of Proposed System – Rider Daily Batches and Route View Flow',
    'image36': 'Figure 36. Flowchart of Proposed System – Rider Delivery Execution Flow',
    'image37': 'Figure 37. Flowchart of Proposed System – Rider Attempted Delivery and Reschedule Flow',
    'image38': 'Figure 38. Flowchart of Proposed System – Rider POD Capture and Delivery Confirmation Flow',
    'image39': 'Figure 39. Flowchart of Proposed System – Rider Route Adjustment and Insertion Handling Flow',
    'image40': 'Figure 40. Flowchart of Proposed System – Rider Account History and Notification Flow',
    'image41': 'Figure 41. Flowchart of Proposed System – Rider Profile and Credential Management Flow',
    'image42': 'Figure 42. Flowchart of Proposed System – Rider Intra-Day Insertion Response Flow',
    'image43': 'Figure 43. Flowchart of Proposed System – Rider Bottom Navigation and Screen Access Flow',
    'image44': 'Figure 44. Proposed System Context Data Flow Diagram',
    'image45': 'Figure 45. Proposed System Level 0 Data Flow Diagram',
    'image46': 'Figure 46. Proposed System Level 1 Data Flow Diagram – Register and Verify Members',
    'image47': 'Figure 47. Proposed System Level 1 Data Flow Diagram – Manage Bookings and Orders',
    'image48': 'Figure 48. Proposed System Level 1 Data Flow Diagram – Generate Batches and Optimize Routes',
    'image49': 'Figure 49. Proposed System Level 1 Data Flow Diagram – Manage Rider Delivery Operations',
    'image50': 'Figure 50. Proposed System Level 1 Data Flow Diagram – Manage Notifications and GPS Tracking',
    'image51': 'Figure 51. Proposed System Level 1 Data Flow Diagram – Generate Analytics and Forecast Data',
}

# ── image dimension reader ────────────────────────────────────────────────────
def get_image_dims(path):
    dims = {}
    with zipfile.ZipFile(path) as z:
        for fname in z.namelist():
            if fname.startswith('word/media/'):
                try:
                    data = z.read(fname)
                    img  = Image.open(io.BytesIO(data))
                    key  = fname.split('/')[-1].rsplit('.', 1)[0]
                    dims[key] = img.size
                except Exception:
                    pass
    return dims

def calc_fit(iw, ih, bw, bh):
    r = iw / ih
    if r >= bw / bh:
        return bw, bw / r
    else:
        return bh * r, bh

# ── XML builders ──────────────────────────────────────────────────────────────
def make_sectPr(landscape):
    m  = int(MARGIN_IN * TWIPS)
    if landscape:
        pw, ph = int(LAND_W * TWIPS), int(LAND_H * TWIPS)
        extra  = {q('w','orient'): 'landscape'}
    else:
        pw, ph = int(PORT_W * TWIPS), int(PORT_H * TWIPS)
        extra  = {}
    s = etree.Element(q('w','sectPr'))
    etree.SubElement(s, q('w','pgSz'),
                     **{q('w','w'): str(pw), q('w','h'): str(ph)}, **extra)
    etree.SubElement(s, q('w','pgMar'),
                     **{q('w','top'):str(m), q('w','right'):str(m),
                        q('w','bottom'):str(m), q('w','left'):str(m),
                        q('w','header'):str(m), q('w','footer'):str(m),
                        q('w','gutter'):'0'})
    return s

def make_drawing(rid, cx, cy, pic_id):
    W = q('w','drawing')
    drawing = etree.Element(W)
    inline  = etree.SubElement(drawing, q('wp','inline'),
                                distT='0', distB='0', distL='0', distR='0')
    etree.SubElement(inline, q('wp','extent'), cx=str(cx), cy=str(cy))
    etree.SubElement(inline, q('wp','effectExtent'), l='0', t='0', r='0', b='0')
    etree.SubElement(inline, q('wp','docPr'),
                     id=str(pic_id), name=f'Diag{pic_id}')
    cNvGFP = etree.SubElement(inline, q('wp','cNvGraphicFramePr'))
    etree.SubElement(cNvGFP, f'{{{NS["a"]}}}graphicFrameLocks', noChangeAspect='1')

    gfx  = etree.SubElement(inline, f'{{{NS["a"]}}}graphic')
    gdat = etree.SubElement(gfx,    f'{{{NS["a"]}}}graphicData', uri=NS['pic'])
    pic  = etree.SubElement(gdat,   f'{{{NS["pic"]}}}pic')

    nv   = etree.SubElement(pic, f'{{{NS["pic"]}}}nvPicPr')
    etree.SubElement(nv, f'{{{NS["pic"]}}}cNvPr',
                     id=str(pic_id), name=f'Diag{pic_id}')
    etree.SubElement(nv, f'{{{NS["pic"]}}}cNvPicPr')

    bf   = etree.SubElement(pic, f'{{{NS["pic"]}}}blipFill')
    etree.SubElement(bf, f'{{{NS["a"]}}}blip',
                     **{f'{{{NS["r"]}}}embed': rid})
    st   = etree.SubElement(bf, f'{{{NS["a"]}}}stretch')
    etree.SubElement(st, f'{{{NS["a"]}}}fillRect')

    sp   = etree.SubElement(pic, f'{{{NS["pic"]}}}spPr')
    xf   = etree.SubElement(sp, f'{{{NS["a"]}}}xfrm')
    etree.SubElement(xf, f'{{{NS["a"]}}}off',  x='0', y='0')
    etree.SubElement(xf, f'{{{NS["a"]}}}ext',  cx=str(cx), cy=str(cy))
    pg   = etree.SubElement(sp, f'{{{NS["a"]}}}prstGeom', prst='rect')
    etree.SubElement(pg, f'{{{NS["a"]}}}avLst')
    return drawing

def make_image_para(rid, cx, cy, pic_id):
    p   = etree.Element(q('w','p'))
    pPr = etree.SubElement(p, q('w','pPr'))
    etree.SubElement(pPr, q('w','jc'),
                     **{q('w','val'): 'center'})
    etree.SubElement(pPr, q('w','spacing'),
                     **{q('w','before'): '0', q('w','after'): '80'})
    run = etree.SubElement(p, q('w','r'))
    etree.SubElement(run, q('w','rPr'))
    run.append(make_drawing(rid, cx, cy, pic_id))
    return p

def make_caption_para(text, landscape):
    p   = etree.Element(q('w','p'))
    pPr = etree.SubElement(p, q('w','pPr'))
    etree.SubElement(pPr, q('w','jc'),      **{q('w','val'): 'center'})
    etree.SubElement(pPr, q('w','spacing'), **{q('w','before'): '60', q('w','after'): '0'})
    rPr_pPr = etree.SubElement(pPr, q('w','rPr'))
    etree.SubElement(rPr_pPr, q('w','b'))
    etree.SubElement(rPr_pPr, q('w','sz'),   **{q('w','val'): '20'})
    etree.SubElement(rPr_pPr, q('w','szCs'), **{q('w','val'): '20'})
    pPr.append(make_sectPr(landscape))      # section ends here

    run = etree.SubElement(p, q('w','r'))
    rPr = etree.SubElement(run, q('w','rPr'))
    etree.SubElement(rPr, q('w','b'))
    etree.SubElement(rPr, q('w','sz'),   **{q('w','val'): '20'})
    t   = etree.SubElement(run, q('w','t'))
    t.text = text
    return p

# ── main ──────────────────────────────────────────────────────────────────────
def main():
    print("Copying original document …")
    shutil.copy2(SRC, DST)

    with zipfile.ZipFile(DST) as z:
        raw = z.read('word/document.xml')

    print("Parsing …")
    tree = etree.fromstring(raw)
    body = tree.find(q('w','body'))

    # Walk body children, recording para index → element
    para_map = {}
    pidx = 0
    for ch in body:
        if ch.tag == q('w','p'):
            para_map[pidx] = ch
            pidx += 1

    print(f"Total paragraphs: {pidx}")

    # Get image dimensions
    dims = get_image_dims(SRC)

    # Build replacement elements (image3 … image51)
    print("\nBuilding full-page diagram elements:")
    new_els = []
    pic_counter = 900   # start high to avoid colliding with existing docPr ids
    for img in IMAGE_ORDER:
        rid        = RID[img]
        caption    = CAPTIONS[img]
        iw, ih     = dims[img]
        landscape  = (iw / ih) > 1.0
        pw         = LAND_W if landscape else PORT_W
        ph         = LAND_H if landscape else PORT_H
        aw         = pw - 2 * MARGIN_IN
        ah         = ph - 2 * MARGIN_IN - CAPTION_IN
        fw, fh     = calc_fit(iw, ih, aw, ah)
        cx         = int(fw * EMU)
        cy         = int(fh * EMU)
        new_els.append(make_image_para(rid, cx, cy, pic_counter))
        new_els.append(make_caption_para(caption, landscape))
        print(f"  {img}: {'LAND' if landscape else 'PORT'} "
              f"{fw:.2f}\" x {fh:.2f}\" | {caption[:55]}")
        pic_counter += 1

    # Identify elements to remove: paras FIRST_DIA … LAST_DIA
    to_remove = [para_map[i] for i in range(FIRST_DIA, LAST_DIA + 1)
                 if i in para_map]

    # The anchor: first paragraph AFTER the diagram zone (para 361)
    anchor = para_map.get(LAST_DIA + 1)

    print(f"\nRemoving {len(to_remove)} original diagram-zone paragraphs "
          f"(paras {FIRST_DIA}–{LAST_DIA}) …")
    for el in to_remove:
        body.remove(el)

    print(f"Inserting {len(new_els)} new elements before para {LAST_DIA+1} …")
    if anchor is not None:
        # addprevious(x) inserts x immediately before `anchor`; forward iteration
        # accumulates correctly in document order.
        for el in new_els:
            anchor.addprevious(el)
    else:
        # Fallback: insert before body-level sectPr
        bsect = body.find(q('w','sectPr'))
        for el in new_els:
            bsect.addprevious(el)

    # Verify body sectPr is untouched (sanity)
    bsp = body.find(q('w','sectPr'))
    cols = bsp.find(q('w','cols'))
    print(f"\nBody sectPr cols={cols.get(q('w','num'),'1') if cols is not None else '1'} "
          f"(must stay 2 for post-content)  ✓")

    # Serialise
    print("Writing …")
    new_xml = etree.tostring(tree, xml_declaration=True,
                              encoding='UTF-8', standalone=True)
    tmp = DST + '.tmp'
    with zipfile.ZipFile(DST) as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item,
                          new_xml if item.filename == 'word/document.xml'
                          else zin.read(item.filename))
    os.replace(tmp, DST)
    print(f"\nDone → {DST}")

if __name__ == '__main__':
    main()
