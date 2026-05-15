"""
Reformats RenierExpress documentation so each diagram occupies its own full page
with appropriate orientation (landscape/portrait) and maximized image size.
Covers pages 13–45 (all diagram figures, image1–image80).
"""

import zipfile
import io
import os
import copy
import shutil
from PIL import Image
from lxml import etree

# ── Paths ────────────────────────────────────────────────────────────────────
SRC  = "/root/.claude/uploads/c8b79956-577a-44b9-a003-e7c18dcd5d92/6cd87e77-May15_RenierExpress_Documentation_Chapter131.docx"
DST  = "/home/user/RenRouter/RenierExpress_DiagramPages_Reformatted.docx"

# ── Page geometry (twips: 1 inch = 1440; EMU: 1 inch = 914400) ───────────────
MARGIN_IN   = 0.5          # tight margin on diagram pages
CAPTION_IN  = 0.35         # vertical space reserved for caption text
TWIPS       = 1440
EMU         = 914400

PORT_W_IN, PORT_H_IN  = 8.5, 11.0
LAND_W_IN, LAND_H_IN  = 11.0,  8.5

# ── rId → image filename map (from document.xml.rels) ────────────────────────
RID_TO_IMG = {
    'rId9':  'image1',  'rId10': 'image2',  'rId11': 'image3',
    'rId12': 'image4',  'rId13': 'image5',  'rId14': 'image6',
    'rId15': 'image7',  'rId16': 'image8',  'rId17': 'image9',
    'rId18': 'image10', 'rId19': 'image11', 'rId20': 'image12',
    'rId21': 'image13', 'rId22': 'image14', 'rId23': 'image15',
    'rId24': 'image16', 'rId25': 'image17', 'rId26': 'image18',
    'rId27': 'image19', 'rId28': 'image20', 'rId29': 'image21',
    'rId30': 'image22', 'rId31': 'image23', 'rId32': 'image24',
    'rId33': 'image25', 'rId34': 'image26', 'rId35': 'image27',
    'rId36': 'image28', 'rId37': 'image29', 'rId38': 'image30',
    'rId39': 'image31', 'rId40': 'image32', 'rId41': 'image33',
    'rId42': 'image34', 'rId43': 'image35', 'rId44': 'image36',
    'rId45': 'image37', 'rId46': 'image38', 'rId47': 'image39',
    'rId48': 'image40', 'rId49': 'image41', 'rId50': 'image42',
    'rId51': 'image43', 'rId52': 'image44', 'rId53': 'image45',
    'rId54': 'image46', 'rId55': 'image47', 'rId56': 'image48',
    'rId57': 'image49', 'rId58': 'image50', 'rId59': 'image51',
    'rId60': 'image52', 'rId61': 'image53', 'rId62': 'image54',
    'rId63': 'image55', 'rId64': 'image56', 'rId65': 'image57',
    'rId66': 'image58', 'rId67': 'image59', 'rId68': 'image60',
    'rId69': 'image61', 'rId70': 'image62', 'rId71': 'image63',
    'rId72': 'image64', 'rId73': 'image65', 'rId74': 'image66',
    'rId75': 'image67', 'rId76': 'image68', 'rId77': 'image69',
    'rId78': 'image70', 'rId79': 'image71', 'rId80': 'image72',
    'rId81': 'image73', 'rId82': 'image74', 'rId83': 'image75',
    'rId84': 'image76', 'rId85': 'image77', 'rId86': 'image78',
    'rId87': 'image79', 'rId88': 'image80',
}
IMG_TO_RID = {v: k for k, v in RID_TO_IMG.items()}

# ── Caption text for each image ───────────────────────────────────────────────
CAPTIONS = {
    'image1':  'Figure 1. Software Development Life Cycle',
    'image2':  'Figure 2. Conceptual Paradigm',
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
    'image52': 'Figure 52. Member Registration Form',
    'image53': 'Figure 53. Login Screen',
    'image54': 'Figure 54. Member Home Dashboard',
    'image55': 'Figure 55 (a). Member Booking Form',
    'image56': 'Figure 55 (b). Member Booking Form',
    'image57': 'Figure 56 (a). Member Booking Details',
    'image58': 'Figure 56 (b). Member Booking Details',
    'image59': 'Figure 57. Member Orders Screen',
    'image60': 'Figure 58. Member Profile Screen',
    'image61': 'Figure 59. Member Notifications Screen',
    'image62': 'Figure 60. Admin Dashboard',
    'image63': 'Figure 61. Batch Management Screen',
    'image64': 'Figure 62 (a). Batch Detail Screen',
    'image65': 'Figure 62 (b). Batch Detail Screen',
    'image66': 'Figure 63. Generate Batches Screen',
    'image67': 'Figure 64. Confirm Rider Availability Screen',
    'image68': 'Figure 64 (a). Batch Generation Review Screen',
    'image69': 'Figure 65. Generating Batches Screen',
    'image70': 'Figure 66. Member Management Screen',
    'image71': 'Figure 67. Rider Management Screen',
    'image72': 'Figure 68. Order Management Screen',
    'image73': 'Figure 69. Settings Screen',
    'image74': 'Figure 70. Today\'s Batches Screen',
    'image75': 'Figure 71. Route View Screen',
    'image76': 'Figure 72. Delivery Execution Screen',
    'image77': 'Figure 73. Proof of Delivery Capture Screen',
    'image78': 'Figure 74. Delivery History Screen',
    'image79': 'Figure 75. Rider Notifications Screen',
    'image80': 'Figure 76. Rider Profile Screen',
}

# ── Ordered list of images to process ────────────────────────────────────────
IMAGE_ORDER = [f'image{i}' for i in range(1, 81)]

# ── Namespace map ─────────────────────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wp14':'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
}
for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)

def qn(ns, tag):
    return f'{{{NS[ns]}}}{tag}'


def get_image_dims(src_path):
    dims = {}
    with zipfile.ZipFile(src_path, 'r') as z:
        for fname in z.namelist():
            if fname.startswith('word/media/'):
                try:
                    data = z.read(fname)
                    img  = Image.open(io.BytesIO(data))
                    key  = fname.split('/')[-1].rsplit('.', 1)[0]
                    dims[key] = img.size  # (width_px, height_px)
                except Exception:
                    pass
    return dims


def calc_fit(img_w, img_h, max_w_in, max_h_in):
    """Return (width_in, height_in) fitting image within bounding box, preserving aspect ratio."""
    ratio = img_w / img_h
    box_r = max_w_in / max_h_in
    if ratio >= box_r:
        w = max_w_in
        h = w / ratio
    else:
        h = max_h_in
        w = h * ratio
    return w, h


def make_sect_pr(landscape: bool) -> etree._Element:
    """Build a w:sectPr element with tight margins and the requested orientation."""
    m = int(MARGIN_IN * TWIPS)
    if landscape:
        pw = int(LAND_W_IN * TWIPS)
        ph = int(LAND_H_IN * TWIPS)
        orient_attr = {qn('w', 'orient'): 'landscape'}
    else:
        pw = int(PORT_W_IN * TWIPS)
        ph = int(PORT_H_IN * TWIPS)
        orient_attr = {}

    sectPr = etree.Element(qn('w', 'sectPr'))
    pgSz   = etree.SubElement(sectPr, qn('w', 'pgSz'),
                               **{qn('w', 'w'): str(pw), qn('w', 'h'): str(ph)},
                               **orient_attr)
    pgMar  = etree.SubElement(sectPr, qn('w', 'pgMar'),
                               **{qn('w', 'top'):    str(m),
                                  qn('w', 'right'):  str(m),
                                  qn('w', 'bottom'): str(m),
                                  qn('w', 'left'):   str(m),
                                  qn('w', 'header'): str(m),
                                  qn('w', 'footer'): str(m),
                                  qn('w', 'gutter'): '0'})
    return sectPr


def make_inline_drawing(rid: str, cx_emu: int, cy_emu: int, pic_id: int) -> etree._Element:
    """Build a w:drawing with wp:inline containing the picture at the given size."""
    drawing = etree.Element(qn('w', 'drawing'))
    inline  = etree.SubElement(drawing, qn('wp', 'inline'),
                                **{'distT': '0', 'distB': '0', 'distL': '0', 'distR': '0'})

    extent = etree.SubElement(inline, qn('wp', 'extent'),
                               cx=str(cx_emu), cy=str(cy_emu))
    effExt = etree.SubElement(inline, qn('wp', 'effectExtent'),
                               l='0', t='0', r='0', b='0')
    docPr  = etree.SubElement(inline, qn('wp', 'docPr'),
                               id=str(pic_id), name=f'Picture {pic_id}')
    cNvGFP = etree.SubElement(inline, qn('wp', 'cNvGraphicFramePr'))
    gfLock = etree.SubElement(cNvGFP,
                               f'{{{NS["a"]}}}graphicFrameLocks',
                               noChangeAspect='1')

    graphic     = etree.SubElement(inline, f'{{{NS["a"]}}}graphic')
    graphicData = etree.SubElement(graphic, f'{{{NS["a"]}}}graphicData',
                                   uri=NS['pic'])

    pic_el  = etree.SubElement(graphicData, f'{{{NS["pic"]}}}pic')
    nvPicPr = etree.SubElement(pic_el, f'{{{NS["pic"]}}}nvPicPr')
    cNvPr   = etree.SubElement(nvPicPr, f'{{{NS["pic"]}}}cNvPr',
                                id=str(pic_id), name=f'Picture {pic_id}')
    cNvPicPr = etree.SubElement(nvPicPr, f'{{{NS["pic"]}}}cNvPicPr')

    blipFill = etree.SubElement(pic_el, f'{{{NS["pic"]}}}blipFill')
    blip     = etree.SubElement(blipFill, f'{{{NS["a"]}}}blip',
                                 **{f'{{{NS["r"]}}}embed': rid})
    stretch  = etree.SubElement(blipFill, f'{{{NS["a"]}}}stretch')
    fillRect = etree.SubElement(stretch,  f'{{{NS["a"]}}}fillRect')

    spPr = etree.SubElement(pic_el, f'{{{NS["pic"]}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{NS["a"]}}}xfrm')
    off  = etree.SubElement(xfrm, f'{{{NS["a"]}}}off', x='0', y='0')
    ext  = etree.SubElement(xfrm, f'{{{NS["a"]}}}ext',
                             cx=str(cx_emu), cy=str(cy_emu))
    prstG = etree.SubElement(spPr, f'{{{NS["a"]}}}prstGeom', prst='rect')
    avLst = etree.SubElement(prstG, f'{{{NS["a"]}}}avLst')

    return drawing


def make_image_paragraph(rid: str, cx_emu: int, cy_emu: int, pic_id: int) -> etree._Element:
    """Centered paragraph containing the inline image."""
    p    = etree.Element(qn('w', 'p'))
    pPr  = etree.SubElement(p, qn('w', 'pPr'))
    jc   = etree.SubElement(pPr, qn('w', 'jc'), **{qn('w', 'val'): 'center'})
    spac = etree.SubElement(pPr, qn('w', 'spacing'),
                             **{qn('w', 'before'): '0', qn('w', 'after'): '120'})
    run  = etree.SubElement(p, qn('w', 'r'))
    rPr  = etree.SubElement(run, qn('w', 'rPr'))
    noP  = etree.SubElement(rPr, qn('w', 'noProof'))
    run.append(make_inline_drawing(rid, cx_emu, cy_emu, pic_id))
    return p


def make_caption_paragraph(text: str, sect_pr: etree._Element) -> etree._Element:
    """Centered bold paragraph with caption text + section break."""
    p   = etree.Element(qn('w', 'p'))
    pPr = etree.SubElement(p, qn('w', 'pPr'))
    jc  = etree.SubElement(pPr, qn('w', 'jc'), **{qn('w', 'val'): 'center'})
    spac = etree.SubElement(pPr, qn('w', 'spacing'),
                              **{qn('w', 'before'): '60', qn('w', 'after'): '0'})
    rPr_pPr = etree.SubElement(pPr, qn('w', 'rPr'))
    b_pPr   = etree.SubElement(rPr_pPr, qn('w', 'b'))
    sz_pPr  = etree.SubElement(rPr_pPr, qn('w', 'sz'),   **{qn('w', 'val'): '20'})
    szCs    = etree.SubElement(rPr_pPr, qn('w', 'szCs'), **{qn('w', 'val'): '20'})
    pPr.append(copy.deepcopy(sect_pr))

    run = etree.SubElement(p, qn('w', 'r'))
    rPr = etree.SubElement(run, qn('w', 'rPr'))
    b   = etree.SubElement(rPr, qn('w', 'b'))
    sz  = etree.SubElement(rPr, qn('w', 'sz'),   **{qn('w', 'val'): '20'})
    t   = etree.SubElement(run, qn('w', 't'))
    t.text = text
    return p


def make_section_break_para(landscape: bool) -> etree._Element:
    """An empty paragraph whose sole purpose is to end the previous section."""
    p   = etree.Element(qn('w', 'p'))
    pPr = etree.SubElement(p, qn('w', 'pPr'))
    pPr.append(make_sect_pr(landscape))
    return p


def build_diagram_pages(image_dims):
    """
    Returns a list of lxml elements (paragraphs) representing all diagram pages.
    Each diagram = [image_para, caption_para_with_sectPr].
    """
    pages = []
    pic_counter = 1

    for img_name in IMAGE_ORDER:
        rid     = IMG_TO_RID[img_name]
        caption = CAPTIONS[img_name]
        w_px, h_px = image_dims[img_name]

        landscape = (w_px / h_px) > 1.0

        if landscape:
            page_w = LAND_W_IN
            page_h = LAND_H_IN
        else:
            page_w = PORT_W_IN
            page_h = PORT_H_IN

        avail_w = page_w  - 2 * MARGIN_IN
        avail_h = page_h  - 2 * MARGIN_IN - CAPTION_IN

        fit_w, fit_h = calc_fit(w_px, h_px, avail_w, avail_h)

        cx_emu = int(fit_w * EMU)
        cy_emu = int(fit_h * EMU)

        sect_pr = make_sect_pr(landscape)

        img_para  = make_image_paragraph(rid, cx_emu, cy_emu, pic_counter)
        cap_para  = make_caption_paragraph(caption, sect_pr)

        pages.append(img_para)
        pages.append(cap_para)
        pic_counter += 1
        print(f"  {img_name}: {'LANDSCAPE' if landscape else 'PORTRAIT'} "
              f"{fit_w:.2f}\" x {fit_h:.2f}\"  ({caption[:50]})")

    return pages


def make_pre_content_section_end(original_sect_pr_177: etree._Element) -> etree._Element:
    """
    Build the sectPr that will close the pre-content section at para 165.
    Clone the original 2-column portrait sectPr (para 177) but use nextPage
    so the first diagram starts on a fresh page.
    """
    sect = copy.deepcopy(original_sect_pr_177)
    # Remove rsid attributes to avoid conflicts
    for attr in list(sect.attrib):
        sect.attrib.pop(attr)
    # Replace continuous break type with nextPage
    typ = sect.find(qn('w', 'type'))
    if typ is not None:
        typ.set(qn('w', 'val'), 'nextPage')
    else:
        typ = etree.SubElement(sect, qn('w', 'type'))
        typ.set(qn('w', 'val'), 'nextPage')
        sect.insert(0, typ)
    return sect


def main():
    print("Loading document …")
    shutil.copy2(SRC, DST)

    with zipfile.ZipFile(DST, 'r') as z:
        doc_xml_bytes = z.read('word/document.xml')

    print("Parsing document XML …")
    tree = etree.fromstring(doc_xml_bytes)
    body = tree.find(qn('w', 'body'))

    # ── Map paragraph index → body child element ──────────────────────────────
    # Body may have w:p and w:tbl children (and final w:sectPr).
    # We walk children to map para indices.
    para_index = 0
    para_to_child = {}   # para_idx → body child element
    for child in body:
        if child.tag == qn('w', 'p'):
            para_to_child[para_index] = child
            para_index += 1
        elif child.tag == qn('w', 'tbl'):
            pass  # tables don't count as paragraphs in our index

    total_paras = para_index
    print(f"Total paragraph elements: {total_paras}")

    # ── Paragraph boundaries ──────────────────────────────────────────────────
    # Para 165: last pre-diagram paragraph ("Requirement Analysis")
    # Para 166: first diagram paragraph (image1)
    # Para 534: last diagram paragraph (Figure 76 caption)
    # Para 535: first post-diagram paragraph (empty line)
    LAST_PRE  = 165
    FIRST_DIA = 166
    LAST_DIA  = 534
    FIRST_POST = 535

    # ── Extract original sectPr from para 177 (2-col portrait, the pre-content section) ──
    para177_el = para_to_child[177]
    para177_pPr = para177_el.find(qn('w', 'pPr'))
    orig_sect_pr_177 = para177_pPr.find(qn('w', 'sectPr'))

    # ── Get image dimensions ──────────────────────────────────────────────────
    print("Reading image dimensions …")
    image_dims = get_image_dims(SRC)

    # ── Build diagram page elements ───────────────────────────────────────────
    print("Building diagram pages:")
    diagram_els = build_diagram_pages(image_dims)

    # ── Surgical body modification ────────────────────────────────────────────
    print("\nModifying document body surgically …")

    # 1. Add a section-ending sectPr to para 165 so the pre-content 2-column
    #    section closes cleanly before the diagram pages begin.
    para165_el = para_to_child[LAST_PRE]
    para165_pPr = para165_el.find(qn('w', 'pPr'))
    if para165_pPr is None:
        para165_pPr = etree.Element(qn('w', 'pPr'))
        para165_el.insert(0, para165_pPr)
    # Don't double-add if already has one (re-run safety)
    if para165_pPr.find(qn('w', 'sectPr')) is None:
        pre_end_sectPr = make_pre_content_section_end(orig_sect_pr_177)
        para165_pPr.append(pre_end_sectPr)

    # 2. Collect body children we need to remove (para 166 through para 534
    #    inclusive) and the insertion anchor (para 535's body child).
    children_to_remove = []
    for idx in range(FIRST_DIA, LAST_DIA + 1):
        if idx in para_to_child:
            children_to_remove.append(para_to_child[idx])

    anchor_el = para_to_child.get(FIRST_POST)

    # 3. Remove the diagram-zone paragraphs from body
    for el in children_to_remove:
        body.remove(el)
    print(f"  Removed {len(children_to_remove)} original diagram-zone paragraphs")

    # 4. Insert new diagram page elements immediately before the post-content anchor.
    #    If anchor is missing (shouldn't happen), append before body sectPr.
    # addprevious(el) inserts el immediately before anchor each time,
    # so iterating forward gives the correct final order before anchor.
    if anchor_el is not None:
        for el in diagram_els:
            anchor_el.addprevious(el)
    else:
        body_sect = body.find(qn('w', 'sectPr'))
        for el in diagram_els:
            body_sect.addprevious(el)
    print(f"  Inserted {len(diagram_els)} diagram page elements ({len(diagram_els)//2} figures)")

    # ── Serialise back ────────────────────────────────────────────────────────
    print("Writing new document.xml …")
    new_xml = etree.tostring(tree, xml_declaration=True,
                              encoding='UTF-8', standalone=True)

    tmp = DST + '.tmp'
    with zipfile.ZipFile(DST, 'r') as zin, \
         zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))

    os.replace(tmp, DST)
    print(f"\nDone! Saved to: {DST}")


if __name__ == '__main__':
    main()
