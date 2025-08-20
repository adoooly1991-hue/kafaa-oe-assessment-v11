
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

def t_i18n(key, lang, i18n):
    try:
        return i18n.get(lang, {}).get(key, i18n.get('en', {}).get(key, key))
    except Exception:
        return key

def _brand_header(slide, brand_primary="#C00000", logo_path=None):
    try:
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.18))
        band.fill.solid(); 
        h = brand_primary.lstrip('#')
        band.fill.fore_color.rgb = RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
        band.line.fill.background()
        if logo_path and os.path.exists(logo_path):
            slide.shapes.add_picture(logo_path, Inches(8.2), Inches(0.2), height=Inches(0.6))
    except Exception:
        pass

def add_material_flow_narrative_slide(prs, text: str, lang='en', i18n=None, brand_primary="#C00000", logo_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    _brand_header(slide, brand_primary, logo_path)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title.text_frame.text = t_i18n("material_flow", lang, i18n or {})
    title.text_frame.paragraphs[0].font.size = Pt(26)
    box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(4.5))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    tf.text = text
    tf.paragraphs[0].font.size = Pt(12)

def export_observations_pptx(observations_df, out_path, steps=None, perstep_top2=None, spacing_mode="Effective CT", ct_eff_map=None, vc_summary=None, material_flow_text=None, photos=None, template_path=None, lang='en', i18n=None, brand_primary="#C00000", logo_path=None):
    default_master = None
    try:
        # Try to load brand pptx master from templates.yaml via env hint
        import yaml
        if os.path.exists('templates.yaml'):
            with open('templates.yaml','r', encoding='utf-8') as _f:
                _tpl = yaml.safe_load(_f) or {}
                default_master = (_tpl.get('brand',{}) or {}).get('pptx_master')
    except Exception:
        default_master = None
    master_to_use = template_path or default_master
    prs = Presentation(master_to_use) if (master_to_use and os.path.exists(master_to_use)) else Presentation()
    title = prs.slides.add_slide(prs.slide_layouts[0])
    _brand_header(title, brand_primary, logo_path)
    if title.shapes.title:
        title.shapes.title.text = t_i18n("title", lang, i18n or {})
    if hasattr(title, "placeholders") and len(title.placeholders)>1:
        try: title.placeholders[1].text = ""
        except Exception: pass

    if steps and perstep_top2:
        add_current_state_map_slide(prs, steps, perstep_top2, spacing_mode=spacing_mode, ct_eff_map=ct_eff_map or {}, lang=lang, i18n=i18n, brand_primary=brand_primary, logo_path=logo_path)
    if vc_summary:
        add_value_chain_slide(prs, vc_summary, lang=lang, i18n=i18n, brand_primary=brand_primary, logo_path=logo_path)
    if material_flow_text:
        add_material_flow_narrative_slide(prs, material_flow_text, lang=lang, i18n=i18n, brand_primary=brand_primary, logo_path=logo_path)

    add_pqcdsm_slides(prs, observations_df, lang=lang, i18n=i18n, brand_primary=brand_primary, logo_path=logo_path)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    _brand_header(slide, brand_primary, logo_path)
    tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    tf = tx.text_frame; tf.text = t_i18n("summary_top", lang, i18n or {})
    tf.paragraphs[0].font.size = Pt(28)
    y = 1.2
    for i, row in observations_df.head(8).iterrows():
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(9), Inches(0.6))
        p = tb.text_frame.paragraphs[0]
        p.text = f"{row['step_name']} — {row['waste'].title()} | Score {row['score_0_5']:.1f} | RPN {row['rpn_pct']:.0f}% | {row.get('evidence','')}"
        p.font.size = Pt(14)
        y += 0.6

    for _, row in observations_df.iterrows():
        s = prs.slides.add_slide(prs.slide_layouts[5])
        _brand_header(s, brand_primary, logo_path)
        header = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        header.text_frame.text = f"{row['step_name']} — {row['waste'].title()}"
        header.text_frame.paragraphs[0].font.size = Pt(26)
        meta = s.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(0.6))
        meta.text_frame.text = f"Score {row['score_0_5']:.1f} | RPN {row['rpn_pct']:.0f}% | Evidence: {row.get('evidence','')}"
        meta.text_frame.paragraphs[0].font.size = Pt(14)
        body = s.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(6.5), Inches(3))
        body.text_frame.text = row['observation']
        body.text_frame.paragraphs[0].font.size = Pt(18)
        # Photos
        try:
            if photos:
                key = (str(row.get('step_id','')), str(row.get('waste','')).lower())
                files = photos.get(key, [])[:2]
                px = Inches(7.2); py = Inches(2.0); ph = Inches(1.9)
                for i, fp in enumerate(files):
                    if os.path.exists(fp):
                        s.shapes.add_picture(fp, px, py + Inches(i*2.1), height=ph)
        except Exception:
            pass

    prs.save(out_path)
    return out_path

def add_pqcdsm_slides(prs, observations_df, lang='en', i18n=None, brand_primary="#C00000", logo_path=None):
    theme_order = [("P","Production"),("Q","Quality"),("C","Cost"),("D","Delivery"),("S","Safety"),("M","Morale")]
    if "theme_code" not in observations_df.columns:
        def cat(w):
            w=str(w).lower()
            if w in ("overproduction","overprocessing","motion","transportation"): return "P"
            if w in ("defects",): return "Q"
            if w in ("inventory",): return "C"
            if w in ("waiting",): return "D"
            if w in ("safety",): return "S"
            if w in ("talent",): return "M"
            return "P"
        observations_df = observations_df.copy()
        observations_df["theme_code"] = observations_df["waste"].apply(cat)
    for code, name in theme_order:
        grp = observations_df[observations_df["theme_code"]==code]
        if grp.empty: 
            continue
        s = prs.slides.add_slide(prs.slide_layouts[5])
        _brand_header(s, brand_primary, logo_path)
        title = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title.text_frame.text = f"{code} — {name}: {t_i18n('pqcdsm_obs', lang, i18n or {})}"
        title.text_frame.paragraphs[0].font.size = Pt(26)
        y = 1.2
        for idx, row in enumerate(grp.itertuples(index=False), start=1):
            tb = s.shapes.add_textbox(Inches(0.5), Inches(y), Inches(9), Inches(0.6))
            tf = tb.text_frame; tf.clear()
            p = tf.paragraphs[0]
            p.text = f"{code}-{idx} {row.step_name} — {row.waste.title()}"
            p.font.size = Pt(14); p.font.bold = True
            q = tf.add_paragraph()
            q.text = getattr(row, "observation", "")
            q.level = 1; q.font.size = Pt(12)
            y += 0.8
            if y > 6.5:
                y = 1.2
                s = prs.slides.add_slide(prs.slide_layouts[5])
                _brand_header(s, brand_primary, logo_path)

def add_current_state_map_slide(prs, steps, perstep_top2, spacing_mode="Effective CT", ct_eff_map=None, lang='en', i18n=None, brand_primary="#C00000", logo_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    _brand_header(slide, brand_primary, logo_path)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title.text_frame.text = t_i18n("csm", lang, i18n or {})
    title.text_frame.paragraphs[0].font.size = Pt(28)
    info_y = Inches(0.9); mat_y = Inches(5.2)
    info_lane = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), info_y, Inches(9.2), Inches(0.5))
    info_lane.fill.solid(); info_lane.fill.fore_color.rgb = RGBColor(230,230,230)
    info_lane.text_frame.text = "Information Flow"
    mat_lane = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), mat_y, Inches(9.2), Inches(0.5))
    mat_lane.fill.solid(); mat_lane.fill.fore_color.rgb = RGBColor(230,230,230)
    mat_lane.text_frame.text = "Material Flow"
    start_x = Inches(0.5); box_w=Inches(2.6); box_h=Inches(2.4); box_y=Inches(2.0); gap = Inches(1.0)
    positions=[]; x=start_x
    for i,s in enumerate(steps):
        rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, box_y, box_w, box_h)
        rect.fill.solid(); rect.fill.fore_color.rgb = RGBColor(198, 239, 206); rect.line.color.rgb = RGBColor(91,155,213)
        tf = rect.text_frame; tf.clear()
        p = tf.paragraphs[0]; p.text = f"{s.id} – {s.name}"; p.font.size=Pt(14); p.font.bold=True
        for line in [f"CT: {int(s.ct_sec or 0)} s", f"WIP: {int(s.wip_units_in or 0)}", f"Defects: {getattr(s,'defect_pct',0):.1f}%", f"Mode: {getattr(s,'push_pull','')}"]:
            q = tf.add_paragraph(); q.text=line; q.level=1; q.font.size=Pt(11)
        positions.append((x, box_y)); x = x + box_w + gap
    for i in range(len(positions)-1):
        (x1,y1) = positions[i]; (x2,y2)=positions[i+1]
        conn = slide.shapes.add_connector(1, int(x1+box_w), int(y1+box_h/2), int(x2), int(y2+box_h/2))
        conn.line.width=Pt(2); conn.line.color.rgb = RGBColor(0,0,0)

def add_value_chain_slide(prs, vc_summary, lang='en', i18n=None, brand_primary="#C00000", logo_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    _brand_header(slide, brand_primary, logo_path)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title.text_frame.text = t_i18n("vc", lang, i18n or {})
    title.text_frame.paragraphs[0].font.size = Pt(26)
    x = Inches(0.5); y = Inches(1.2); w = Inches(2.3); h = Inches(0.9); gap = Inches(0.4)
    for i, st in enumerate(vc_summary):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(230,242,255); box.line.color.rgb = RGBColor(91,155,213)
        tf = box.text_frame; tf.clear(); tf.text = st["stage_name"]; tf.paragraphs[0].font.size=Pt(12); tf.paragraphs[0].font.bold=True
        if i < len(vc_summary)-1:
            slide.shapes.add_connector(1, int(x+w), int(y+h/2), int(x+w+gap), int(y+h/2)).line.width=Pt(2)
        tb = slide.shapes.add_textbox(x, y+h+Inches(0.1), w, Inches(1.5))
        tf2 = tb.text_frame; tf2.clear(); tf2.text = "Top 3 Mudas"; tf2.paragraphs[0].font.size=Pt(11); tf2.paragraphs[0].font.bold=True
        for (m, sc) in st.get("top3", [])[:3]:
            p = tf2.add_paragraph(); p.text = f"• {m.title()} ({sc:.1f})"; p.font.size=Pt(10); p.level=1
        issues = st.get("issues", [])[:2]
        if issues:
            p = tf2.add_paragraph(); p.text = "Examples:"; p.font.size=Pt(10); p.level=1
            for iss in issues:
                q = tf2.add_paragraph(); q.text = f"– {iss}"; q.font.size=Pt(9); q.level=2
        x = x + w + gap

def export_observations_pdf(observations_df, out_path, brand_primary="#C00000", logo_path="assets/kafaa_logo.png"):
    c = canvas.Canvas(out_path, pagesize=landscape(A4))
    w, h = landscape(A4)
    # Watermark on each page
    def _watermark():
        try:
            from reportlab.lib.utils import ImageReader
            c.saveState()
            c.translate(w*0.5, h*0.35)
            c.rotate(25)
            c.setFillAlpha(0.08)
            img = ImageReader(logo_path) if (logo_path and os.path.exists(logo_path)) else None
            if img:
                c.drawImage(img, -w*0.25, -h*0.15, width=w*0.5, height=h*0.3, mask='auto', preserveAspectRatio=True)
            c.restoreState()
        except Exception:
            pass
    _watermark()
    w, h = landscape(A4)
    c.setFont("Helvetica-Bold", 20); c.drawString(2*cm, h-1.5*cm, "Automated VSM – Observations")
    c.setFont("Helvetica-Bold", 12); y = h-3.0*cm
    for _, row in observations_df.iterrows():
        c.drawString(1.5*cm, y, f"{row['step_name']} — {row['waste'].title()} (Score {row['score_0_5']:.1f} | RPN {row['rpn_pct']:.0f}% | {row.get('evidence','')})")
        y -= 0.7*cm; c.setFont("Helvetica", 11)
        for line in split_text(row['observation'], max_chars=140):
            c.drawString(2.0*cm, y, line); y -= 0.55*cm
            if y < 2.0*cm: c.showPage(); y = h-2.0*cm; c.setFont("Helvetica", 11)
        c.setFont("Helvetica-Bold", 12); y -= 0.3*cm
        if y < 3.0*cm: c.showPage(); y = h-3.0*cm; c.setFont("Helvetica-Bold", 12)
    _watermark(); c.showPage(); c.save(); return out_path

def split_text(text, max_chars=100):
    words = text.split(); out=[]; cur=""
    for w in words:
        if len(cur)+len(w)+1 <= max_chars: cur=(cur+" "+w).strip()
        else: out.append(cur); cur=w
    if cur: out.append(cur)
    return out
