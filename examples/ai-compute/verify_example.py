"""Narrow package/render-text verifier; never certifies diagram meaning or aesthetics."""
import argparse, hashlib, json, re, zipfile
from pathlib import Path
import pymupdf
from pptx import Presentation

def shapes(items):
    for shape in items:
        yield shape
        if shape.shape_type == 6:
            yield from shapes(shape.shapes)

def norm(value):
    return re.sub(r'\s+', '', value)

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--pptx',type=Path,required=True)
    p.add_argument('--pdf',type=Path,required=True,help='Actual render of the supplied PPTX, not a separate SVG preview.')
    p.add_argument('--report',type=Path,required=True)
    p.add_argument('--preview-dir',type=Path)
    args=p.parse_args()
    result={'scope':'AI-compute six-slide example only','semantic_review':'unchecked','visual_review':'unchecked','complete':False}
    try:
        result['pptx_sha256']=hashlib.sha256(args.pptx.read_bytes()).hexdigest()
        result['pdf_sha256']=hashlib.sha256(args.pdf.read_bytes()).hexdigest()
        result['render_provenance']='Caller must establish this PDF was rendered from the supplied PPTX; matching text alone cannot prove provenance.'
        with zipfile.ZipFile(args.pptx) as z:
            if z.testzip() is not None:raise ValueError('Corrupt ZIP member')
        prs=Presentation(args.pptx)
        with pymupdf.open(args.pdf) as pdf:
            if len(prs.slides)!=6 or len(pdf)!=6:raise ValueError('Expected six PPTX and PDF pages')
            if abs(prs.slide_width/prs.slide_height-16/9)>.001:raise ValueError('Expected 16:9')
            pages=[]
            if args.preview_dir:args.preview_dir.mkdir(parents=True,exist_ok=True)
            for i,slide in enumerate(prs.slides):
                ss=list(shapes(slide.shapes));actual=norm(pdf[i].get_text())
                missing=[s.text for s in ss if s.has_text_frame and norm(s.text) and norm(s.text) not in actual]
                notes=slide.has_notes_slide and len(slide.notes_slide.notes_text_frame.text)>100
                pages.append({'page':i+1,'missing_render_text':missing,'notes_present':notes,'text_objects':sum(bool(s.text) for s in ss if s.has_text_frame),'picture_objects':sum(s.shape_type==13 for s in ss),'native_component_count':sum(s.shape_type not in (6,13) for s in ss)})
                if args.preview_dir:pdf[i].get_pixmap(matrix=pymupdf.Matrix(1.34,1.34)).save(args.preview_dir/f'page-{i+1}.png')
            result['pages']=pages
            result['automated_status']='passed' if all(not x['missing_render_text'] and x['notes_present'] and x['picture_objects']==0 for x in pages) else 'failed'
    except Exception as e:
        result['automated_status']='failed';result['error']=str(e)
    args.report.parent.mkdir(parents=True,exist_ok=True)
    args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f"Automated checks: {result['automated_status']}; semantic/visual review: unchecked; report: {args.report}")
    return 0 if result['automated_status']=='passed' else 1

if __name__=='__main__':
    raise SystemExit(main())
