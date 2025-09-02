# ============================================================
# Nombre del proyecto: elBullistatistics
# Archivo: pdf.py
# Descripción: Módulo encargado de generar el informe en PDF a partir 
#              de las capturas de pantalla proporcionadas, colocando 
#              cada imagen de forma centrada en su propia página.
# Fecha de creación: Julio - Agosto 2025
# ============================================================

# --------------------------
# Librerías
# --------------------------

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Image, PageBreak, NextPageTemplate)
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Flowable
from reportlab.lib.utils import ImageReader

# --------------------------
# Fuentes y estilos
# --------------------------

pdfmetrics.registerFont(TTFont("Raleway", "fonts/Raleway-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Raleway-Bold", "fonts/Raleway-Bold.ttf"))
styles = getSampleStyleSheet()
H1 = styles["Heading1"]
N  = styles["BodyText"]

# --------------------------
# Clase auxiliar
# --------------------------

class CenteredImage(Flowable):
    
    def __init__(self, path, draw_w=None, draw_h=None, scale=None, fit=False):
        
        super().__init__()
        self.path = path
        self.draw_w = draw_w
        self.draw_h = draw_h
        self.scale = scale
        self.fit = fit
        self._dw = None
        self._dh = None
        self.availW = None
        self.availH = None

    def wrap(self, availWidth, availHeight):
        
        img = ImageReader(self.path)
        iw, ih = img.getSize()

        if self.scale is not None:
            dw, dh = iw * self.scale, ih * self.scale
        elif self.draw_w and self.draw_h:
            r = min(self.draw_w / iw, self.draw_h / ih)
            dw, dh = iw * r, ih * r
        elif self.draw_w:
            r = self.draw_w / iw
            dw, dh = iw * r, ih * r
        elif self.draw_h:
            r = self.draw_h / ih
            dw, dh = iw * r, ih * r
        else:
            dw, dh = iw, ih

        if self.fit:
            rfit = min(availWidth / dw, availHeight / dh, 1.0)
            dw, dh = dw * rfit, dh * rfit

        self._dw, self._dh = dw, dh
        self.availW, self.availH = availWidth, availHeight
        return (availWidth, availHeight)

    def draw(self):

        x = (self.availW - self._dw) / 2 - 1*cm
        y = (self.availH - self._dh) / 2
        self.canv.drawImage(
            self.path, x, y, width=self._dw, height=self._dh,
            preserveAspectRatio=True, mask='auto'
        )

# --------------------------
# Funciones de dibujo
# --------------------------

def draw_cover(canvas, doc, image_path: str, file_name: str, image_offset_cm: float = -1.0, text_from_bottom_cm: float = 5.8):
    
    canvas.saveState()
    width, height = doc.pagesize

    img = ImageReader(image_path)
    iw, ih = img.getSize()
    max_w = width * 1.2
    max_h = height * 0.9
    scale = min(max_w / iw, max_h / ih)
    dw, dh = iw * scale, ih * scale
    image_y = (height - dh) / 2 + (image_offset_cm * cm)
    image_x = (width - dw) / 2
    canvas.drawImage(image_path, image_x, image_y, width=dw, height=dh, preserveAspectRatio=True, mask='auto')

    canvas.setFont("Raleway-Bold", 28)
    text_y = max(1.0*cm, min(text_from_bottom_cm*cm, height - 1.5*cm))
    canvas.drawCentredString(width / 2, text_y, file_name)
    canvas.restoreState()

def draw_common_elements(canvas, doc, file_name: str, total_visitors: int, title: str = "Estadísticas elBulli1846"):

    canvas.saveState()
    width, height = doc.pagesize

    canvas.setFont("Raleway-Bold", 14)
    canvas.drawCentredString(width / 2, height - 1.2*cm, title)
    canvas.setFont("Raleway-Bold", 10)
    canvas.drawString(2*cm, height - 1.2*cm, "Fecha:")
    off = canvas.stringWidth("Fecha:", "Raleway-Bold", 10) + 4
    canvas.setFont("Raleway", 10)
    canvas.drawString(2*cm + off, height - 1.2*cm, str(file_name))

    try:
        img = ImageReader("resources/water.png")
        iw, ih = img.getSize()
        target_h = 1.2*cm
        scale = target_h / ih
        dw, dh = iw * scale, ih * scale
        x = width - 2*cm - dw
        y = height - 1.2*cm - (dh * 0.5)
        canvas.drawImage("resources/water.png", x, y, width=dw, height=dh, preserveAspectRatio=True, mask='auto')
    except Exception:
        pass

    vis_str = f"{int(total_visitors):,}".replace(",", ".")
    label = "Visitantes totales:"
    canvas.setFont("Raleway-Bold", 10)
    lw = canvas.stringWidth(label, "Raleway-Bold", 10)
    vw = canvas.stringWidth(vis_str, "Raleway", 10)
    start_x = (width - (lw + vw + 4)) / 2
    canvas.drawString(start_x, 1.2*cm, label)
    canvas.setFont("Raleway", 10)
    canvas.drawString(start_x + lw + 4, 1.2*cm, vis_str)

    shown_page = doc.page - 1
    canvas.drawRightString(width - 2*cm, 1.2*cm, f"{shown_page}")
    canvas.restoreState()

def draw_image_page(canvas, doc, image_path):

    canvas.saveState()
    width, height = doc.pagesize

    img = ImageReader(image_path)
    iw, ih = img.getSize()

    max_w, max_h = 30*cm, 15.5*cm
    scale = min(max_w / iw, max_h / ih, 1)
    dw, dh = iw * scale, ih * scale

    x = (width - dw) / 2
    y = (height - dh) / 2
    canvas.drawImage(image_path, x, y, width=dw, height=dh, preserveAspectRatio=True, mask='auto')

    canvas.restoreState()

# --------------------------
# Generación del documento
# --------------------------

def build_pdf(output_path, titulo="Informe", file_name="Informe", total_visitors=0, page_images=None):

    frame_cover = Frame(0*cm, 0*cm, 1000*cm, 1000*cm, id="COVER")
    frame_body  = Frame(2*cm, 2.5*cm, 27.7*cm, 16*cm, id="BODY")

    cover_tpl = PageTemplate(
        id="COVER",
        frames=[frame_cover],
        onPage=lambda c, d: draw_cover(c, d, "resources/cover.jpg", file_name)
    )
    body_tpl = PageTemplate(
        id="BODY",
        frames=[frame_body],
        onPage=lambda c, d: draw_common_elements(
            c, d, file_name=file_name, total_visitors=total_visitors, title="Estadísticas elBulli1846"
        )
    )

    doc = BaseDocTemplate(
        output_path,
        pagesize=landscape(A4),
        pageTemplates=[cover_tpl, body_tpl],
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2.5*cm, bottomMargin=2.5*cm
    )

    story = []
    story.append(Paragraph(" ", N))
    story.append(NextPageTemplate("BODY"))
    story.append(PageBreak())

    if page_images:
        for i, p in enumerate(page_images):
            story.append(CenteredImage(p, draw_w=26*cm))
            if i < len(page_images) - 1:
                story.append(PageBreak())

    doc.build(story)