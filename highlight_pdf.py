import os
import pandas as pd
import fitz  # PyMuPDF

pasta_resumidos = "arquivos_por_documento"
pasta_pdf = "full_contract_pdf"
saida_pdf = "pdf_destacados"
os.makedirs(saida_pdf, exist_ok=True)

CORES = {
    "change of control": (1, 0, 0),
    "anti-assignment": (0.8, 0, 0),
    "audit rights": (0.6, 0, 0),

    "agreement date": (0, 0.4, 1),
    "agreement date-answer": (0, 0.4, 1),

    "effective date": (0, 0.6, 0.9),
    "effective date-answer": (0, 0.6, 0.9),

    "expiration date": (0, 0.8, 0.8),
    "expiration date-answer": (0, 0.8, 0.8),

    "renewal term": (0, 0.5, 0.5),
    "renewal term-answer": (0, 0.5, 0.5),

    "notice period to terminate renewal": (0.4, 0.3, 0.8),
    "notice period to terminate renewal-answer": (0.4, 0.3, 0.8),

    "document name": (0.9, 0.4, 0),
    "governing law": (0.9, 0.6, 0),
    "insurance": (0.9, 0.8, 0),

    "license grant": (0.3, 0.7, 0.3),
    "non-transferable license": (0.3, 0.7, 0.5),
    "affiliate license-licensor": (0.3, 0.7, 0.7),
    "affiliate license-licensee": (0.3, 0.7, 0.9),
    "irrevocable or perpetual license": (0.3, 0.7, 0.2),

    "liquidated damages": (0.5, 0.2, 0.2),
    "competitive restriction exception": (0.5, 0.4, 0.2),
    "non-compete": (0.5, 0.6, 0.2),

    "exclusivity": (0.6, 0.3, 0.7),
    "no-solicit of customers": (0.6, 0.5, 0.7),

    "parties": (0.2, 0.6, 0.2),
    "parties-answer": (0.2, 0.6, 0.2),

    "post-termination services": (0.2, 0.4, 0.6),
    "termination for convenience": (0.2, 0.2, 0.6),
}

def cor_para_categoria(cat):
    cat = cat.lower().strip().replace("_", " ").replace("-", " ")
    return CORES.get(cat)  # sem cor default


def extrair_texto_area(page, rect):
    clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1)
    return page.get_text("text", clip=clip).strip()


def highlight_text(page, texto, cor):
    texto = texto.strip()

    if cor is None:
        return

    if len(texto) < 20:  # bloqueio absoluto
        return

    tamanho_ref = len(texto)

    # ---------- trechos curtos ----------
    if len(texto) < 30:
        areas = page.search_for(texto)
        for a in areas:
            texto_area = extrair_texto_area(page, a)
            if len(texto_area) < max(20, tamanho_ref * 0.5):
                continue
            h = page.add_highlight_annot(a)
            h.set_colors(stroke=cor)
            h.update()
        return

    # ---------- trechos longos ----------
    window = 45
    overlaps = []

    for i in range(0, len(texto), window):
        chunk = texto[i:i+window].strip()
        if len(chunk) < 20:
            continue
        areas = page.search_for(chunk)
        for a in areas:
            overlaps.append(a)

    if not overlaps:
        return

    overlaps = sorted(overlaps, key=lambda r: (r.y0, r.x0))

    final_boxes = []
    atual = overlaps[0]

    for box in overlaps[1:]:
        if abs(box.y0 - atual.y1) < 20 or abs(box.x0 - atual.x1) < 50:
            atual = fitz.Rect(
                min(atual.x0, box.x0),
                min(atual.y0, box.y0),
                max(atual.x1, box.x1),
                max(atual.y1, box.y1)
            )
        else:
            final_boxes.append(atual)
            atual = box

    final_boxes.append(atual)

    for b in final_boxes:
        texto_area = extrair_texto_area(page, b)
        if len(texto_area) < max(20, tamanho_ref * 0.5):
            continue
        h = page.add_highlight_annot(b)
        h.set_colors(stroke=cor)
        h.update()


# ---------- processamento ----------
for pdf_nome in os.listdir(pasta_pdf):
    if not pdf_nome.lower().endswith(".pdf"):
        continue

    nome_base = os.path.splitext(pdf_nome)[0]
    xlsx_caminho = os.path.join(pasta_resumidos, nome_base + ".pdf.xlsx")

    if not os.path.exists(xlsx_caminho):
        print("Sem XLSX para:", nome_base)
        continue

    df = pd.read_excel(xlsx_caminho)
    categorias = [c for c in df.columns if c != "categoria_origem"]

    pdf_caminho = os.path.join(pasta_pdf, pdf_nome)
    doc = fitz.open(pdf_caminho)

    for _, row in df.iterrows():
        for categoria in categorias:
            trecho = str(row[categoria]).strip()
            if trecho.lower() in ["", "nan", "none"]:
                continue

            cor = cor_para_categoria(categoria)

            for page in doc:
                highlight_text(page, trecho, cor)

    # ---------- legenda ----------
    legend_page = doc.new_page(-1)
    y = 50
    legend_page.insert_text((50, 20), "Legenda das Cores", fontsize=14)

    for categoria, cor in CORES.items():
        legend_page.insert_text((50, y), categoria, fontsize=11)
        legend_page.draw_rect(fitz.Rect(150, y - 10, 250, y + 10), color=cor, fill=cor)
        y += 30

    saida = os.path.join(saida_pdf, nome_base + "_destacado.pdf")
    doc.save(saida)
    doc.close()

print("Processo concluído.")