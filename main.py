"""
Etiquetadora SKU — Inventa Unlock
Versão Python Desktop | Impressão ZPL direta via socket TCP
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
import pandas as pd
try:
    import win32print
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
import threading
import os
import time
from dataclasses import dataclass, field
from typing import Optional
import io

# ── Configurações visuais ──────────────────────────────────────
ctk.set_appearance_mode("light")

C_NAVY   = "#000020"
C_BLUE   = "#142F5E"
C_CYAN   = "#92E8F1"
C_MID    = "#3B4EC8"
C_LIGHT  = "#ECF5FE"
C_GRAY   = "#EBEFF2"
C_WHITE  = "#FFFFFF"
C_AMBER  = "#BA7517"
C_RED    = "#A32D2D"
C_GREEN  = "#1D9E75"
C_TEXT   = "#1a1a18"
C_TEXT2  = "#4a4a48"

# ── Temas ──────────────────────────────────────────────────────
THEMES = {
    "light": {
        "bg":       "#F5F5F5",
        "surface":  "#FFFFFF",
        "surface2": "#FAFAFA",
        "border":   "#E0E0E0",
        "text":     "#1a1a18",
        "text2":    "#4a4a48",
        "text3":    "#999999",
        "tbl_alt":  "#F8FAFF",
        "row_ok":   "#F0FAF5",
        "row_warn": "#FFFDF0",
        "btn_ghost":"#F0F0F0",
        "btn_txt":  "#444444",
        "nb_tab":   "#EBEBEB",
        "nb_sel":   "#FFFFFF",
        "nb_txt":   "#142F5E",
    },
    "dark": {
        "bg":       "#282a36",
        "surface":  "#1e1f29",
        "surface2": "#21222c",
        "border":   "#44475a",
        "text":     "#f8f8f2",
        "text2":    "#6272a4",
        "text3":    "#44475a",
        "tbl_alt":  "#2d2f3f",
        "row_ok":   "#1a3a2a",
        "row_warn": "#3a2e1a",
        "btn_ghost":"#44475a",
        "btn_txt":  "#f8f8f2",
        "nb_tab":   "#21222c",
        "nb_sel":   "#282a36",
        "nb_txt":   "#8be9fd",
    }
}
_theme = "light"

def T(key):
    return THEMES[_theme][key]

# ── Modelo de dados ────────────────────────────────────────────
@dataclass
class Product:
    code: str           # código de barras
    interno: str        # SKU interno
    externo: str        # SKU externo
    needed: int = 1     # qtd necessária
    scanned: int = 0    # qtd bipada
    checked: bool = True
    descricao: str = ""  # descrição do produto
    variacao: str = ""   # variação (cor, tamanho, etc)

    @property
    def done(self):
        return self.scanned >= self.needed

    @property
    def status(self):
        if self.scanned == 0:       return "pendente"
        if self.scanned < self.needed: return "parcial"
        return "completo"

    @property
    def pct(self):
        return min(100, int((self.scanned / self.needed) * 100)) if self.needed else 0


# ── Impressão ZPL via win32print (driver Windows) ────────────
def list_printers() -> list[str]:
    """Lista impressoras instaladas no Windows."""
    if not WIN32_AVAILABLE:
        return []
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        printers = win32print.EnumPrinters(flags, None, 1)
        return [p[2] for p in printers]
    except Exception:
        return []

def get_default_printer() -> str:
    if not WIN32_AVAILABLE:
        return ""
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""

def send_zpl_win32(printer_name: str, zpl: str) -> tuple[bool, str]:
    """
    Envia ZPL para impressora Windows via modo RAW.
    O driver ZDesigner deve estar configurado para receber dados RAW (ZPL).
    """
    if not WIN32_AVAILABLE:
        return False, "win32print não está instalado. Execute: pip install pywin32"
    try:
        # Abre a impressora
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            # Modo RAW — passa ZPL direto para o firmware da Zebra
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta ZPL", None, "RAW"))
            try:
                win32print.StartPagePrinter(hPrinter)
                # Envia como bytes — crítico: sem BOM, sem \r\n extras
                data = zpl.encode("ascii", errors="replace")
                win32print.WritePrinter(hPrinter, data)
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
        return True, "Enviado com sucesso"
    except Exception as e:
        return False, str(e)


def send_zpl_socket(host: str, port: int, zpl: str, timeout: int = 5) -> tuple[bool, str]:
    """
    Alternativa: envia ZPL via socket TCP direto (para impressoras em rede).
    Funciona mesmo sem driver instalado.
    """
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(timeout)
            s.connect((host, port))
            s.sendall(zpl.encode("ascii", errors="replace"))
        return True, "Enviado com sucesso"
    except Exception as e:
        return False, str(e)


# ── Gerador ZPL ───────────────────────────────────────────────
# Tamanhos disponíveis: (largura_mm, altura_mm) → dots @ 203dpi (1mm ≈ 8dots)
LABEL_SIZES = {
    "40x25": (320, 200),  # 40mm × 25mm
    "40x15": (320, 120),  # 40mm × 15mm
}
# Bobina 2 colunas: total 90mm
# |--3mm--|--40mm--|--4mm--|--40mm--|--3mm--|
# Margem lateral = 3mm = 24 dots
# Gap entre colunas = 4mm = 32 dots
# Largura total da página = 90mm = 720 dots
COL_MARGIN  = 24   # margem lateral (3mm)
COL_GAP     = 32   # gap entre colunas (4mm)
COL_PAGE_W  = 720  # largura total bobina 2 colunas (90mm)
# Offset X de cada coluna
COL_LEFT_X  = COL_MARGIN                        # 24 dots
COL_RIGHT_X = COL_MARGIN + 320 + COL_GAP        # 376 dots

def encode_zpl(text: str) -> str:
    """Converte caracteres especiais para hex ZPL (_XX format) com ^FH."""
    result = []
    for ch in text:
        if ord(ch) > 127:
            for b in ch.encode("utf-8"):
                result.append(f"_{b:02X}")
        else:
            result.append(ch)
    return "".join(result)

def _label_fields(product: Product, x_offset: int,
                   label_w: int, label_h: int,
                   show_interno: bool, show_codbarras_rodape: bool) -> list[str]:
    """
    Campos ZPL de uma etiqueta com offset X para posicionamento em coluna.
    Adapta proporções automaticamente para 40×25mm ou 40×15mm.
    """
    ext  = encode_zpl(product.externo)
    desc = encode_zpl(product.descricao) if product.descricao else ""
    sku  = encode_zpl(product.interno)
    cod  = encode_zpl(product.code)
    x    = x_offset
    is_tall = label_h >= 180  # 40×25mm

    lines = []

    if is_tall:
        # Layout 40×25mm — barcode grande, SKU ao lado, descrição abaixo
        lines.append(f"^FO{x+25},15^BY2,,0^BCN,55,N,N^FD{ext}^FS")
        lines.append(f"^FT{x+110},98^A0N,22,22^FH^FD{ext}^FS")
        lines.append(f"^FT{x+109},98^A0N,22,22^FH^FD{ext}^FS")
        if desc:
            lines.append(f"^FO{x+22},115^A0N,18,18^FB300,2,0,L^FH^FD{desc}^FS")
        rodape_y = 175
    else:
        # Layout 40×15mm — mesmo barcode grande do 40×25, sem descrição
        # Barcode ocupa toda a altura, SKU em texto abaixo do barcode
        lines.append(f"^FO{x+4},2^BY2,,0^BCN,55,N,N^FD{ext}^FS")
        lines.append(f"^FT{x+110},78^A0N,18,14^FH^FD{ext}^FS")
        lines.append(f"^FT{x+109},78^A0N,18,14^FH^FD{ext}^FS")
        rodape_y = 95

    # Rodapé: SKU interno e/ou código de barras
    rodape_items = []
    if show_interno:
        rodape_items.append(f"SKU:{sku}")
    if show_codbarras_rodape:
        rodape_items.append(cod)
    if rodape_items:
        font_h = 18 if is_tall else 14
        font_w = 14 if is_tall else 10
        rodape_text = encode_zpl("  ".join(rodape_items))
        lines.append(f"^FO{x+4},{rodape_y}^A0N,{font_h},{font_w}^FH^FD{rodape_text}^FS")

    return lines


def build_zpl(product: Product, label_w: int = 320, label_h: int = 120,
              show_border: bool = True, show_barcode: bool = True,
              show_interno: bool = False, show_codbarras_rodape: bool = False,
              text_size: str = "normal") -> str:
    """ZPL para 1 etiqueta simples (1 coluna) — usa ^PQ para cópias."""
    lines = ["^XA", "^CI28", "^LH0,0",
             f"^PW{label_w}", f"^LL{label_h}"]
    if show_border:
        lines.append(f"^FO2,2^GB{label_w-4},{label_h-4},2^FS")
    lines.extend(_label_fields(product, 0, label_w, label_h,
                                show_interno, show_codbarras_rodape))
    qty = max(1, product.scanned)
    lines.append(f"^PQ{qty},0,1,Y^XZ")
    return "\n".join(lines)


def build_zpl_batch(products_with_qty: list, cols: int,
                    label_w: int = 320, label_h: int = 120,
                    show_border: bool = True,
                    show_interno: bool = False,
                    show_codbarras_rodape: bool = False) -> str:
    """
    Gera ZPL para múltiplos produtos em lote.
    cols=1: 1 etiqueta por ^XA
    cols=2: 2 etiquetas lado a lado por ^XA (bobina dupla)
    """
    if cols == 1:
        blocks = []
        for p in products_with_qty:
            lines = ["^XA", "^CI28", "^LH0,0",
                     f"^PW{label_w}", f"^LL{label_h}"]
            if show_border:
                lines.append(f"^FO2,2^GB{label_w-4},{label_h-4},2^FS")
            lines.extend(_label_fields(p, 0, label_w, label_h,
                                        show_interno, show_codbarras_rodape))
            lines.append("^PQ1,0,1,Y^XZ")
            blocks.append("\n".join(lines))
        return "\n".join(blocks)

    else:  # cols == 2
        blocks = []
        items = list(products_with_qty)
        for i in range(0, len(items), 2):
            left  = items[i]
            right = items[i+1] if i+1 < len(items) else None
            lines = ["^XA", "^CI28", "^LH0,0",
                     f"^PW{COL_PAGE_W}", f"^LL{label_h}"]
            # Etiqueta esquerda com margem lateral
            if show_border:
                lines.append(f"^FO{COL_LEFT_X+2},2^GB{label_w-4},{label_h-4},2^FS")
            lines.extend(_label_fields(left, COL_LEFT_X, label_w, label_h,
                                        show_interno, show_codbarras_rodape))
            # Etiqueta direita
            if right:
                if show_border:
                    lines.append(f"^FO{COL_RIGHT_X+2},2^GB{label_w-4},{label_h-4},2^FS")
                lines.extend(_label_fields(right, COL_RIGHT_X, label_w, label_h,
                                            show_interno, show_codbarras_rodape))
            lines.append("^PQ1,0,1,Y^XZ")
            blocks.append("\n".join(lines))
        return "\n".join(blocks)



# ── PDF via reportlab ─────────────────────────────────────────
def build_pdf(products: list[Product], cfg: dict) -> bytes:
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.graphics.barcode.code128 import Code128

    label_h_dots = cfg.get("label_h", 120)
    is_tall  = label_h_dots >= 180
    cols     = cfg.get("cols", 1)
    LW       = 40 * mm
    LH       = (25 if is_tall else 15) * mm
    GAP      = 4 * mm   # gap entre colunas (igual ao ZPL: 4mm)
    MARGIN   = 3 * mm   # margem lateral (igual ao ZPL: 3mm)

    # Tamanho da página
    if cols == 2:
        PAGE_W = MARGIN + LW + GAP + LW + MARGIN  # 90mm
    else:
        PAGE_W = LW
    PAGE_H = LH

    font_sizes = {"normal": 5, "medio": 6, "grande": 7}
    fsize = font_sizes.get(cfg.get("text_size", "normal"), 5)

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    # Expande produtos pela quantidade bipada
    expanded = []
    for p in products:
        expanded.extend([p] * max(1, p.scanned))

    def draw_label(product: Product, x_offset: float):
        """Desenha uma etiqueta na posição x_offset."""
        # Borda
        if cfg.get("show_border", True):
            c.setStrokeColorRGB(0.5, 0.5, 0.5)
            c.setLineWidth(0.4)
            c.rect(x_offset + 0.5*mm, 0.5*mm, LW - 1*mm, LH - 1*mm, stroke=1, fill=0)

        # SKU externo — topo
        c.setFillColorRGB(0, 0, 0.1)
        c.setFont("Helvetica-Bold", fsize + 1)
        ext = product.externo
        max_chars = max(1, int((LW - 3*mm) / ((fsize + 1) * 0.6)))
        line1 = ext[:max_chars]
        line2 = ext[max_chars:max_chars*2] if len(ext) > max_chars else None
        y_text = LH - 2.5*mm
        c.drawString(x_offset + 1.5*mm, y_text, line1)
        if line2 and is_tall:
            y_text -= (fsize + 1.5) * 0.35 * mm
            c.drawString(x_offset + 1.5*mm, y_text, line2)

        # Descrição (só 40×25mm)
        if is_tall and product.descricao:
            c.setFillColorRGB(0.1, 0.1, 0.1)
            c.setFont("Helvetica", fsize - 0.5)
            desc = product.descricao
            max_d = max(1, int((LW - 3*mm) / (fsize * 0.55)))
            d1 = desc[:max_d]
            d2 = desc[max_d:max_d*2] if len(desc) > max_d else None
            desc_y = y_text - (fsize + 1) * 0.4 * mm
            c.drawString(x_offset + 1.5*mm, desc_y, d1)
            if d2:
                c.drawString(x_offset + 1.5*mm, desc_y - fsize * 0.35 * mm, d2)

        # Rodapé
        rodape_items = []
        if cfg.get("show_interno", False):
            rodape_items.append(f"SKU: {product.interno}")
        if cfg.get("show_codbarras_rodape", False):
            rodape_items.append(product.code)
        rodape_h = 3.5*mm if rodape_items else 0

        # Código de barras (codifica SKU externo)
        if cfg.get("show_barcode", True):
            bc_top    = y_text - 1*mm
            bc_bottom = 1*mm + rodape_h
            bc_h = max(4*mm, bc_top - bc_bottom)
            try:
                bc = Code128(product.externo,
                              barHeight=bc_h,
                              barWidth=0.28*mm,
                              humanReadable=False)
                bc_w = bc.width
                x_bc = x_offset + max(1*mm, (LW - bc_w) / 2)
                bc.drawOn(c, x_bc, bc_bottom)
            except Exception:
                c.setFont("Courier", 3.5)
                c.setFillColorRGB(0, 0, 0)
                c.drawString(x_offset + 1.5*mm, bc_bottom + 1*mm, product.externo)

        # Rodapé texto
        if rodape_items:
            c.setFillColorRGB(0.3, 0.3, 0.3)
            c.setFont("Helvetica", 3.2)
            c.drawString(x_offset + 1.5*mm, 1.2*mm, "  |  ".join(rodape_items))

    # Renderiza páginas
    if cols == 1:
        for product in expanded:
            c.setPageSize((PAGE_W, PAGE_H))
            draw_label(product, 0)
            c.showPage()
    else:  # 2 colunas
        for i in range(0, len(expanded), 2):
            c.setPageSize((PAGE_W, PAGE_H))
            left  = expanded[i]
            right = expanded[i+1] if i+1 < len(expanded) else None
            draw_label(left, MARGIN)
            if right:
                draw_label(right, MARGIN + LW + GAP)
            c.showPage()

    c.save()
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
# INTERFACE GRÁFICA
# ══════════════════════════════════════════════════════════════

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Etiquetadora SKU — Inventa Unlock")
        self.geometry("1280x800")
        self.minsize(1100, 700)
        self.configure(fg_color=T("bg"))

        # Estado
        self.map_data: Optional[pd.DataFrame] = None
        self.wms_data: Optional[pd.DataFrame] = None
        self.product_list: list[Product] = []
        self.barcode_index: dict[str, int] = {}
        self.selected_printer = tk.StringVar(value="")
        self.print_type = tk.StringVar(value="pdf")
        self.text_size = tk.StringVar(value="normal")
        self.show_border = tk.BooleanVar(value=True)
        self.show_barcode = tk.BooleanVar(value=True)
        self.show_interno = tk.BooleanVar(value=False)
        self.show_codbarras_rodape = tk.BooleanVar(value=False)
        self.label_cols = tk.IntVar(value=1)
        self.label_size = tk.StringVar(value="40x15")
        self.zebra_dpi = tk.IntVar(value=203)
        self.auto_print = tk.BooleanVar(value=False)
        self.use_qty_mode = tk.BooleanVar(value=False)  # imprime qtd total sem bipar N vezes

        self._build_ui()
        self._apply_treeview_style_init()

    def _apply_treeview_style(self):
        """Força cores completas na Treeview usando tema 'clam' — funciona no Windows."""
        s = ttk.Style()
        s.theme_use("clam")  # clam permite sobrescrever todas as cores no Windows
        if _theme == "dark":
            bg       = "#1e1f29"
            fg       = "#f8f8f2"
            hdr_bg   = "#44475a"
            hdr_fg   = "#f8f8f2"
            sel_bg   = "#44475a"
            sel_fg   = "#f8f8f2"
            border   = "#44475a"
        else:
            bg       = "#ffffff"
            fg       = "#1a1a18"
            hdr_bg   = "#000020"
            hdr_fg   = "#ffffff"
            sel_bg   = "#dce8ff"
            sel_fg   = "#000020"
            border   = "#e0e0e0"

        s.configure("Custom.Treeview",
                    background=bg,
                    foreground=fg,
                    fieldbackground=bg,
                    bordercolor=border,
                    darkcolor=border,
                    lightcolor=border,
                    rowheight=30,
                    font=("Arial", 10),
                    borderwidth=0,
                    relief="flat")
        s.configure("Custom.Treeview.Heading",
                    background=hdr_bg,
                    foreground=hdr_fg,
                    font=("Arial", 10, "bold"),
                    relief="flat",
                    borderwidth=0,
                    padding=[8, 6])
        s.map("Custom.Treeview",
              background=[("selected", sel_bg)],
              foreground=[("selected", sel_fg)])
        s.map("Custom.Treeview.Heading",
              background=[("active", hdr_bg)],
              foreground=[("active", hdr_fg)])
        # Row tag colors
        row_ok   = "#1a3a2a" if _theme == "dark" else "#f0faf5"
        row_warn = "#3a2e1a" if _theme == "dark" else "#fffdf0"
        if hasattr(self, "tree"):
            self.tree.tag_configure("pendente", background=bg,     foreground=fg)
            self.tree.tag_configure("parcial",  background=row_warn, foreground=fg)
            self.tree.tag_configure("completo", background=row_ok,   foreground=fg)

    def _apply_treeview_style_init(self):
        """Versão chamada no __init__ antes da tree existir."""
        s = ttk.Style()
        s.theme_use("clam")

    # ── Layout principal ───────────────────────────────────────
    def _toggle_theme(self):
        global _theme
        _theme = "dark" if _theme == "light" else "light"
        ctk.set_appearance_mode("dark" if _theme == "dark" else "light")
        self.theme_btn.configure(text="☀ Light" if _theme == "dark" else "🌙 Dark")
        self._apply_theme()

    def _apply_theme(self):
        t = THEMES[_theme]
        self.configure(fg_color=t["bg"])
        for tab in [self.tab_map, self.tab_wms, self.tab_scan, self.tab_hist]:
            tab.configure(fg_color=t["bg"])
        style = ttk.Style()
        style.configure("TNotebook", background=t["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=t["nb_tab"], foreground=t["nb_txt"],
                         font=("Arial", 11), padding=[20, 8])
        style.map("TNotebook.Tab",
                  background=[("selected", t["nb_sel"])],
                  foreground=[("selected", "#000020" if _theme=="light" else "#92E8F1")])
        # Treeview — aplica via método dedicado com tema clam
        self._apply_treeview_style()
        # Update progress bar
        pb_color = "#bd93f9" if _theme == "dark" else C_NAVY
        self.prog_bar.configure(progress_color=pb_color, fg_color=T("border"))
        # Update scan cards bg
        for attr in ["prog_frame_ref", "scan_card_ref", "tbl_card_ref"]:
            w = getattr(self, attr, None)
            if w:
                w.configure(fg_color=T("surface"), border_color=T("border"))
        # Update prog labels
        prog_txt = "#f8f8f2" if _theme == "dark" else C_TEXT
        for lbl in [self.prog_geradas, self.prog_faltam, self.prog_pct]:
            lbl.configure(text_color=prog_txt)
        # Sidebar
        self.sidebar_ref.configure(fg_color=T("surface2"), border_color=T("border"))
        # Refresh table
        self._refresh_scan_table()

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, fg_color=C_NAVY, corner_radius=0, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="⬛ Etiquetadora SKU",
                     font=ctk.CTkFont("Arial", 17, "bold"),
                     text_color=C_WHITE).pack(side="left", padx=20)
        ctk.CTkLabel(header, text="inventa unlock",
                     font=ctk.CTkFont("Arial", 11),
                     text_color=C_CYAN).pack(side="left", padx=0)

        self.status_lbl = ctk.CTkLabel(header, text="● mapeamento  ● wms",
                                        font=ctk.CTkFont("Arial", 11),
                                        text_color="#556688")
        self.status_lbl.pack(side="right", padx=(0, 12))

        self.theme_btn = ctk.CTkButton(
            header, text="🌙 Dark", width=90, height=30,
            fg_color="transparent", hover_color="#1a2a4a",
            border_width=1, border_color="#334466",
            text_color=C_CYAN, font=ctk.CTkFont("Arial", 11),
            corner_radius=8, command=self._toggle_theme
        )
        self.theme_btn.pack(side="right", padx=(0, 8))

        # Notebook (abas)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=0, pady=0)

        style = ttk.Style()
        style.configure("TNotebook", background=T("bg"), borderwidth=0)
        style.configure("TNotebook.Tab", background=T("nb_tab"), foreground=T("nb_txt"),
                         font=("Arial", 11), padding=[20, 8])
        style.map("TNotebook.Tab",
                  background=[("selected", T("nb_sel"))],
                  foreground=[("selected", C_NAVY)])

        self.tab_map  = ctk.CTkFrame(self.notebook, fg_color=T("bg"))
        self.tab_wms  = ctk.CTkFrame(self.notebook, fg_color=T("bg"))
        self.tab_scan = ctk.CTkFrame(self.notebook, fg_color=T("bg"))
        self.tab_hist = ctk.CTkFrame(self.notebook, fg_color=T("bg"))

        self.notebook.add(self.tab_map,  text="1. Mapeamento")
        self.notebook.add(self.tab_wms,  text="2. WMS")
        self.notebook.add(self.tab_scan, text="3. Bipagem")
        self.notebook.add(self.tab_hist, text="4. Histórico")

        self._build_tab_map()
        self._build_tab_wms()
        self._build_tab_scan()
        self._build_tab_hist()

    # ── ABA 1: MAPEAMENTO ─────────────────────────────────────
    def _build_tab_map(self):
        f = self.tab_map
        ctk.CTkLabel(f, text="Tabela de Mapeamento",
                     font=ctk.CTkFont("Arial", 14, "bold"),
                     text_color=C_NAVY).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkLabel(f, text="CSV, XLS ou XLSX — SKU Interno + SKU Externo + Quantidade",
                     font=ctk.CTkFont("Arial", 11), text_color=C_TEXT2).pack(anchor="w", padx=24)

        btn_frame = ctk.CTkFrame(f, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=12)
        ctk.CTkButton(btn_frame, text="Importar arquivo", command=self._load_map,
                       fg_color=C_NAVY, hover_color=C_BLUE,
                       font=ctk.CTkFont("Arial", 12, "bold")).pack(side="left")
        self.map_status = ctk.CTkLabel(btn_frame, text="Nenhum arquivo importado",
                                        font=ctk.CTkFont("Arial", 11), text_color=C_TEXT2)
        self.map_status.pack(side="left", padx=16)

        # Seletores de colunas
        col_frame = ctk.CTkFrame(f, fg_color=T("surface2"), corner_radius=10)
        col_frame.pack(fill="x", padx=24, pady=4)

        ctk.CTkLabel(col_frame, text="Configurar colunas",
                     font=ctk.CTkFont("Arial", 11, "bold"),
                     text_color=C_TEXT2).grid(row=0, column=0, columnspan=6, sticky="w", padx=16, pady=(10,4))

        for i, (lbl, var_name) in enumerate([
            ("● SKU Interno", "map_col_interno"),
            ("● SKU Externo", "map_col_externo"),
            ("● Quantidade",  "map_col_qty"),
        ]):
            ctk.CTkLabel(col_frame, text=lbl, font=ctk.CTkFont("Arial", 11),
                          text_color=[C_GREEN, C_MID, C_AMBER][i]).grid(row=1, column=i*2, padx=(16,4), pady=(0,10), sticky="w")
            cb = ctk.CTkComboBox(col_frame, values=["—"], width=160,
                                  font=ctk.CTkFont("Arial", 11))
            cb.grid(row=1, column=i*2+1, padx=(0,16), pady=(0,10))
            setattr(self, var_name, cb)
        self.map_col_qty.set("— sem quantidade —")

        # Preview
        self.map_preview = self._make_treeview(f)
        ctk.CTkButton(f, text="Próximo →", command=lambda: self.notebook.select(1),
                       fg_color=C_NAVY, hover_color=C_BLUE,
                       font=ctk.CTkFont("Arial", 12, "bold")).pack(anchor="e", padx=24, pady=12)

    # ── ABA 2: WMS ────────────────────────────────────────────
    def _build_tab_wms(self):
        f = self.tab_wms
        ctk.CTkLabel(f, text="Tabela do WMS",
                     font=ctk.CTkFont("Arial", 14, "bold"),
                     text_color=C_NAVY).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkLabel(f, text='CSV, XLS ou XLSX — "Código do Produto" + "barra"',
                     font=ctk.CTkFont("Arial", 11), text_color=C_TEXT2).pack(anchor="w", padx=24)

        btn_frame = ctk.CTkFrame(f, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=12)
        ctk.CTkButton(btn_frame, text="Importar arquivo", command=self._load_wms,
                       fg_color=C_NAVY, hover_color=C_BLUE,
                       font=ctk.CTkFont("Arial", 12, "bold")).pack(side="left")
        self.wms_status = ctk.CTkLabel(btn_frame, text="Nenhum arquivo importado",
                                        font=ctk.CTkFont("Arial", 11), text_color=C_TEXT2)
        self.wms_status.pack(side="left", padx=16)

        col_frame = ctk.CTkFrame(f, fg_color=T("surface2"), corner_radius=10)
        col_frame.pack(fill="x", padx=24, pady=4)
        ctk.CTkLabel(col_frame, text="Configurar colunas",
                     font=ctk.CTkFont("Arial", 11, "bold"),
                     text_color=C_TEXT2).grid(row=0, column=0, columnspan=4, sticky="w", padx=16, pady=(10,4))

        for i, (lbl, var_name) in enumerate([
            ("● SKU Interno",     "wms_col_interno"),
            ("● Código de Barras","wms_col_barcode"),
        ]):
            ctk.CTkLabel(col_frame, text=lbl, font=ctk.CTkFont("Arial", 11),
                          text_color=[C_GREEN, C_AMBER][i]).grid(row=1, column=i*2, padx=(16,4), pady=(0,6), sticky="w")
            cb = ctk.CTkComboBox(col_frame, values=["—"], width=200,
                                  font=ctk.CTkFont("Arial", 11))
            cb.grid(row=1, column=i*2+1, padx=(0,16), pady=(0,6))
            setattr(self, var_name, cb)

        # Descrição e variação
        ctk.CTkLabel(col_frame, text="● Descrição do Produto", font=ctk.CTkFont("Arial", 11),
                      text_color=C_MID).grid(row=2, column=0, padx=(16,4), pady=(0,6), sticky="w")
        self.wms_col_desc = ctk.CTkComboBox(col_frame, values=["— não incluir —"], width=320,
                                              font=ctk.CTkFont("Arial", 11))
        self.wms_col_desc.grid(row=2, column=1, columnspan=3, padx=(0,16), pady=(0,10))



        self.wms_preview = self._make_treeview(f)
        ctk.CTkButton(f, text="Ir para Bipagem →", command=self._build_lookup_and_go,
                       fg_color=C_NAVY, hover_color=C_BLUE,
                       font=ctk.CTkFont("Arial", 12, "bold")).pack(anchor="e", padx=24, pady=12)

    # ── ABA 3: BIPAGEM ────────────────────────────────────────
    def _build_tab_scan(self):
        f = self.tab_scan
        f.configure(fg_color=T("bg"))

        # Barra de progresso
        prog_frame = ctk.CTkFrame(f, fg_color=T("surface"), corner_radius=12, border_width=1, border_color=T("border"))
        self.prog_frame_ref = prog_frame
        prog_frame.pack(fill="x", padx=16, pady=(12, 6))

        top_row = ctk.CTkFrame(prog_frame, fg_color="transparent")
        top_row.pack(fill="x", padx=16, pady=(10, 4))
        self.prog_geradas = ctk.CTkLabel(top_row, text="0 geradas",
                                          font=ctk.CTkFont("Arial", 12, "bold"), text_color=C_NAVY)
        self.prog_geradas.pack(side="left")
        ctk.CTkLabel(top_row, text=" | ", text_color=C_TEXT2).pack(side="left")
        self.prog_faltam = ctk.CTkLabel(top_row, text="0 faltam",
                                         font=ctk.CTkFont("Arial", 12), text_color=C_AMBER)
        self.prog_faltam.pack(side="left")
        self.prog_pct = ctk.CTkLabel(top_row, text="0%",
                                      font=ctk.CTkFont("Arial", 12, "bold"), text_color=C_MID)
        self.prog_pct.pack(side="right")

        self.prog_bar = ctk.CTkProgressBar(prog_frame, height=6,
                                            progress_color=C_NAVY, fg_color=T("border"))
        self.prog_bar.pack(fill="x", padx=16, pady=(0, 8))
        self.prog_bar.set(0)

        badges = ctk.CTkFrame(prog_frame, fg_color="transparent")
        badges.pack(fill="x", padx=16, pady=(0, 10))
        self.badge_comp  = self._badge(badges, "0 COMPLETOS", C_GREEN)
        self.badge_parc  = self._badge(badges, "0 PARCIAIS",  C_AMBER)
        self.badge_pend  = self._badge(badges, "0 PENDENTES", C_TEXT2)
        self.badge_total = self._badge(badges, "0 PRODUTOS",  C_MID)
        for b in [self.badge_comp, self.badge_parc, self.badge_pend, self.badge_total]:
            b.pack(side="left", padx=(0, 6))

        # Layout: esquerda (scanner + tabela) | direita (config)
        main = ctk.CTkFrame(f, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=6)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)

        left = ctk.CTkFrame(main, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # Scanner
        scan_card = ctk.CTkFrame(left, fg_color=T("surface"), corner_radius=12, border_width=1, border_color=T("border"))
        self.scan_card_ref = scan_card
        scan_card.pack(fill="x", pady=(0, 8))
        scan_row = ctk.CTkFrame(scan_card, fg_color="transparent")
        scan_row.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(scan_row, text="⬛ Scanner",
                     font=ctk.CTkFont("Arial", 12, "bold"),
                     text_color=C_NAVY).pack(side="left")

        self.barcode_entry = ctk.CTkEntry(scan_row, placeholder_text="Escaneie o código de barras...",
                                           font=ctk.CTkFont("Arial", 14),
                                           width=340, height=36,
                                           border_color="#CCCCCC",
                                           fg_color="white",
                                           corner_radius=8)
        self.barcode_entry.pack(side="left", padx=12)
        self.barcode_entry.bind("<Return>", lambda e: self._process_barcode())

        ctk.CTkButton(scan_row, text="✕ Limpar", width=82, height=32,
                       fg_color=T("btn_ghost"), text_color=T("btn_txt"), hover_color=T("border"),
                       font=ctk.CTkFont("Arial", 11), corner_radius=8,
                       command=self._clear_entry).pack(side="left", padx=2)
        ctk.CTkButton(scan_row, text="↺ Reset", width=82, height=32,
                       fg_color=T("btn_ghost"), text_color=T("btn_txt"), hover_color=T("border"),
                       font=ctk.CTkFont("Arial", 11), corner_radius=8,
                       command=self._reset_all).pack(side="left", padx=2)
        self.auto_btn = ctk.CTkButton(scan_row, text="Auto: Off", width=96, height=32,
                                       fg_color=T("btn_ghost"), text_color=T("btn_txt"), hover_color=T("border"),
                                       font=ctk.CTkFont("Arial", 11), corner_radius=8,
                                       command=self._toggle_auto)
        self.auto_btn.pack(side="left", padx=2)

        self.qty_mode_btn = ctk.CTkButton(scan_row, text="Qtd Total: Off", width=120, height=32,
                                           fg_color=T("btn_ghost"), text_color=T("btn_txt"), hover_color=T("border"),
                                           font=ctk.CTkFont("Arial", 11), corner_radius=8,
                                           command=self._toggle_qty_mode)
        self.qty_mode_btn.pack(side="left", padx=2)

        self.scan_error = ctk.CTkLabel(scan_card, text="",
                                        font=ctk.CTkFont("Arial", 11), text_color=C_RED,
                                        wraplength=500)
        self._scan_error_widget = self.scan_error
        self.scan_error.pack(anchor="w", padx=12, pady=(0, 4))

        # Tabela de produtos
        tbl_card = ctk.CTkFrame(left, fg_color=T("surface"), corner_radius=12, border_width=1, border_color=T("border"))
        self.tbl_card_ref = tbl_card
        tbl_card.pack(fill="both", expand=True)
        tbl_header = ctk.CTkFrame(tbl_card, fg_color=T("surface"))
        tbl_header.pack(fill="x", padx=12, pady=(10, 4))
        ctk.CTkLabel(tbl_header, text="Produtos Etiquetados",
                     font=ctk.CTkFont("Arial", 12, "bold"),
                     text_color=C_NAVY).pack(side="left")
        ctk.CTkButton(tbl_header, text="↓ Exportar CSV", width=120,
                       fg_color=C_LIGHT, text_color=C_MID, hover_color=C_GRAY,
                       font=ctk.CTkFont("Arial", 11),
                       command=self._export_csv).pack(side="right")

        cols = ("sel", "interno", "externo", "barcode", "qtd", "acoes")
        self.tree = ttk.Treeview(tbl_card, columns=cols, show="headings",
                                  height=16, style="Custom.Treeview")
        self._apply_treeview_style()

        self.tree.heading("sel",      text="✓")
        self.tree.heading("interno",  text="SKU Interno")
        self.tree.heading("externo",  text="SKU Externo")
        self.tree.heading("barcode",  text="Código de Barras")
        self.tree.heading("qtd",      text="Qtd")
        self.tree.heading("acoes",    text="Ações")
        self.tree.column("sel",      width=36,  anchor="center")
        self.tree.column("interno",  width=150, anchor="w")
        self.tree.column("externo",  width=110, anchor="w")
        self.tree.column("barcode",  width=150, anchor="w")
        self.tree.column("qtd",      width=90,  anchor="center")
        self.tree.column("acoes",    width=80,  anchor="center")

        self._apply_treeview_style()  # aplica cores corretas após criar tree

        scroll = ttk.Scrollbar(tbl_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8,0), pady=(0,8))
        scroll.pack(side="right", fill="y", pady=(0,8), padx=(0,8))
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # Sidebar direita
        sidebar = ctk.CTkScrollableFrame(main, fg_color=T("surface2"), corner_radius=12, border_width=1, border_color=T("border"), width=250)
        self.sidebar_ref = sidebar
        sidebar.grid(row=0, column=1, sticky="nsew")

        def sec(text):
            fr = ctk.CTkFrame(sidebar, fg_color="transparent", height=1)
            fr.pack(fill="x", padx=12, pady=(12,0))
            ctk.CTkLabel(sidebar, text=text, font=ctk.CTkFont("Arial", 9, "bold"),
                          text_color="#999999").pack(anchor="w", padx=12, pady=(2, 4))

        sec("CONFIGURAÇÕES")
        sec("TAMANHO DA ETIQUETA")
        size_opts = [
            ("4 × 2,5 cm — 1 coluna",  "40x25", 1),
            ("4 × 2,5 cm — 2 colunas", "40x25", 2),
            ("4 × 1,5 cm — 1 coluna",  "40x15", 1),
            ("4 × 1,5 cm — 2 colunas", "40x15", 2),
        ]
        for lbl, sz, col in size_opts:
            f_rb = ctk.CTkFrame(sidebar, fg_color="transparent")
            f_rb.pack(anchor="w", padx=16, pady=1)
            ctk.CTkRadioButton(
                f_rb, text=lbl,
                variable=self.label_size, value=f"{sz}_{col}",
                font=ctk.CTkFont("Arial", 11),
                fg_color=C_NAVY, hover_color=C_BLUE,
                command=lambda s=sz, c=col: (
                    self.label_size.set(f"{s}_{c}"),
                    self.label_cols.set(c)
                )
            ).pack(side="left")

        sec("TAMANHO DO TEXTO")
        tsz_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        tsz_frame.pack(fill="x", padx=12, pady=2)
        for lbl, val in [("Normal","normal"),("Médio","medio"),("Grande","grande")]:
            ctk.CTkRadioButton(tsz_frame, text=lbl, variable=self.text_size, value=val,
                                font=ctk.CTkFont("Arial", 11),
                                fg_color=C_NAVY, hover_color=C_BLUE).pack(side="left", padx=4)

        sec("OPÇÕES")
        ctk.CTkCheckBox(sidebar, text="Exibir bordas", variable=self.show_border,
                         font=ctk.CTkFont("Arial", 11),
                         fg_color=C_NAVY, hover_color=C_BLUE).pack(anchor="w", padx=16, pady=2)
        ctk.CTkCheckBox(sidebar, text="Código de barras (barcode)", variable=self.show_barcode,
                         font=ctk.CTkFont("Arial", 11),
                         fg_color=C_NAVY, hover_color=C_BLUE).pack(anchor="w", padx=16, pady=2)

        sec("RODAPÉ DA ETIQUETA")
        ctk.CTkLabel(sidebar, text="Exibir no rodapé:",
                      font=ctk.CTkFont("Arial", 10), text_color=C_TEXT2).pack(anchor="w", padx=16)
        ctk.CTkCheckBox(sidebar, text="SKU Interno", variable=self.show_interno,
                         font=ctk.CTkFont("Arial", 11),
                         fg_color=C_NAVY, hover_color=C_BLUE).pack(anchor="w", padx=24, pady=2)
        ctk.CTkCheckBox(sidebar, text="Código de Barras", variable=self.show_codbarras_rodape,
                         font=ctk.CTkFont("Arial", 11),
                         fg_color=C_NAVY, hover_color=C_BLUE).pack(anchor="w", padx=24, pady=2)

        sec("TIPO DE IMPRESSÃO")
        for lbl, val in [("PDF — Impressora comum","pdf"),("ZPL — Zebra térmica","zpl")]:
            ctk.CTkRadioButton(sidebar, text=lbl, variable=self.print_type, value=val,
                                font=ctk.CTkFont("Arial", 11),
                                fg_color=C_NAVY, hover_color=C_BLUE,
                                command=self._on_print_type_change).pack(anchor="w", padx=16, pady=2)

        # Config ZPL
        self.zpl_frame = ctk.CTkFrame(sidebar, fg_color=C_LIGHT, corner_radius=8)
        self.zpl_frame.pack(fill="x", padx=12, pady=(6,0))
        self.zpl_frame.pack_forget()

        ctk.CTkLabel(self.zpl_frame, text="Impressora",
                     font=ctk.CTkFont("Arial", 10), text_color=C_TEXT2).pack(anchor="w", padx=10, pady=(8,2))
        self.printer_cb = ctk.CTkComboBox(self.zpl_frame, variable=self.selected_printer,
                                           values=["Clique em Atualizar →"],
                                           font=ctk.CTkFont("Arial", 10), height=32)
        self.printer_cb.pack(fill="x", padx=10)

        dpi_row = ctk.CTkFrame(self.zpl_frame, fg_color="transparent")
        dpi_row.pack(fill="x", padx=10, pady=(6,0))
        ctk.CTkLabel(dpi_row, text="DPI", font=ctk.CTkFont("Arial", 10),
                      text_color=C_TEXT2).pack(side="left")
        dpi_cb = ctk.CTkComboBox(dpi_row, values=["203","300"], width=80,
                                   font=ctk.CTkFont("Arial", 11))
        dpi_cb.set("203")
        dpi_cb.configure(command=lambda v: self.zebra_dpi.set(int(v)))
        dpi_cb.pack(side="left", padx=8)

        self.conn_lbl = ctk.CTkLabel(self.zpl_frame, text="",
                                      font=ctk.CTkFont("Arial", 10), text_color=C_TEXT2,
                                      wraplength=200)
        self.conn_lbl.pack(anchor="w", padx=10, pady=(4,0))
        ctk.CTkButton(self.zpl_frame, text="🔄 Atualizar impressoras", height=28,
                       fg_color=C_BLUE, hover_color=C_MID,
                       font=ctk.CTkFont("Arial", 11),
                       command=self._refresh_printers).pack(fill="x", padx=10, pady=(6,4))
        ctk.CTkButton(self.zpl_frame, text="🖨 Imprimir teste", height=28,
                       fg_color=C_GRAY, text_color=C_NAVY, hover_color="#d0d0cc",
                       font=ctk.CTkFont("Arial", 11),
                       command=self._test_zpl_print).pack(fill="x", padx=10, pady=(0,8))

        # Botão imprimir
        ctk.CTkButton(sidebar, text="🖨  Imprimir Selecionados",
                       fg_color=C_NAVY, hover_color="#1a1a50", height=44,
                       font=ctk.CTkFont("Arial", 13, "bold"),
                       corner_radius=10,
                       command=self._print_selected).pack(fill="x", padx=12, pady=(16, 4))

    # ── ABA 4: HISTÓRICO ──────────────────────────────────────
    def _build_tab_hist(self):
        f = self.tab_hist
        ctk.CTkLabel(f, text="Histórico de Bipagens",
                     font=ctk.CTkFont("Arial", 14, "bold"),
                     text_color=C_NAVY).pack(anchor="w", padx=24, pady=(20, 8))
        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", padx=24)
        ctk.CTkButton(btn_row, text="↓ Exportar CSV", command=self._export_csv,
                       fg_color=C_LIGHT, text_color=C_MID, hover_color=C_GRAY,
                       font=ctk.CTkFont("Arial", 11)).pack(side="left")
        ctk.CTkButton(btn_row, text="↺ Limpar", command=self._reset_all,
                       fg_color=C_LIGHT, text_color=C_RED, hover_color=C_GRAY,
                       font=ctk.CTkFont("Arial", 11)).pack(side="left", padx=8)

        self.hist_tree = self._make_treeview(f, columns=("externo","interno","barcode","bipado","necessario","horario"),
                                              headings=["SKU Externo","SKU Interno","Código","Bipado","Necessário","Horário"])

    # ── Helpers de UI ─────────────────────────────────────────
    def _make_treeview(self, parent, columns=None, headings=None):
        frame = ctk.CTkFrame(parent, fg_color=C_WHITE, corner_radius=10)
        frame.pack(fill="both", expand=True, padx=24, pady=8)
        if columns is None:
            columns = ("col1","col2","col3","col4","col5")
        tv = ttk.Treeview(frame, columns=columns, show="headings", height=8)
        if headings:
            for col, head in zip(columns, headings):
                tv.heading(col, text=head)
                tv.column(col, width=130)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=scroll.set)
        tv.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        scroll.pack(side="right", fill="y", pady=8)
        return tv

    def _badge(self, parent, text, color):
        lbl = ctk.CTkLabel(parent, text=text,
                            font=ctk.CTkFont("Arial", 10, "bold"),
                            fg_color=C_LIGHT, text_color=color,
                            corner_radius=6, padx=8, pady=3)
        return lbl

    # ── Lógica: importação ────────────────────────────────────
    def _load_file(self, title="Selecionar arquivo"):
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Planilhas", "*.csv *.xlsx *.xls *.txt"), ("Todos", "*.*")]
        )

    def _read_file(self, path: str) -> Optional[pd.DataFrame]:
        try:
            if path.endswith(".csv") or path.endswith(".txt"):
                return pd.read_csv(path, dtype=str).fillna("")
            else:
                return pd.read_excel(path, dtype=str).fillna("")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{e}")
            return None

    def _auto_detect(self, cols, hints):
        for h in hints:
            for col in cols:
                normalized = col.lower().replace(" ","").replace("_","")
                hint_norm  = h.lower().replace(" ","").replace("_","")
                if hint_norm == normalized or hint_norm in normalized:
                    return col
        return cols[0] if cols else ""

    def _normalize(self, s):
        import unicodedata
        return unicodedata.normalize("NFD", s.lower()).encode("ascii","ignore").decode()

    def _load_map(self):
        path = self._load_file("Importar tabela de mapeamento")
        if not path: return
        df = self._read_file(path)
        if df is None: return
        self.map_data = df
        cols = list(df.columns)
        self.map_status.configure(text=f"✓ {os.path.basename(path)} — {len(df)} registros",
                                    text_color=C_GREEN)
        for cb, hints in [(self.map_col_interno, ["interno","skuint","sku","id","codigo","code"]),
                           (self.map_col_externo, ["externo","skuext","ext","destino","cliente"]),
                           (self.map_col_qty,     ["qtd","qty","quantidade","quant","etiqueta","total"])]:
            cb.configure(values=["— sem quantidade —"] + cols if cb == self.map_col_qty else cols)
            detected = self._auto_detect(cols, hints)
            cb.set(detected if detected else cols[0])
        self._refresh_preview(self.map_preview, df)

    def _load_wms(self):
        path = self._load_file("Importar tabela do WMS")
        if not path: return
        df = self._read_file(path)
        if df is None: return
        self.wms_data = df
        cols = list(df.columns)
        self.wms_status.configure(text=f"✓ {os.path.basename(path)} — {len(df)} registros",
                                    text_color=C_GREEN)
        blank = ["— não incluir —"]
        for cb, hints, extra in [
            (self.wms_col_interno, ["cod_produto","codigodoproduto","produto","interno","sku","id","codigo"], False),
            (self.wms_col_barcode, ["cod_barras","barra","barcode","codbar","ean","gtin","bar"],              False),
            (self.wms_col_desc,    ["descr_produto","descricao","descr","desc","nome","titulo","title"],      True),
        ]:
            cb.configure(values=(blank + cols) if extra else cols)
            detected = self._auto_detect(cols, hints)
            cb.set((detected if detected else (blank[0] if extra else cols[0])))
        self._refresh_preview(self.wms_preview, df)

    def _refresh_preview(self, tv: ttk.Treeview, df: pd.DataFrame, limit=8):
        tv.delete(*tv.get_children())
        cols = list(df.columns)[:6]
        tv["columns"] = cols
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=120)
        for _, row in df.head(limit).iterrows():
            tv.insert("", "end", values=[str(row[c]) for c in cols])

    # ── Lógica: lookup e bipagem ──────────────────────────────
    def _build_lookup_and_go(self):
        if self.map_data is None or self.wms_data is None:
            messagebox.showwarning("Aviso", "Importe os dois arquivos antes de continuar.")
            return
        self._build_lookup()
        self._refresh_scan_table()
        self._update_progress()
        self.notebook.select(2)
        self.barcode_entry.focus()

    def _build_lookup(self):
        ci_map = self.map_col_interno.get()
        ce_map = self.map_col_externo.get()
        cq_map = self.map_col_qty.get()
        ci_wms = self.wms_col_interno.get()
        cb_wms = self.wms_col_barcode.get()
        cd_wms = self.wms_col_desc.get()
        blank  = "— não incluir —"

        sku_to_ext: dict[str, str] = {}
        sku_needed: dict[str, int] = {}
        for _, row in self.map_data.iterrows():
            intern = str(row.get(ci_map, "")).strip()
            if not intern: continue
            sku_to_ext[intern] = str(row.get(ce_map, "")).strip()
            qty = 1
            if cq_map and cq_map in self.map_data.columns:
                try: qty = max(1, int(float(str(row.get(cq_map, 1)) or 1)))
                except: qty = 1
            sku_needed[intern] = qty

        seen: set[str] = set()
        self.product_list = []
        self.barcode_index = {}
        for _, row in self.wms_data.iterrows():
            bar    = str(row.get(cb_wms, "")).strip()
            intern = str(row.get(ci_wms, "")).strip()
            if not bar or not intern or bar in seen: continue
            ext = sku_to_ext.get(intern)
            if not ext: continue
            seen.add(bar)
            desc = str(row.get(cd_wms, "")).strip() if cd_wms and cd_wms != blank and cd_wms in self.wms_data.columns else ""
            idx = len(self.product_list)
            self.product_list.append(Product(
                code=bar, interno=intern, externo=ext,
                needed=sku_needed.get(intern, 1),
                descricao=desc
            ))
            self.barcode_index[bar] = idx

    def _process_barcode(self):
        code = self.barcode_entry.get().strip()
        if not code: return
        self.scan_error.configure(text="")

        if not self.product_list:
            self.scan_error.configure(text="Importe as tabelas nas etapas 1 e 2 primeiro.")
            return

        idx = self.barcode_index.get(code)
        if idx is None:
            self.scan_error.configure(text=f"Código '{code}' não encontrado.")
            self.barcode_entry.configure(border_color=C_RED)
            self.after(1200, lambda: self.barcode_entry.configure(border_color=("#979DA2","#565B5E")))
            return

        item = self.product_list[idx]
        if item.done:
            self.scan_error.configure(
                text=f"{item.externo} — quantidade já completa ({item.needed}/{item.needed}).")
            return

        # Modo Qtd Total: 1 bipagem = quantidade total necessária
        if self.use_qty_mode.get():
            item.scanned = item.needed
        else:
            item.scanned += 1

        self.barcode_entry.configure(border_color=C_GREEN)
        self.after(600, lambda: (
            self.barcode_entry.configure(border_color=("#979DA2","#565B5E")),
            self.barcode_entry.delete(0, "end")
        ))

        self._refresh_scan_table()
        self._update_progress()

        if self.auto_print.get():
            threading.Thread(target=self._print_item, args=(idx,), daemon=True).start()

    def _clear_entry(self):
        self.barcode_entry.delete(0, "end")
        self.scan_error.configure(text="")
        self.barcode_entry.focus()

    def _toggle_auto(self):
        self.auto_print.set(not self.auto_print.get())
        if self.auto_print.get():
            on_bg  = "#1a3a2a" if _theme == "dark" else "#E8F7F0"
            on_fg  = "#50fa7b" if _theme == "dark" else "#0F6E56"
            on_bdr = "#50fa7b" if _theme == "dark" else "#0F6E56"
            self.auto_btn.configure(text="● Auto: On", fg_color=on_bg, text_color=on_fg,
                                     border_width=1, border_color=on_bdr)
        else:
            self.auto_btn.configure(text="Auto: Off",
                                     fg_color=T("btn_ghost"), text_color=T("btn_txt"),
                                     border_width=0, border_color="transparent")

    def _toggle_qty_mode(self):
        """Liga/desliga modo Qtd Total: ao bipar, marca scanned = needed de uma vez."""
        self.use_qty_mode.set(not self.use_qty_mode.get())
        if self.use_qty_mode.get():
            on_bg  = "#1a2a3a" if _theme == "dark" else "#E6F0FE"
            on_fg  = "#8be9fd" if _theme == "dark" else "#1a4fa8"
            on_bdr = "#8be9fd" if _theme == "dark" else "#1a4fa8"
            self.qty_mode_btn.configure(text="● Qtd Total: On",
                                         fg_color=on_bg, text_color=on_fg,
                                         border_width=1, border_color=on_bdr)
        else:
            self.qty_mode_btn.configure(text="Qtd Total: Off",
                                         fg_color=T("btn_ghost"), text_color=T("btn_txt"),
                                         border_width=0, border_color="transparent")

    def _reset_all(self):
        if not messagebox.askyesno("Confirmar", "Zerar todas as bipagens?"):
            return
        for p in self.product_list:
            p.scanned = 0
        self._refresh_scan_table()
        self._update_progress()

    # ── Tabela de scan ────────────────────────────────────────
    def _refresh_scan_table(self):
        self.tree.delete(*self.tree.get_children())
        sorted_list = sorted(enumerate(self.product_list),
                              key=lambda x: (x[1].done, -x[1].scanned))
        for orig_idx, p in sorted_list:
            qtd_str = f"{p.scanned}/{p.needed}"
            sel_str = "☑" if p.checked else "☐"
            self.tree.insert("", "end",
                              iid=str(orig_idx),
                              values=(sel_str, p.interno, p.externo, p.code, qtd_str, "🖨 ↺"),
                              tags=(p.status,))

    def _on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col  = self.tree.identify_column(event.x)
        if not item: return
        idx = int(item)
        if col == "#1":  # toggle checked
            self.product_list[idx].checked = not self.product_list[idx].checked
            self._refresh_scan_table()
        elif col == "#6":  # ações
            x = event.x
            if x < self.tree.column("acoes","width") // 2 + sum(
                self.tree.column(c,"width") for c in ("sel","interno","externo","barcode","qtd")):
                self._print_item(idx)
            else:
                self.product_list[idx].scanned = 0
                self._refresh_scan_table()
                self._update_progress()

    # ── Progresso ─────────────────────────────────────────────
    def _update_progress(self):
        total_needed  = sum(p.needed  for p in self.product_list)
        total_scanned = sum(p.scanned for p in self.product_list)
        completos = sum(1 for p in self.product_list if p.done)
        parciais  = sum(1 for p in self.product_list if 0 < p.scanned < p.needed)
        pendentes = sum(1 for p in self.product_list if p.scanned == 0)
        pct = total_scanned / total_needed if total_needed else 0

        self.prog_geradas.configure(text=f"{total_scanned} geradas")
        self.prog_faltam.configure(text=f"{total_needed - total_scanned} faltam")
        self.prog_pct.configure(text=f"{int(pct*100)}%")
        self.prog_bar.set(pct)
        self.badge_comp.configure(text=f"{completos} COMPLETOS",
                                   text_color="#50fa7b" if _theme=="dark" else C_GREEN)
        self.badge_parc.configure(text=f"{parciais} PARCIAIS",
                                   text_color="#ffb86c" if _theme=="dark" else C_AMBER)
        self.badge_pend.configure(text=f"{pendentes} PENDENTES",
                                   text_color="#6272a4" if _theme=="dark" else C_TEXT2)
        self.badge_total.configure(text=f"{len(self.product_list)} PRODUTOS",
                                    text_color="#8be9fd" if _theme=="dark" else C_MID)

        # Status header
        map_dot  = "🟢" if self.map_data is not None else "⚪"
        wms_dot  = "🟢" if self.wms_data is not None else "⚪"
        self.status_lbl.configure(text=f"{map_dot} mapeamento   {wms_dot} wms")

    # ── Impressão ─────────────────────────────────────────────
    def _get_cfg(self):
        size_key = self.label_size.get()  # e.g. "40x25_2"
        parts = size_key.rsplit("_", 1)
        sz_name = parts[0] if len(parts) == 2 else "40x15"
        cols    = int(parts[1]) if len(parts) == 2 else self.label_cols.get()
        lw, lh  = LABEL_SIZES.get(sz_name, (320, 120))
        return {
            "text_size":             self.text_size.get(),
            "show_border":           self.show_border.get(),
            "show_barcode":          self.show_barcode.get(),
            "show_interno":          self.show_interno.get(),
            "show_codbarras_rodape": self.show_codbarras_rodape.get(),
            "cols":                  cols,
            "label_w":               lw,
            "label_h":               lh,
        }

    def _on_print_type_change(self):
        if self.print_type.get() == "zpl":
            self.zpl_frame.pack(fill="x", padx=12, pady=(6,0))
        else:
            self.zpl_frame.pack_forget()

    def _refresh_printers(self):
        printers = list_printers()
        if not printers:
            self.conn_lbl.configure(
                text="Nenhuma impressora encontrada.\nInstale o driver ZDesigner.",
                text_color=C_RED)
            return
        self.printer_cb.configure(values=printers)
        zebra = next((p for p in printers if "zebra" in p.lower() or "zdesigner" in p.lower()), None)
        default = get_default_printer()
        selected = zebra or default or printers[0]
        self.printer_cb.set(selected)
        self.selected_printer.set(selected)
        self.conn_lbl.configure(
            text=f"✓ {len(printers)} impressora(s) encontrada(s)",
            text_color=C_GREEN)

    def _test_zpl_print(self):
        """Envia um ZPL mínimo de teste para diagnosticar."""
        printer_name = self.selected_printer.get()
        # Fallback para impressora padrão se nenhuma selecionada
        if not printer_name or printer_name == "Clique em Atualizar →":
            printer_name = get_default_printer()
        if not printer_name:
            messagebox.showwarning("Impressora", "Nenhuma impressora encontrada.\nClique em 'Atualizar impressoras'.")
            return
        # ZPL de teste com dimensões da etiqueta selecionada
        cfg = self._get_cfg()
        lw, lh = cfg.get("label_w", 320), cfg.get("label_h", 120)
        is_tall = lh >= 180
        if is_tall:
            zpl_test = (
                f"^XA^CI28^LH0,0^PW{lw}^LL{lh}"
                f"^FO25,15^BY2,,0^BCN,55,N,N^FDTESTE123^FS"
                f"^FT110,98^A0N,22,22^FH^FDTESTE123^FS"
                f"^FO22,115^A0N,18,18^FB300,2,0,L^FH^FDEtiquetadora SKU - Inventa Unlock^FS"
                f"^PQ1,0,1,Y^XZ"
            )
        else:
            zpl_test = (
                f"^XA^CI28^LH0,0^PW{lw}^LL{lh}"
                f"^FO4,2^BY2,,0^BCN,55,N,N^FDTESTE123^FS"
                f"^FT110,78^A0N,18,14^FH^FDTESTE123^FS"
                f"^FT109,78^A0N,18,14^FH^FDTESTE123^FS"
                f"^PQ1,0,1,Y^XZ"
            )
        ok, msg = send_zpl_win32(printer_name, zpl_test)
        if ok:
            messagebox.showinfo("Teste enviado",
                f"✓ ZPL enviado para '{printer_name}'.\n\n"
                f"Se nada saiu, verifique se o driver está em modo RAW/ZPL:\n"
                f"Painel de Controle → Impressoras → Clique direito na Zebra\n"
                f"→ Preferências → Language = ZPL")
        else:
            messagebox.showerror("Erro", f"Falha ao enviar para '{printer_name}':\n{msg}")

    def _print_item(self, idx: int):
        p = self.product_list[idx]
        if not p.scanned:
            messagebox.showinfo("Aviso", "Este produto ainda não foi bipado.")
            return
        cfg = self._get_cfg()
        if self.print_type.get() == "zpl":
            self._send_zpl_items([p], cfg)
        else:
            self._print_pdf([p], cfg)

    def _print_selected(self):
        sel = [p for p in self.product_list if p.checked and p.scanned > 0]
        if not sel:
            messagebox.showwarning("Aviso", "Selecione ao menos um produto bipado para imprimir.")
            return
        cfg = self._get_cfg()
        if self.print_type.get() == "zpl":
            threading.Thread(target=self._send_zpl_items, args=(sel, cfg), daemon=True).start()
        else:
            self._print_pdf(sel, cfg)

    def _send_zpl_items(self, items: list[Product], cfg: dict):
        printer_name = self.selected_printer.get()
        if not printer_name or printer_name == "Clique em Atualizar →":
            self.after(0, lambda: messagebox.showwarning(
                "Impressora", "Selecione uma impressora.\nClique em 'Atualizar impressoras'."))
            return

        # Expande cada produto pela quantidade bipada
        expanded = []
        for item in items:
            expanded.extend([item] * max(1, item.scanned))

        cols    = cfg.get("cols", 1)
        label_w = cfg.get("label_w", 320)
        label_h = cfg.get("label_h", 120)
        zpl = build_zpl_batch(
            expanded,
            cols=cols,
            label_w=label_w,
            label_h=label_h,
            show_border=cfg["show_border"],
            show_interno=cfg["show_interno"],
            show_codbarras_rodape=cfg.get("show_codbarras_rodape", False)
        )

        ok, err_msg = send_zpl_win32(printer_name, zpl)
        total = len(expanded)
        if ok:
            if not self.auto_print.get():
                self.after(0, lambda: messagebox.showinfo("Sucesso",
                    f"✓ {total} etiqueta(s) enviada(s) para {printer_name}!"))
        else:
            self.after(0, lambda m=err_msg: messagebox.showerror("Erro ao imprimir", m))

    def _print_pdf(self, items: list[Product], cfg: dict):
        try:
            pdf_bytes = build_pdf(items, cfg)
            path = filedialog.asksaveasfilename(
                title="Salvar PDF",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile="etiquetas.pdf"
            )
            if not path: return
            with open(path, "wb") as f:
                f.write(pdf_bytes)
            os.startfile(path)  # abre no visualizador padrão
        except Exception as e:
            messagebox.showerror("Erro ao gerar PDF", str(e))

    # ── Exportar CSV ──────────────────────────────────────────
    def _export_csv(self):
        scanned = [p for p in self.product_list if p.scanned > 0]
        if not scanned:
            messagebox.showinfo("Aviso", "Nenhuma bipagem para exportar.")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="historico_bipagem.csv"
        )
        if not path: return
        df = pd.DataFrame([{
            "SKU Externo": p.externo,
            "SKU Interno": p.interno,
            "Código de Barras": p.code,
            "Bipado": p.scanned,
            "Necessário": p.needed,
        } for p in scanned])
        df.to_csv(path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("Exportado", f"✓ CSV salvo em:\n{path}")
        self._refresh_hist()

    def _refresh_hist(self):
        self.hist_tree.delete(*self.hist_tree.get_children())
        for p in self.product_list:
            if p.scanned > 0:
                self.hist_tree.insert("", "end", values=(
                    p.externo, p.interno, p.code, p.scanned, p.needed, "—"
                ))


# ── Entry point ───────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
