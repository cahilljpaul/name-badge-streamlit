import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics

# =========================
# Utilities
# =========================

def apply_opacity(image, opacity):
    """Ensure image has RGBA and apply opacity (0..1)."""
    if image.mode != "RGBA":
        image = image.convert("RGBA")
    alpha = image.split()[3]
    alpha = alpha.point(lambda p: int(p * opacity))
    image.putalpha(alpha)
    return image

def image_to_reader(pil_image):
    """
    Robustly convert a PIL image to ReportLab ImageReader.
    Handles CMYK and exotic modes; preserves alpha via PNG,
    falls back to JPEG if needed.
    """
    mode = pil_image.mode
    has_alpha = "A" in mode
    try:
        if mode in ("CMYK", "YCbCr", "LAB"):
            pil_image = pil_image.convert("RGB")
            has_alpha = False
        elif mode in ("P", "L"):
            pil_image = pil_image.convert("RGBA") if has_alpha else pil_image.convert("RGB")
            has_alpha = "A" in pil_image.mode
        elif mode not in ("RGB", "RGBA"):
            pil_image = pil_image.convert("RGB")
            has_alpha = False

        buf = BytesIO()
        if has_alpha:
            pil_image.save(buf, format="PNG")
        else:
            try:
                pil_image.save(buf, format="PNG")
            except OSError:
                buf = BytesIO()
                pil_image.convert("RGB").save(buf, format="JPEG", quality=92)
        buf.seek(0)
        return ImageReader(buf)
    except Exception:
        pil_image = pil_image.convert("RGB")
        buf = BytesIO()
        pil_image.save(buf, format="JPEG", quality=92)
        buf.seek(0)
        return ImageReader(buf)

def compute_draw_size(original_size, max_w_pts, max_h_pts):
    """Fit image (px) into box (pt) preserving aspect."""
    ow, oh = original_size
    if ow == 0 or oh == 0:
        return max_w_pts, max_h_pts
    aspect = ow / oh
    if aspect >= (max_w_pts / max_h_pts):
        draw_w = max_w_pts
        draw_h = max_w_pts / aspect
    else:
        draw_h = max_h_pts
        draw_w = max_h_pts * aspect
    return draw_w, draw_h

def fit_text_to_width(text, font_name, font_size, max_width_pts):
    """Truncate text with ellipsis to not exceed width."""
    if text is None:
        return ""
    s = str(text)
    w = pdfmetrics.stringWidth(s, font_name, font_size)
    if w <= max_width_pts:
        return s
    ell = "…"
    ell_w = pdfmetrics.stringWidth(ell, font_name, font_size)
    if ell_w >= max_width_pts:
        return ""
    lo, hi = 0, len(s)
    while lo < hi:
        mid = (lo + hi) // 2
        candidate = s[:mid] + ell
        cw = pdfmetrics.stringWidth(candidate, font_name, font_size)
        if cw <= max_width_pts:
            lo = mid + 1
        else:
            hi = mid
    return s[:max(0, lo - 1)] + ell

def ranges_intersect(a1, a2, b1, b2):
    """Check 1D interval overlap."""
    return not (a2 <= b1 or b2 <= a1)

def place_by_anchor(x, y, bw, bh, draw_w, draw_h, anchor, padding, offx_pts, offy_pts):
    """Position an image inside a badge by anchor + offsets."""
    if anchor == "Top Left":
        px = x + padding + offx_pts
        py = y + bh - draw_h - padding + offy_pts
    elif anchor == "Top Right":
        px = x + bw - draw_w - padding + offx_pts
        py = y + bh - draw_h - padding + offy_pts
    elif anchor == "Bottom Left":
        px = x + padding + offx_pts
        py = y + padding + offy_pts
    elif anchor == "Bottom Right":
        px = x + bw - draw_w - padding + offx_pts
        py = y + padding + offy_pts
    else:  # Center
        px = x + (bw - draw_w) / 2 + offx_pts
        py = y + (bh - draw_h) / 2 + offy_pts
    return px, py

def largest_free_interval(L, R, blocked_intervals):
    """
    From [L,R], remove blocked intervals and return largest free (start,end).
    """
    clipped = []
    for a, b in blocked_intervals:
        left = max(L, a)
        right = min(R, b)
        if left < right:
            clipped.append((left, right))
    if not clipped:
        return (L, R)
    clipped.sort(key=lambda x: x[0])
    merged = [clipped[0]]
    for a, b in clipped[1:]:
        m_a, m_b = merged[-1]
        if a <= m_b:
            merged[-1] = (m_a, max(m_b, b))
        else:
            merged.append((a, b))
    free = []
    cur = L
    for a, b in merged:
        if cur < a:
            free.append((cur, a))
        cur = max(cur, b)
    if cur < R:
        free.append((cur, R))
    if not free:
        return (L, L)
    free.sort(key=lambda x: x[1] - x[0], reverse=True)
    return free[0]

def hex_to_rgb01(h):
    """Convert #RRGGBB -> (r,g,b) floats 0..1."""
    h = str(h).lstrip("#")
    try:
        return (int(h[0:2], 16)/255.0, int(h[2:4], 16)/255.0, int(h[4:6], 16)/255.0)
    except Exception:
        return (0, 0, 0)

# Accepted (case-insensitive, trimmed) spellings for the required columns
NAME_ALIASES = {"name", "full name", "attendee name", "delegate name"}
ORG_ALIASES = {"organisation", "organization", "company", "organisation ", "organization "}

def _norm_cell(v):
    """Lowercase + trim a header cell for matching."""
    return str(v).strip().lower()

def find_header_row(raw_df, max_scan_rows=15):
    """Find the row index containing both a Name-like and Organisation-like header, anywhere in the first rows."""
    for i in range(min(max_scan_rows, len(raw_df))):
        row_vals = [_norm_cell(v) for v in raw_df.iloc[i].tolist()]
        if any(v in NAME_ALIASES for v in row_vals) and any(v in ORG_ALIASES for v in row_vals):
            return i
    return None

def load_badge_excel(uploaded_file):
    """
    Read an uploaded Excel file, locate the header row wherever it is (not just row 1),
    and normalize the Name/Organisation column names regardless of case, spacing or
    American/British spelling. Returns (dataframe_or_None, error_message_or_None).
    """
    raw = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
    header_row = find_header_row(raw)
    if header_row is None:
        return None, (
            "Couldn't find columns for 'Name' and 'Organisation' (or 'Organization') "
            "in the first 15 rows of the sheet. Please ensure both column headers exist somewhere near the top."
        )
    columns = [str(c).strip() for c in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = columns
    data = data.dropna(how="all")

    rename_map = {}
    for col in data.columns:
        norm = _norm_cell(col)
        if norm in NAME_ALIASES:
            rename_map[col] = "Name"
        elif norm in ORG_ALIASES:
            rename_map[col] = "Organisation"
    data = data.rename(columns=rename_map)
    return data, None

# =========================
# PDF Generation
# =========================

def generate_badges_pdf(
    data, logo_image, background_image, font_type,
    name_font_size, org_font_size, extra_font_size, extra_fields,
    bold_organisation, color_column, color_mapping, border_style,
    fixed_border_color, border_thickness, logo_position,
    second_image=None, second_logo_position="Left Middle",
    bottom_image=None, bottom_image_max_height_cm=1.2,
    primary_logo_max_width_cm=3.25, primary_logo_max_height_cm=5.0,
    second_logo_max_width_cm=3.0, second_logo_max_height_cm=4.5,
    extra_image1=None, extra1_anchor="Top Left", extra1_max_width_cm=2.5,
    extra1_max_height_cm=2.5, extra1_offset_x_cm=0.0, extra1_offset_y_cm=0.0,
    extra_image2=None, extra2_anchor="Top Right", extra2_max_width_cm=2.5,
    extra2_max_height_cm=2.5, extra2_offset_x_cm=0.0, extra2_offset_y_cm=0.0,
    show_color_circle=True, text_horizontal="Left", text_vertical="Top",
    gap_name_org_cm=0.6, gap_before_first_extra_cm=1.2, gap_between_extras_cm=0.6,
    show_text_border=False, text_border_color=(0, 0, 0), text_border_thickness=0.5,
    show_name_border=False, name_border_color="#000000", name_border_thickness=1.0,
    text_box_mode="Full available width", text_box_fixed_width_cm=6.0,
    text_box_padding_cm=0.25, text_box_offset_x_cm=0.0, text_box_offset_y_cm=0.0,
    name_border_padding_cm=0.25,
    debug_draw_boxes=False
):
    """
    Create multi-badge PDF with robust layout & overlap avoidance.
    Returns BytesIO positioned at start.
    """
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    # Badge layout
    badge_width = 9.01 * cm
    badge_height = 5.51 * cm
    left_margin = 1.0 * cm
    top_margin = 1.0 * cm
    padding = 0.3 * cm
    badges_per_row = 2
    badges_per_column = 4

    # Convert images to ImageReaders (robust)
    logo_reader = image_to_reader(logo_image) if logo_image else None
    bg_reader = image_to_reader(background_image) if background_image else None
    second_reader = image_to_reader(second_image) if second_image else None
    bottom_reader = image_to_reader(bottom_image) if bottom_image else None
    extra1_reader = image_to_reader(extra_image1) if extra_image1 else None
    extra2_reader = image_to_reader(extra_image2) if extra_image2 else None

    # Precompute draw sizes
    logo_draw_width, logo_draw_height = compute_draw_size(
        logo_image.size, primary_logo_max_width_cm * cm, primary_logo_max_height_cm * cm
    ) if logo_image else (0, 0)

    logo2_draw_width, logo2_draw_height = compute_draw_size(
        second_image.size, second_logo_max_width_cm * cm, second_logo_max_height_cm * cm
    ) if second_image else (0, 0)

    bottom_draw_w, bottom_draw_h = compute_draw_size(
        bottom_image.size, badge_width - 2 * padding, bottom_image_max_height_cm * cm
    ) if bottom_image else (0, 0)

    extra1_draw_w, extra1_draw_h = compute_draw_size(
        extra_image1.size, extra1_max_width_cm * cm, extra1_max_height_cm * cm
    ) if extra_image1 else (0, 0)

    extra2_draw_w, extra2_draw_h = compute_draw_size(
        extra_image2.size, extra2_max_width_cm * cm, extra2_max_height_cm * cm
    ) if extra_image2 else (0, 0)

    # Spacing pts
    gap_name_org = gap_name_org_cm * cm
    gap_before_first_extra = gap_before_first_extra_cm * cm
    gap_between_extras = gap_between_extras_cm * cm

    # Ensure we have rows to process
    if data is None or data.empty:
        # Draw a single page note to avoid "blank" confusion
        c.setFont("Helvetica-Bold", 16)
        c.drawString(2 * cm, A4[1] - 3 * cm, "No data rows to render (DataFrame empty).")
        c.save()
        buffer.seek(0)
        return buffer

    for i, row in data.iterrows():
        col = i % badges_per_row
        row_num = (i // badges_per_row) % badges_per_column

        # New page when current full
        if i % (badges_per_row * badges_per_column) == 0 and i != 0:
            c.showPage()

        x = left_margin + col * (badge_width + padding)
        y = A4[1] - top_margin - (row_num + 1) * (badge_height + padding)

        # Debug draw badge frame
        if debug_draw_boxes:
            c.setStrokeColorRGB(1, 0, 0)
            c.setLineWidth(0.5)
            c.rect(x, y, badge_width, badge_height)

        # Background
        if bg_reader:
            bg_w, bg_h = background_image.size
            bg_aspect = bg_w / bg_h if bg_h else 1.0
            badge_aspect = badge_width / badge_height
            if bg_aspect >= badge_aspect:
                bg_draw_width = badge_width
                bg_draw_height = badge_width / bg_aspect
            else:
                bg_draw_height = badge_height
                bg_draw_width = badge_height * bg_aspect
            bg_x = x + (badge_width - bg_draw_width) / 2
            bg_y = y + (badge_height - bg_draw_height) / 2
            c.drawImage(bg_reader, bg_x, bg_y, width=bg_draw_width, height=bg_draw_height, mask='auto')

        # Badge border (fixed or color-coded)
        if border_style == "Color-coded" and color_column and pd.notna(row.get(color_column)):
            color_hex = color_mapping.get(str(row[color_column]), "#CCCCCC")
        else:
            color_hex = fixed_border_color
        hex_clean = str(color_hex).lstrip("#")
        try:
            r, g, b = tuple(int(hex_clean[i:i+2], 16)/255.0 for i in (0, 2, 4))
        except Exception:
            r, g, b = (0.8, 0.8, 0.8)
        c.setStrokeColorRGB(r, g, b)
        c.setLineWidth(border_thickness)
        c.rect(x, y, badge_width, badge_height)

        # Primary logo (right)
        if logo_reader and logo_draw_width > 0 and logo_draw_height > 0:
            right_logo_x = x + badge_width - logo_draw_width - padding
            if logo_position == "Right Top":
                right_logo_y = y + badge_height - logo_draw_height - padding
            elif logo_position == "Right Bottom":
                right_logo_y = y + padding
            else:
                right_logo_y = y + (badge_height - logo_draw_height) / 2
        else:
            right_logo_x = right_logo_y = None

        # Second logo (left)
        if second_reader and logo2_draw_width > 0 and logo2_draw_height > 0:
            left_logo_x = x + padding
            if second_logo_position == "Left Top":
                left_logo_y = y + badge_height - logo2_draw_height - padding
            elif second_logo_position == "Left Bottom":
                left_logo_y = y + padding
            else:
                left_logo_y = y + (badge_height - logo2_draw_height) / 2
        else:
            left_logo_x = left_logo_y = None

        # Extra images positions
        extra1_px = extra1_py = None
        if extra1_reader and extra1_draw_w > 0 and extra1_draw_h > 0:
            extra1_px, extra1_py = place_by_anchor(
                x, y, badge_width, badge_height, extra1_draw_w, extra1_draw_h,
                extra1_anchor, padding, extra1_offset_x_cm * cm, extra1_offset_y_cm * cm
            )
            if debug_draw_boxes:
                c.setStrokeColorRGB(0, 1, 0)
                c.rect(extra1_px, extra1_py, extra1_draw_w, extra1_draw_h)

        extra2_px = extra2_py = None
        if extra2_reader and extra2_draw_w > 0 and extra2_draw_h > 0:
            extra2_px, extra2_py = place_by_anchor(
                x, y, badge_width, badge_height, extra2_draw_w, extra2_draw_h,
                extra2_anchor, padding, extra2_offset_x_cm * cm, extra2_offset_y_cm * cm
            )
            if debug_draw_boxes:
                c.setStrokeColorRGB(0, 0, 1)
                c.rect(extra2_px, extra2_py, extra2_draw_w, extra2_draw_h)

        # Text block vertical span
        present_extras = [f for f in extra_fields if (f in row and pd.notna(row[f]))]
        if present_extras:
            bottom_offset = gap_before_first_extra + gap_between_extras * (len(present_extras) - 1)
        else:
            bottom_offset = gap_name_org

        if text_vertical == "Top":
            name_y = y + badge_height - 1.2 * cm
        elif text_vertical == "Middle":
            name_y = y + badge_height / 2 + bottom_offset / 2
        else:
            name_y = y + padding + bottom_offset

        text_top_y = name_y
        text_bottom_y = name_y - bottom_offset

        # Build blocked intervals for overlap with all images
        L = x + padding
        R = x + badge_width - padding
        blocked = []

        if logo_reader and right_logo_y is not None and ranges_intersect(text_bottom_y, text_top_y, right_logo_y, right_logo_y + logo_draw_height):
            blocked.append((right_logo_x, right_logo_x + logo_draw_width))
        if second_reader and left_logo_y is not None and ranges_intersect(text_bottom_y, text_top_y, left_logo_y, left_logo_y + logo2_draw_height):
            blocked.append((left_logo_x, left_logo_x + logo2_draw_width))
        if extra1_px is not None and ranges_intersect(text_bottom_y, text_top_y, extra1_py, extra1_py + extra1_draw_h):
            blocked.append((extra1_px, extra1_px + extra1_draw_w))
        if extra2_px is not None and ranges_intersect(text_bottom_y, text_top_y, extra2_py, extra2_py + extra2_draw_h):
            blocked.append((extra2_px, extra2_px + extra2_draw_w))

        free_start, free_end = largest_free_interval(L, R, blocked)
        available_width = max(0.0, free_end - free_start)
        if available_width < 1.0 * cm:
            free_start = L
            available_width = max(1.0 * cm, R - L)

        # Debug: text box outline
        if debug_draw_boxes:
            c.setStrokeColorRGB(1, 0.5, 0)
            c.rect(free_start, text_bottom_y, available_width, text_top_y - text_bottom_y)

        # TEXT — Name
        c.setFillColorRGB(0, 0, 0)
        try:
            c.setFont(font_type + "-Bold", name_font_size)
            name_font_used = font_type + "-Bold"
        except:
            c.setFont(font_type, name_font_size)
            name_font_used = font_type
        name_text = fit_text_to_width(row.get('Name', ''), name_font_used, name_font_size, available_width)
        name_w = pdfmetrics.stringWidth(name_text, name_font_used, name_font_size)
        if text_horizontal == "Center":
            name_x = free_start + (available_width - name_w) / 2
        elif text_horizontal == "Right":
            name_x = free_end - name_w
        else:
            name_x = free_start
        c.drawString(name_x, name_y, name_text)

        # Name-only border (with padding)
        if show_name_border and name_text:
            hex_clean = name_border_color.lstrip("#")
            nr, ng, nb = tuple(int(hex_clean[i:i+2], 16)/255.0 for i in (0, 2, 4))
            c.setStrokeColorRGB(nr, ng, nb)
            c.setLineWidth(name_border_thickness)
            pad_box = name_border_padding_cm * cm
            c.rect(name_x - pad_box, name_y - pad_box, name_w + 2 * pad_box, name_font_size + 2 * pad_box)

        # TEXT — Organisation
        if bold_organisation:
            try:
                c.setFont(font_type + "-Bold", org_font_size)
                org_font_used = font_type + "-Bold"
            except:
                c.setFont(font_type, org_font_size)
                org_font_used = font_type
        else:
            c.setFont(font_type, org_font_size)
            org_font_used = font_type

        org_text = fit_text_to_width(row.get('Organisation', ''), org_font_used, org_font_size, available_width)
        org_w = pdfmetrics.stringWidth(org_text, org_font_used, org_font_size)
        org_y = name_y - gap_name_org
        if text_horizontal == "Center":
            org_x = free_start + (available_width - org_w) / 2
        elif text_horizontal == "Right":
            org_x = free_end - org_w
        else:
            org_x = free_start
        c.drawString(org_x, org_y, org_text)

        # TEXT — Extra fields
        extra_fitted_widths = []  # capture widths for fit-to-content box mode
        offset = gap_before_first_extra
        for f in present_extras:
            c.setFont(font_type, extra_font_size)
            extra_text = fit_text_to_width(str(row[f]), font_type, extra_font_size, available_width)
            extra_w = pdfmetrics.stringWidth(extra_text, font_type, extra_font_size)
            extra_fitted_widths.append(extra_w)
            extra_y = name_y - offset
            if text_horizontal == "Center":
                extra_x = free_start + (available_width - extra_w) / 2
            elif text_horizontal == "Right":
                extra_x = free_end - extra_w
            else:
                extra_x = free_start
            c.drawString(extra_x, extra_y, extra_text)
            offset += gap_between_extras

        # Draw logos and extras
        if logo_reader and right_logo_x is not None:
            c.drawImage(logo_reader, right_logo_x, right_logo_y, width=logo_draw_width, height=logo_draw_height, mask='auto')
        if second_reader and left_logo_x is not None:
            c.drawImage(second_reader, left_logo_x, left_logo_y, width=logo2_draw_width, height=logo2_draw_height, mask='auto')
        if extra1_reader and extra1_px is not None:
            c.drawImage(extra1_reader, extra1_px, extra1_py, width=extra1_draw_w, height=extra1_draw_h, mask='auto')
        if extra2_reader and extra2_px is not None:
            c.drawImage(extra2_reader, extra2_px, extra2_py, width=extra2_draw_w, height=extra2_draw_h, mask='auto')

        # Color circle (optional)
        if show_color_circle and color_column and color_column in row and pd.notna(row[color_column]):
            color_hex2 = color_mapping.get(str(row[color_column]), "#000000")
            hex_clean2 = str(color_hex2).lstrip("#")
            try:
                rr, gg, bb = tuple(int(hex_clean2[i:i+2], 16)/255.0 for i in (0, 2, 4))
            except Exception:
                rr, gg, bb = (0, 0, 0)
            circle_x = x + padding + 0.3 * cm
            circle_y = y + padding + 0.3 * cm
            c.setFillColorRGB(rr, gg, bb)
            c.circle(circle_x, circle_y, 0.3 * cm, fill=1)

        # Bottom strip
        if bottom_reader and bottom_draw_w > 0 and bottom_draw_h > 0:
            bottom_x = x + (badge_width - bottom_draw_w) / 2
            bottom_y = y + padding
            c.drawImage(bottom_reader, bottom_x, bottom_y, width=bottom_draw_w, height=bottom_draw_h, mask='auto')

        # Text block border (adjustable modes, padding, offsets)
        if show_text_border:
            pad_pts = text_box_padding_cm * cm
            fixed_w_pts = text_box_fixed_width_cm * cm

            # Compute max content width using fitted widths
            content_line_widths = [name_w, org_w] + extra_fitted_widths
            max_content_w = max(content_line_widths) if content_line_widths else 0.0

            if text_box_mode == "Full available width":
                box_w = available_width
                box_x = free_start
            elif text_box_mode == "Fit to content":
                box_w = min(max_content_w + 2 * pad_pts, available_width)
                # align with text block: center around the text free interval
                if text_horizontal == "Left":
                    box_x = free_start
                elif text_horizontal == "Center":
                    box_x = free_start + (available_width - box_w) / 2
                else:
                    box_x = free_end - box_w
            else:  # Fixed width
                box_w = min(fixed_w_pts, available_width)
                if text_horizontal == "Left":
                    box_x = free_start
                elif text_horizontal == "Center":
                    box_x = free_start + (available_width - box_w) / 2
                else:
                    box_x = free_end - box_w

            # Apply offsets
            box_x += text_box_offset_x_cm * cm
            box_y = (text_bottom_y - pad_pts) + text_box_offset_y_cm * cm
            box_h = (text_top_y - text_bottom_y) + 2 * pad_pts

            # Clamp box within badge horizontally (optional safety)
            box_x = max(x + padding, min(box_x, x + badge_width - padding - box_w))

            c.setStrokeColorRGB(*text_border_color)
            c.setLineWidth(text_border_thickness)
            c.rect(box_x, box_y, box_w, box_h)

    c.save()
    buffer.seek(0)
    return buffer

# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Name Badge Generator", layout="wide")
st.title("📛 Name Badge Generator – Full Layout (No Overlap & Adjustable Text Box)")

st.markdown(
    """
This tool turns a simple spreadsheet of attendees into print-ready name badges (PDF).

**How to use it:**
1. **Upload an Excel file** with a column for each person's **Name** and **Organisation** (the header names
   can be anywhere near the top of the sheet, in any case, and "Organization" spelled the American way works too).
2. **Optionally upload a logo and/or background image** to brand the badges.
3. Use the options in the sidebar on the left to tweak fonts, colors, borders and layout — changes update live.
4. Click **Download PDF** when you're happy with the preview, then print on your badge stock.

No Excel to hand? Tick **"Use sample data"** in the sidebar to try it out with example names first.
"""
)

# Add new entries to the top of this list as the app changes.
WHATS_NEW = [
    ("2026-10-01", "Added this plain-language intro and usage guide."),
    ("2026-10-01", "Excel upload is now more forgiving: the Name/Organisation header row can be anywhere "
                   "near the top, in any case, with extra spaces, and 'Organization' (US spelling) is accepted."),
]
with st.expander("🆕 What's new", expanded=False):
    for date, note in WHATS_NEW:
        st.markdown(f"- **{date}** — {note}")

# Large, clear upload prompts
st.markdown("### 📂 Upload your files")
uploaded_excel = st.file_uploader("**Upload Excel with at least 'Name' and 'Organisation' columns**", type=["xlsx"])
uploaded_logo = st.file_uploader("**Primary Logo** (PNG/JPG)", type=["png", "jpg", "jpeg"])
uploaded_logo2 = st.file_uploader("Optional: **Second Logo** (PNG/JPG)", type=["png", "jpg", "jpeg"])
uploaded_background = st.file_uploader("Optional: **Background Image** (PNG/JPG)", type=["png", "jpg", "jpeg"])
uploaded_bottom_image = st.file_uploader("Optional: **Bottom Strip Image** (PNG/JPG)", type=["png", "jpg", "jpeg"])
uploaded_extra1 = st.file_uploader("Optional: **Extra Image 1** (PNG/JPG)", type=["png", "jpg", "jpeg"])
uploaded_extra2 = st.file_uploader("Optional: **Extra Image 2** (PNG/JPG)", type=["png", "jpg", "jpeg"])

# Preview & debug
st.sidebar.header("🧪 Preview & Debug")
use_sample_data = st.sidebar.checkbox("Use sample data (8 badges) if no Excel", value=True)
debug_draw_boxes = st.sidebar.checkbox("Debug: draw layout guides", value=False)

# Fonts
st.sidebar.header("🖋️ Fonts")
font_type = st.sidebar.selectbox("Font", ["Helvetica", "Times-Roman", "Courier"])
name_font_size = st.sidebar.slider("Name font size", 8, 28, 14)
org_font_size = st.sidebar.slider("Organisation font size", 6, 22, 11)
extra_font_size = st.sidebar.slider("Extra fields font size", 6, 20, 10)
bold_organisation = st.sidebar.checkbox("Bold Organisation", value=False)

# Badge border
st.sidebar.header("🎨 Badge Border")
border_style = st.sidebar.radio("Style", ["Fixed", "Color-coded"])
fixed_border_color = st.sidebar.color_picker("Fixed border color", "#CCCCCC")
border_thickness = st.sidebar.slider("Border thickness", 0.5, 5.0, 1.0)

# Logos
st.sidebar.header("🖼️ Primary Logo (Right)")
logo_position = st.sidebar.selectbox("Primary logo position", ["Right Top", "Right Middle", "Right Bottom"])
primary_logo_max_width_cm = st.sidebar.slider("Max width (cm)", 1.0, 6.0, 3.25)
primary_logo_max_height_cm = st.sidebar.slider("Max height (cm)", 1.0, 6.0, 5.0)

st.sidebar.header("🖼️ Second Logo (Left)")
second_logo_position = st.sidebar.selectbox("Second logo position", ["Left Top", "Left Middle", "Left Bottom"])
second_logo_max_width_cm = st.sidebar.slider("Max width (cm) (2nd)", 1.0, 6.0, 3.0)
second_logo_max_height_cm = st.sidebar.slider("Max height (cm) (2nd)", 1.0, 6.0, 4.5)

# Background & bottom strip
st.sidebar.header("🌗 Background")
background_opacity = st.sidebar.slider("Background opacity", 0.0, 1.0, 1.0)

st.sidebar.header("⬇️ Bottom Strip")
bottom_image_max_height_cm = st.sidebar.slider("Bottom strip max height (cm)", 0.5, 3.0, 1.2)

# Text layout
st.sidebar.header("🅰️ Text Layout")
text_horizontal = st.sidebar.selectbox("Horizontal alignment", ["Left", "Center", "Right"])
text_vertical = st.sidebar.selectbox("Vertical position", ["Top", "Middle", "Bottom"])

st.sidebar.header("↕️ Line Spacing (cm)")
gap_name_org_cm = st.sidebar.slider("Gap: Name → Organisation", 0.3, 1.5, 0.6, step=0.05)
gap_before_first_extra_cm = st.sidebar.slider("Gap: Organisation → First Extra", 0.6, 2.0, 1.2, step=0.05)
gap_between_extras_cm = st.sidebar.slider("Gap: Between Extras", 0.3, 1.5, 0.6, step=0.05)

# Color circle
st.sidebar.header("🔵 Color Circle")
show_color_circle = st.sidebar.checkbox("Show color circle", value=True)

# Text borders
st.sidebar.header("🧩 Text Borders")
show_text_border = st.sidebar.checkbox("Show border around entire text block", value=False)
text_border_color_hex = st.sidebar.color_picker("Text block border color", "#000000")
text_border_thickness = st.sidebar.slider("Text block border thickness", 0.5, 5.0, 1.0)

# Box width / padding / offsets
st.sidebar.subheader("Text Box Border – Size & Padding")
text_box_mode = st.sidebar.selectbox("Text box width mode", ["Full available width", "Fit to content", "Fixed width"])
text_box_fixed_width_cm = st.sidebar.slider("Fixed width (cm)", 1.0, 9.0, 6.0, step=0.1)
text_box_padding_cm = st.sidebar.slider("Text box padding (cm)", 0.1, 1.0, 0.25, step=0.05)
text_box_offset_x_cm = st.sidebar.slider("Text box X offset (cm)", -2.0, 2.0, 0.0, step=0.1)
text_box_offset_y_cm = st.sidebar.slider("Text box Y offset (cm)", -1.0, 1.0, 0.0, step=0.1)

# Name border
st.sidebar.header("👤 Name Border")
show_name_border = st.sidebar.checkbox("Show border around Name only", value=False)
name_border_color = st.sidebar.color_picker("Name border color", "#000000")
name_border_thickness = st.sidebar.slider("Name border thickness", 0.5, 5.0, 1.0)
st.sidebar.subheader("Name Border – Padding")
name_border_padding_cm = st.sidebar.slider("Padding (cm) around name", 0.1, 1.0, 0.25, step=0.05)

# Extra images controls
st.sidebar.header("🖼️ Extra Image 1")
extra1_anchor = st.sidebar.selectbox("Anchor (Img 1)", ["Top Left", "Top Right", "Bottom Left", "Bottom Right", "Center"])
extra1_max_width_cm = st.sidebar.slider("Max width (cm) – Img 1", 0.5, 6.0, 2.5)
extra1_max_height_cm = st.sidebar.slider("Max height (cm) – Img 1", 0.5, 6.0, 2.5)
extra1_offset_x_cm = st.sidebar.slider("Offset X (cm) – Img 1", -3.0, 3.0, 0.0, step=0.1)
extra1_offset_y_cm = st.sidebar.slider("Offset Y (cm) – Img 1", -3.0, 3.0, 0.0, step=0.1)

st.sidebar.header("🖼️ Extra Image 2")
extra2_anchor = st.sidebar.selectbox("Anchor (Img 2)", ["Top Left", "Top Right", "Bottom Left", "Bottom Right", "Center"])
extra2_max_width_cm = st.sidebar.slider("Max width (cm) – Img 2", 0.5, 6.0, 2.5)
extra2_max_height_cm = st.sidebar.slider("Max height (cm) – Img 2", 0.5, 6.0, 2.5)
extra2_offset_x_cm = st.sidebar.slider("Offset X (cm) – Img 2", -3.0, 3.0, 0.0, step=0.1)
extra2_offset_y_cm = st.sidebar.slider("Offset Y (cm) – Img 2", -3.0, 3.0, 0.0, step=0.1)

# ======== Data handling ========

def open_img(upl):
    try:
        return Image.open(upl) if upl else None
    except Exception:
        return None

df = None
if uploaded_excel is not None:
    try:
        df, load_err = load_badge_excel(uploaded_excel)
        if load_err:
            st.error(load_err)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")

if (df is None or df.empty) and use_sample_data:
    df = pd.DataFrame({
        "Name": [f"Sample Name {i+1}" for i in range(8)],
        "Organisation": ["WSP"] * 8,
        "Role": ["Delegate"] * 8,
        "Track": ["A", "B", "C", "A", "B", "C", "A", "B"],
        "Zone": ["North", "South", "East", "West", "North", "South", "East", "West"]
    })
    st.info("Using sample data (8 rows). Upload an Excel to replace this.")

if df is not None and not df.empty:
    if "Name" not in df.columns or "Organisation" not in df.columns:
        st.error("Excel must contain columns 'Name' and 'Organisation'.")
        df = None

# Extra field selection and color mapping
selected_fields = []
color_column = None
color_mapping = {}
if df is not None:
    available_fields = [c for c in df.columns if c not in ["Name", "Organisation"]]
    selected_fields = st.multiselect("Select up to 4 extra fields to show", available_fields, max_selections=4)
    if available_fields:
        color_column = st.selectbox("Select a column for color coding (badge border & circle)", [None] + available_fields)
        if color_column:
            st.markdown("#### Assign colors per category")
            for val in pd.Series(df[color_column]).dropna().unique():
                color_mapping[str(val)] = st.color_picker(f"Color for '{val}'", "#000000", key=f"cc_{val}")

# ======== Images ========

logo_image = open_img(uploaded_logo)
logo2_image = open_img(uploaded_logo2)
bg_image = open_img(uploaded_background)
bottom_image = open_img(uploaded_bottom_image)
extra1_image = open_img(uploaded_extra1)
extra2_image = open_img(uploaded_extra2)

if bg_image:
    bg_image = apply_opacity(bg_image, background_opacity)

text_border_color_rgb = hex_to_rgb01(text_border_color_hex)

# ======== Generate ========

generate = st.button("🔄 Generate PDF Badges", type="primary")

if generate:
    if df is None or df.empty:
        st.error("No rows to render. Upload a valid Excel or enable sample data.")
    else:
        pdf_bytes = generate_badges_pdf(
            data=df,
            logo_image=logo_image,
            background_image=bg_image,
            font_type=font_type,
            name_font_size=name_font_size,
            org_font_size=org_font_size,
            extra_font_size=extra_font_size,
            extra_fields=selected_fields,
            bold_organisation=bold_organisation,
            color_column=color_column,
            color_mapping=color_mapping,
            border_style=border_style,
            fixed_border_color=fixed_border_color,
            border_thickness=border_thickness,
            logo_position=logo_position,
            second_image=logo2_image,
            second_logo_position=second_logo_position,
            bottom_image=bottom_image,
            bottom_image_max_height_cm=bottom_image_max_height_cm,
            primary_logo_max_width_cm=primary_logo_max_width_cm,
            primary_logo_max_height_cm=primary_logo_max_height_cm,
            second_logo_max_width_cm=second_logo_max_width_cm,
            second_logo_max_height_cm=second_logo_max_height_cm,
            extra_image1=extra1_image,
            extra1_anchor=extra1_anchor,
            extra1_max_width_cm=extra1_max_width_cm,
            extra1_max_height_cm=extra1_max_height_cm,
            extra1_offset_x_cm=extra1_offset_x_cm,
            extra1_offset_y_cm=extra1_offset_y_cm,
            extra_image2=extra2_image,
            extra2_anchor=extra2_anchor,
            extra2_max_width_cm=extra2_max_width_cm,
            extra2_max_height_cm=extra2_max_height_cm,
            extra2_offset_x_cm=extra2_offset_x_cm,
            extra2_offset_y_cm=extra2_offset_y_cm,
            show_color_circle=show_color_circle,
            text_horizontal=text_horizontal,
            text_vertical=text_vertical,
            gap_name_org_cm=gap_name_org_cm,
            gap_before_first_extra_cm=gap_before_first_extra_cm,
            gap_between_extras_cm=gap_between_extras_cm,
            show_text_border=show_text_border,
            text_border_color=text_border_color_rgb,   # (r,g,b) 0..1
            text_border_thickness=text_border_thickness,
            show_name_border=show_name_border,
            name_border_color=name_border_color,       # hex string
            name_border_thickness=name_border_thickness,
            text_box_mode=text_box_mode,
            text_box_fixed_width_cm=text_box_fixed_width_cm,
            text_box_padding_cm=text_box_padding_cm,
            text_box_offset_x_cm=text_box_offset_x_cm,
            text_box_offset_y_cm=text_box_offset_y_cm,
            name_border_padding_cm=name_border_padding_cm,
            debug_draw_boxes=debug_draw_boxes
        )
        st.download_button(
            "📥 Download PDF",
            pdf_bytes,
            file_name="name_badges.pdf",
            mime="application/pdf"
        )
        st.success("PDF generated. If something looks off, enable 'Debug: draw layout guides' and tune padding/offsets.")
