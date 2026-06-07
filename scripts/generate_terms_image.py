#-----------------------------------------------------------------------------------------------------------------------
# Program: Year Planner Generator
# Version: 1.1
# Author:  Rohin Gosling
#
# Description:
#
#   Generate a PNG image of the Terms and Definitions page.
#
#   Creates a visual representation of page 245 (a Terms and Definitions page) matching the Year Planner document
#   styling.
#
# Usage:
#
#   python scripts/generate_terms_image.py
#-----------------------------------------------------------------------------------------------------------------------

from PIL import Image, ImageDraw, ImageFont
import os

# Global constants.

DPI = 300                   # DPI for high-quality output.

PAGE_WIDTH_CM  = 21.0       # A4 width in cm.
PAGE_HEIGHT_CM = 29.7       # A4 height in cm.


#-----------------------------------------------------------------------------------------------------------------------
# Function: cm_to_px
#
# Description:
#
#   Convert centimeters to pixels at 300 DPI.
#
# Arguments:
#
#   cm : Measurement in centimeters.
#
# Returns:
#
#   Pixel count as an integer.
#-----------------------------------------------------------------------------------------------------------------------

def cm_to_px(cm: float) -> int:

    # Convert centimeters to pixels at 300 DPI.

    inches = cm / 2.54

    # Return data to caller.

    return int(inches * DPI)


# Page dimensions in pixels.

PAGE_WIDTH_PX  = cm_to_px(PAGE_WIDTH_CM)
PAGE_HEIGHT_PX = cm_to_px(PAGE_HEIGHT_CM)

# Margins (from config).

MARGIN_TOP_CM    = 0.6
MARGIN_BOTTOM_CM = 1.2
MARGIN_LEFT_CM   = 0.6
MARGIN_RIGHT_CM  = 0.6
GUTTER_CM        = 1.5

# For a recto (right/odd) page, gutter is on the left.

LEFT_MARGIN_CM  = MARGIN_LEFT_CM + GUTTER_CM  # 2.1 cm
RIGHT_MARGIN_CM = MARGIN_RIGHT_CM              # 0.6 cm

# Content area.

CONTENT_LEFT_PX   = cm_to_px(LEFT_MARGIN_CM)
CONTENT_RIGHT_PX  = PAGE_WIDTH_PX - cm_to_px(RIGHT_MARGIN_CM)
CONTENT_TOP_PX    = cm_to_px(MARGIN_TOP_CM)
CONTENT_BOTTOM_PX = PAGE_HEIGHT_PX - cm_to_px(MARGIN_BOTTOM_CM)

CONTENT_WIDTH_PX  = CONTENT_RIGHT_PX - CONTENT_LEFT_PX
CONTENT_HEIGHT_PX = CONTENT_BOTTOM_PX - CONTENT_TOP_PX

# Table configuration.

ROW_COUNT          = 16   # Content rows.
TERM_WIDTH_PERCENT = 25

# Row heights in points.

TITLE_ROW_HEIGHT_PT  = 14.2
HEADER_ROW_HEIGHT_PT = 14.2


#-----------------------------------------------------------------------------------------------------------------------
# Function: pt_to_px
#
# Description:
#
#   Convert points to pixels at 300 DPI.
#
# Arguments:
#
#   pt : Measurement in points.
#
# Returns:
#
#   Pixel count as an integer.
#-----------------------------------------------------------------------------------------------------------------------

def pt_to_px(pt: float) -> int:

    # Convert points to pixels at 300 DPI.

    inches = pt / 72.0

    # Return data to caller.

    return int(inches * DPI)


TITLE_ROW_HEIGHT_PX  = pt_to_px(TITLE_ROW_HEIGHT_PT)
HEADER_ROW_HEIGHT_PX = pt_to_px(HEADER_ROW_HEIGHT_PT)

# Safety margin and minimized paragraph height.

SAFETY_MARGIN_PX  = pt_to_px(2)   # 40 twips = 2pt
MIN_PARA_HEIGHT_PX = pt_to_px(1)  # 20 twips = 1pt

# Calculate content row height to fill page.

TOTAL_FIXED_HEIGHT_PX      = TITLE_ROW_HEIGHT_PX + HEADER_ROW_HEIGHT_PX + SAFETY_MARGIN_PX + MIN_PARA_HEIGHT_PX
AVAILABLE_FOR_CONTENT_PX   = CONTENT_HEIGHT_PX - TOTAL_FIXED_HEIGHT_PX
CONTENT_ROW_HEIGHT_PX      = AVAILABLE_FOR_CONTENT_PX // ROW_COUNT


#-----------------------------------------------------------------------------------------------------------------------
# Function: grayscale_to_rgb
#
# Description:
#
#   Convert grayscale percentage to RGB tuple.
#
# Arguments:
#
#   percent : Grayscale percentage (0 = white, 100 = black).
#
# Returns:
#
#   RGB tuple as (int, int, int).
#-----------------------------------------------------------------------------------------------------------------------

def grayscale_to_rgb(percent: int) -> tuple:

    # Convert grayscale percentage to RGB tuple.

    value = int(255 * (100 - percent) / 100)

    # Return data to caller.

    return (value, value, value)


# Color definitions from config.

TITLE_BG_COLOR   = grayscale_to_rgb(75)    # 75% gray background.
TITLE_TEXT_COLOR = grayscale_to_rgb(0)     # White text (0%).
HEADER_BG_COLOR  = grayscale_to_rgb(25)    # 25% gray background.
HEADER_TEXT_COLOR = grayscale_to_rgb(100)  # Black text (100%).
BORDER_COLOR     = grayscale_to_rgb(75)    # 75% gray borders.
PAGE_BG_COLOR    = (255, 255, 255)         # White background.

# Border thickness.

BORDER_THICKNESS_PT = 0.75
BORDER_THICKNESS_PX = max(1, pt_to_px(BORDER_THICKNESS_PT))

# Font settings.

TITLE_FONT_SIZE_PT  = 10
HEADER_FONT_SIZE_PT = 10


#-----------------------------------------------------------------------------------------------------------------------
# Function: pt_to_font_size
#
# Description:
#
#   Convert point size to PIL font size (approximately).
#
# Arguments:
#
#   pt : Font size in points.
#
# Returns:
#
#   Approximate PIL font size as an integer.
#-----------------------------------------------------------------------------------------------------------------------

def pt_to_font_size(pt: float) -> int:

    # Convert point size to PIL font size (approximately).

    # PIL font sizes scale differently; approximate.

    # Return data to caller.

    return int(pt * DPI / 72)


#-----------------------------------------------------------------------------------------------------------------------
# Function: get_font
#
# Description:
#
#   Get a font at the specified size.
#
#   Attempts to load Arial (or DejaVu Sans on Linux) and falls back to the PIL default if neither is available.
#
# Arguments:
#
#   size_pt : Desired font size in points.
#   bold    : If True, load the bold variant.
#
# Returns:
#
#   An ImageFont.FreeTypeFont instance, or the PIL default font on failure.
#-----------------------------------------------------------------------------------------------------------------------

def get_font(size_pt: float, bold: bool = False) -> ImageFont.FreeTypeFont:

    # Get a font at the specified size.

    size = pt_to_font_size(size_pt)

    # Try to use Arial or a similar font.

    font_paths = [
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]

    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                return ImageFont.truetype(font_path, size)
            except Exception:
                pass

    # Fallback to default font.

    # Return data to caller.

    return ImageFont.load_default()


#-----------------------------------------------------------------------------------------------------------------------
# Function: draw_terms_definitions_page
#
# Description:
#
#   Generate the Terms and Definitions page image.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   A PIL Image object containing the rendered page.
#-----------------------------------------------------------------------------------------------------------------------

def draw_terms_definitions_page():

    # Generate the Terms and Definitions page image.

    # Create image.

    img = Image.new('RGB', (PAGE_WIDTH_PX, PAGE_HEIGHT_PX), PAGE_BG_COLOR)
    draw = ImageDraw.Draw(img)

    # Column widths.

    term_col_width_px = int(CONTENT_WIDTH_PX * TERM_WIDTH_PERCENT / 100)
    def_col_width_px = CONTENT_WIDTH_PX - term_col_width_px

    # Starting position.

    table_left = CONTENT_LEFT_PX
    table_top = CONTENT_TOP_PX + MIN_PARA_HEIGHT_PX  # Account for minimized paragraph.
    current_y = table_top

    # Get fonts.

    title_font = get_font(TITLE_FONT_SIZE_PT, bold=True)
    header_font = get_font(HEADER_FONT_SIZE_PT, bold=True)

    # === TITLE ROW ===

    title_rect = (table_left, current_y, table_left + CONTENT_WIDTH_PX, current_y + TITLE_ROW_HEIGHT_PX)
    draw.rectangle(title_rect, fill=TITLE_BG_COLOR)

    # Title text.

    title_text = "Terms and Definitions"
    bbox = draw.textbbox((0, 0), title_text, font=title_font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    text_x = table_left + 10  # Left padding.
    text_y = current_y + (TITLE_ROW_HEIGHT_PX - text_height) // 2
    draw.text((text_x, text_y), title_text, fill=TITLE_TEXT_COLOR, font=title_font)

    current_y += TITLE_ROW_HEIGHT_PX

    # === HEADER ROW ===

    # Term column header.

    term_header_rect = (table_left, current_y, table_left + term_col_width_px, current_y + HEADER_ROW_HEIGHT_PX)
    draw.rectangle(term_header_rect, fill=HEADER_BG_COLOR)

    term_header_text = "Term / Abbreviation"
    bbox = draw.textbbox((0, 0), term_header_text, font=header_font)
    text_height = bbox[3] - bbox[1]
    text_x = table_left + 10
    text_y = current_y + (HEADER_ROW_HEIGHT_PX - text_height) // 2
    draw.text((text_x, text_y), term_header_text, fill=HEADER_TEXT_COLOR, font=header_font)

    # Definition column header.

    def_header_rect = (table_left + term_col_width_px, current_y,
                       table_left + CONTENT_WIDTH_PX, current_y + HEADER_ROW_HEIGHT_PX)
    draw.rectangle(def_header_rect, fill=HEADER_BG_COLOR)

    def_header_text = "Definition"
    bbox = draw.textbbox((0, 0), def_header_text, font=header_font)
    text_height = bbox[3] - bbox[1]
    text_x = table_left + term_col_width_px + 10
    text_y = current_y + (HEADER_ROW_HEIGHT_PX - text_height) // 2
    draw.text((text_x, text_y), def_header_text, fill=HEADER_TEXT_COLOR, font=header_font)

    current_y += HEADER_ROW_HEIGHT_PX

    # === CONTENT ROWS ===

    for row_idx in range(ROW_COUNT):
        row_top = current_y
        row_bottom = current_y + CONTENT_ROW_HEIGHT_PX

        # Draw horizontal line at top of row.

        draw.line([(table_left, row_top), (table_left + CONTENT_WIDTH_PX, row_top)],
                  fill=BORDER_COLOR, width=BORDER_THICKNESS_PX)

        # Draw vertical divider.

        draw.line([(table_left + term_col_width_px, row_top),
                   (table_left + term_col_width_px, row_bottom)],
                  fill=BORDER_COLOR, width=BORDER_THICKNESS_PX)

        current_y = row_bottom

    # === DRAW TABLE BORDERS ===

    table_bottom = table_top + TITLE_ROW_HEIGHT_PX + HEADER_ROW_HEIGHT_PX + (ROW_COUNT * CONTENT_ROW_HEIGHT_PX)

    # Outer border.

    draw.rectangle(
        (table_left, table_top, table_left + CONTENT_WIDTH_PX, table_bottom),
        outline=BORDER_COLOR, width=BORDER_THICKNESS_PX
    )

    # Horizontal line after title row.

    draw.line(
        [(table_left, table_top + TITLE_ROW_HEIGHT_PX),
         (table_left + CONTENT_WIDTH_PX, table_top + TITLE_ROW_HEIGHT_PX)],
        fill=BORDER_COLOR, width=BORDER_THICKNESS_PX
    )

    # Horizontal line after header row.

    draw.line(
        [(table_left, table_top + TITLE_ROW_HEIGHT_PX + HEADER_ROW_HEIGHT_PX),
         (table_left + CONTENT_WIDTH_PX, table_top + TITLE_ROW_HEIGHT_PX + HEADER_ROW_HEIGHT_PX)],
        fill=BORDER_COLOR, width=BORDER_THICKNESS_PX
    )

    # Vertical divider through title row.

    draw.line(
        [(table_left + term_col_width_px, table_top + TITLE_ROW_HEIGHT_PX),
         (table_left + term_col_width_px, table_bottom)],
        fill=BORDER_COLOR, width=BORDER_THICKNESS_PX
    )

    # Bottom line.

    draw.line(
        [(table_left, table_bottom), (table_left + CONTENT_WIDTH_PX, table_bottom)],
        fill=BORDER_COLOR, width=BORDER_THICKNESS_PX
    )

    # === PAGE NUMBER ===

    page_num_font = get_font(10, bold=False)
    page_num_text = "245"
    bbox = draw.textbbox((0, 0), page_num_text, font=page_num_font)
    text_width = bbox[2] - bbox[0]

    # Recto page: page number at bottom right.

    page_num_y = PAGE_HEIGHT_PX - cm_to_px(0.55) - (bbox[3] - bbox[1])
    page_num_x = CONTENT_RIGHT_PX - text_width
    draw.text((page_num_x, page_num_y), page_num_text, fill=(0, 0, 0), font=page_num_font)

    # Return data to caller.

    return img


#-----------------------------------------------------------------------------------------------------------------------
# Function: main
#
# Description:
#
#   Generate and save the Terms and Definitions page image.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def main():

    # Generate and save the Terms and Definitions page image.

    print("Generating Terms and Definitions page image...")

    # Generate image.

    img = draw_terms_definitions_page()

    # Ensure output directory exists.

    output_path = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                               "assets", "images", "terms_and_definitions.png")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Save image.

    img.save(output_path, "PNG", dpi=(DPI, DPI))
    print(f"Saved: {output_path}")
    print(f"Dimensions: {img.width} x {img.height} pixels ({DPI} DPI)")


if __name__ == "__main__":
    main()
