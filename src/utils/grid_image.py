#-----------------------------------------------------------------------------------------------------------------------
# Module:  grid_image.py
# Project: Year Planner Generator
# Version: 1.1
# Author:  Rohin Gosling
#
# Description:
#
#   Grid image generation for graph paper. Creates grid images using Pillow for insertion into Word documents.
#-----------------------------------------------------------------------------------------------------------------------

from PIL import Image, ImageDraw


#-----------------------------------------------------------------------------------------------------------------------
# Function: generate_grid_image
#
# Description:
#
#   Generate a graph paper grid image and save to file.
#
#   Creates a white background image with evenly spaced grid lines and a distinct outer border.
#
# Arguments:
#
#   width_px            : Image width in pixels.
#   height_px           : Image height in pixels.
#   columns             : Number of vertical grid divisions.
#   rows                : Number of horizontal grid divisions.
#   grid_color_percent  : Interior grid line color (0=white, 100=black).
#   border_color_percent : Outer border color (0=white, 100=black).
#   output_path         : File path to save the PNG image.
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def generate_grid_image (
    width_px: int,
    height_px: int,
    columns: int,
    rows: int,
    grid_color_percent: int,
    border_color_percent: int,
    output_path: str
) -> None:

    # Generate a graph paper grid image and save to file.

    # Convert grayscale percentages to RGB values (0% = white, 100% = black).

    grid_gray   = int ( 255 * ( 1 - grid_color_percent / 100 ) )
    border_gray = int ( 255 * ( 1 - border_color_percent / 100 ) )

    grid_color   = ( grid_gray, grid_gray, grid_gray )
    border_color = ( border_gray, border_gray, border_gray )

    # Create white background image.

    image = Image.new ( 'RGB', ( width_px, height_px ), 'white' )
    draw  = ImageDraw.Draw ( image )

    # Calculate cell dimensions.

    cell_width  = width_px / columns
    cell_height = height_px / rows

    # Draw interior vertical grid lines (1px).

    for i in range ( 1, columns ):
        x = int ( i * cell_width )
        draw.line ( [ ( x, 0 ), ( x, height_px - 1 ) ], fill = grid_color, width = 1 )

    # Draw interior horizontal grid lines (1px).

    for i in range ( 1, rows ):
        y = int ( i * cell_height )
        draw.line ( [ ( 0, y ), ( width_px - 1, y ) ], fill = grid_color, width = 1 )

    # Draw outer border (2px for definition).

    # Top border.

    draw.line ( [ ( 0, 0 ), ( width_px - 1, 0 ) ], fill = border_color, width = 2 )
    draw.line ( [ ( 0, 1 ), ( width_px - 1, 1 ) ], fill = border_color, width = 1 )

    # Bottom border.

    draw.line ( [ ( 0, height_px - 1 ), ( width_px - 1, height_px - 1 ) ],
                fill = border_color, width = 2 )
    draw.line ( [ ( 0, height_px - 2 ), ( width_px - 1, height_px - 2 ) ],
                fill = border_color, width = 1 )

    # Left border.

    draw.line ( [ ( 0, 0 ), ( 0, height_px - 1 ) ], fill = border_color, width = 2 )
    draw.line ( [ ( 1, 0 ), ( 1, height_px - 1 ) ], fill = border_color, width = 1 )

    # Right border.

    draw.line ( [ ( width_px - 1, 0 ), ( width_px - 1, height_px - 1 ) ],
                fill = border_color, width = 2 )
    draw.line ( [ ( width_px - 2, 0 ), ( width_px - 2, height_px - 1 ) ],
                fill = border_color, width = 1 )

    # Save as PNG.

    image.save ( output_path, 'PNG' )
