#-----------------------------------------------------------------------------------------------------------------------
# Program: Year Planner Generator
# Version: 1.1
# Author:  Rohin Gosling
#
# Description:
#
#   Convert a single-page PDF to a high-resolution PNG image.
#
# Usage:
#
#   python scripts/pdf_to_png.py input.pdf output.png [--dpi 300]
#-----------------------------------------------------------------------------------------------------------------------

import argparse
import sys
from pathlib import Path

import fitz  # PyMuPDF


#-----------------------------------------------------------------------------------------------------------------------
# Function: pdf_to_png
#
# Description:
#
#   Convert a single-page PDF to a high-resolution PNG.
#
# Arguments:
#
#   pdf_path    : Path to the input PDF file.
#   output_path : Path for the output PNG file.
#   dpi         : Resolution in dots per inch (default: 300).
#
# Returns:
#
#   None.
#
# Raises:
#
#   FileNotFoundError : If the PDF file doesn't exist.
#   ValueError        : If the PDF has no pages.
#-----------------------------------------------------------------------------------------------------------------------

def pdf_to_png ( pdf_path: str, output_path: str, dpi: int = 300 ) -> None:

    # Convert a single-page PDF to a high-resolution PNG.

    pdf_file = Path ( pdf_path )
    if not pdf_file.exists ():
        raise FileNotFoundError ( f"PDF file not found: {pdf_path}" )

    doc = fitz.open ( pdf_path )

    if len ( doc ) == 0:
        doc.close ()
        raise ValueError ( f"PDF has no pages: {pdf_path}" )

    # Get the first page.

    page = doc [ 0 ]

    # Scale factor for desired DPI (default PDF resolution is 72 DPI).

    scale  = dpi / 72
    matrix = fitz.Matrix ( scale, scale )

    # Render page to pixmap.

    pix = page.get_pixmap ( matrix = matrix )

    # Save as PNG.

    output_file = Path ( output_path )
    output_file.parent.mkdir ( parents = True, exist_ok = True )
    pix.save ( str ( output_file ) )

    doc.close ()

    print ( f"Converted: {pdf_path}" )
    print ( f"Output:    {output_path}" )
    print ( f"Size:      {pix.width} x {pix.height} pixels ({dpi} DPI)" )


#-----------------------------------------------------------------------------------------------------------------------
# Function: main
#
# Description:
#
#   Command-line entry point.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def main () -> None:

    # Command-line entry point.

    parser = argparse.ArgumentParser (
        description = "Convert a single-page PDF to a high-resolution PNG image."
    )
    parser.add_argument ( "input", help = "Path to the input PDF file" )
    parser.add_argument ( "output", help = "Path for the output PNG file" )
    parser.add_argument (
        "--dpi",
        type    = int,
        default = 300,
        help    = "Resolution in DPI (default: 300)"
    )

    args = parser.parse_args ()

    try:
        pdf_to_png ( args.input, args.output, args.dpi )
    except ( FileNotFoundError, ValueError ) as e:
        print ( f"Error: {e}", file = sys.stderr )
        sys.exit ( 1 )


if __name__ == "__main__":
    main ()
