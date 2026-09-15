#-----------------------------------------------------------------------------------------------------------------------
# Program: Year Planner Generator
# Version: 1.2
# Author:  Rohin Gosling
#
# Description:
#
#   Generates a configurable Year Planner Microsoft Word document from a bundled YAML configuration file. DOCX is
#   always produced; a PDF is an optional opt-in (--pdf) available only on Windows with Microsoft Word installed.
#
# Usage:
#
#   python src/main.py --year 2027
#   python src/main.py --year 2027 --pdf
#   python src/main.py --year 2027 --filename my-planner
#-----------------------------------------------------------------------------------------------------------------------

import argparse
import re
import shutil
import sys
from datetime import date
from pathlib  import Path

# Add project root to path for imports.

project_root = Path ( __file__ ).parent.parent
sys.path.insert ( 0, str ( project_root ) )

from src.config                     import load_config, Config
from src.document                   import create_document, add_page_break, add_section_break, add_numbered_section_break, add_non_numbered_section_break
from src.paths                      import resource_path, resolve_output_paths, word_available
from src.sections.cover             import generate_cover_page
from src.sections.instructions      import generate_instructions_page
from src.sections.calendar          import generate_calendar_section
from src.sections.toc               import generate_toc
from src.sections.goals             import generate_goals_page
from src.sections.week_planner      import generate_week_planner
from src.sections.backlog           import generate_backlog
from src.sections.graph_paper       import generate_graph_paper
from src.sections.monthly           import generate_monthly_sections
from src.sections.terms_definitions import generate_terms_definitions
from src.sections.rear_cover        import generate_rear_cover


#-----------------------------------------------------------------------------------------------------------------------
# Function: _changelog_version
#
# Description:
#
#   Read the highest version from docs/CHANGELOG.md. Scans for "## [X.Y]" headings (skipping "[Unreleased]") and
#   returns the first match, which is the highest version. Returns the fallback when the file is missing or contains
#   no versioned headings (e.g. a frozen .exe, where docs/ is not bundled). build.py stamps the fallback at build time.
#
# Arguments:
#
#   fallback : Version string to return when no version is found.
#
# Returns:
#
#   The version string from the changelog, or the fallback value.
#-----------------------------------------------------------------------------------------------------------------------

def _changelog_version ( fallback: str = "1.2" ) -> str:

    # Read the highest version from docs/CHANGELOG.md.

    changelog_path = resource_path ( "docs", "CHANGELOG.md" )

    try:
        with open ( changelog_path, "r", encoding = "utf-8" ) as changelog_file:
            for line in changelog_file:
                match = re.match ( r"^##\s+\[(\d+\.\d+)\]", line )
                if match:

                    # Return the first versioned heading, which is the highest version.

                    return match.group ( 1 )
    except OSError:
        pass

    # Return fallback when the file is missing or contains no versioned headings.

    return fallback

# Program metadata. BUILD_DATE is stamped by scripts/build.py at build time; VERSION's fallback is stamped likewise.

VERSION      = _changelog_version ( "1.2" )
BUILD_DATE   = "2026-06-07"
TITLE        = "Year Planner Generator"
AUTHOR       = "Rohin Gosling"
BANNER_WIDTH = 60


#-----------------------------------------------------------------------------------------------------------------------
# Function: _banner_text
#
# Description:
#
#   Build the program banner as a box-drawing string showing program, version, build date, and author.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   The formatted banner string.
#-----------------------------------------------------------------------------------------------------------------------

def _banner_text () -> str:

    # Build the program banner as a box-drawing string.

    inner = BANNER_WIDTH - 2  # Space between the vertical borders.

    lines = [
        "┌" + "─" * inner + "┐",
        f"│{'  Program:    ' + TITLE:<{inner}}│",
        f"│{'  Version:    ' + VERSION:<{inner}}│",
        f"│{'  Build Date: ' + BUILD_DATE:<{inner}}│",
        f"│{'  Author:     ' + AUTHOR:<{inner}}│",
        "└" + "─" * inner + "┘",
    ]

    # Return data to caller.

    return "\n".join ( lines )


#-----------------------------------------------------------------------------------------------------------------------
# Class: _BannerParser
#
# Description:
#
#   ArgumentParser that prints the program banner ahead of help text, and routes argument errors to stderr with a
#   usage message and exit code 2.
#-----------------------------------------------------------------------------------------------------------------------

class _BannerParser ( argparse.ArgumentParser ):

    #-------------------------------------------------------------------------------------------------------------------
    # Function: print_help
    #
    # Description:
    #
    #   Print the banner followed by the standard help text.
    #
    # Arguments:
    #
    #   file : Output stream (defaults to stdout).
    #
    # Returns:
    #
    #   None.
    #-------------------------------------------------------------------------------------------------------------------

    def print_help ( self, file = None ):

        # Print the banner followed by the standard help text.

        if file is None:
            file = sys.stdout

        file.write ( _banner_text () + "\n\n" )
        super ().print_help ( file )

    #-------------------------------------------------------------------------------------------------------------------
    # Function: error
    #
    # Description:
    #
    #   Print a usage message and the error to stderr, then exit with status 2.
    #
    # Arguments:
    #
    #   message : The argparse error message.
    #
    # Returns:
    #
    #   None (exits the process).
    #-------------------------------------------------------------------------------------------------------------------

    def error ( self, message ):

        # Print a usage message and the error to stderr, then exit with status 2.

        sys.stderr.write ( "\n" )
        self.print_usage ( sys.stderr )
        sys.stderr.write ( f"\n{self.prog}: error: {message}\n" )
        sys.exit ( 2 )


#-----------------------------------------------------------------------------------------------------------------------
# Function: _build_parser
#
# Description:
#
#   Build the command-line argument parser. The only user-facing knobs are --year, --filename, and --pdf; --config is
#   a hidden power-user escape hatch suppressed from --help.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   A configured ArgumentParser.
#-----------------------------------------------------------------------------------------------------------------------

def _build_parser () -> argparse.ArgumentParser:

    # Build the command-line argument parser.

    parser = _BannerParser (
        description = "Generate a configurable, printable Year Planner as a Microsoft Word document.",
    )

    parser.add_argument ( "--year", type = int, default = None, help = "Planner year (default: current year)" )
    parser.add_argument ( "--filename", type = str, default = None, help = "Output filename or path (default: year-planner-<year> in the current directory)" )
    parser.add_argument ( "--pdf", action = "store_true", help = "Also produce a PDF (Windows with Microsoft Word only)" )
    parser.add_argument ( "--version", action = "version", version = f"{TITLE} {VERSION} ({BUILD_DATE})" )
    parser.add_argument ( "--config", type = str, default = None, help = argparse.SUPPRESS )

    # Return data to caller.

    return parser


#-----------------------------------------------------------------------------------------------------------------------
# Function: backup_existing_file
#
# Description:
#
#   Backup existing output file by renaming to .bak extension.
#
# Arguments:
#
#   output_path : Path to the output file.
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def backup_existing_file ( output_path: Path ) -> None:

    # Backup existing output file by renaming to .bak extension.

    if output_path.exists ():

        backup_path = output_path.with_suffix ( '.bak' )

        shutil.move ( str ( output_path ), str ( backup_path ) )
        print ( f"  Backed up existing file to: {backup_path}" )


#-----------------------------------------------------------------------------------------------------------------------
# Function: generate_year_planner
#
# Description:
#
#   Generate the Year Planner document. DOCX is always saved; a PDF is produced only when produce_pdf is True, with
#   docx2pdf lazily imported inside that branch so DOCX-only / non-Windows runs never load Microsoft Word or pywin32.
#
# Arguments:
#
#   config      : Loaded configuration object.
#   docx_path   : Path where the .docx document will be saved.
#   produce_pdf : Whether to also convert the document to PDF.
#   pdf_path    : Path where the .pdf document will be saved (used only when produce_pdf is True).
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def generate_year_planner ( config: Config, docx_path: Path, produce_pdf: bool, pdf_path: Path ) -> None:

    # Generate the Year Planner document.

    print ( f"Creating Year Planner for {config.document.year}..." )

    # Ensure output directory exists.

    docx_path = Path ( docx_path )
    docx_path.parent.mkdir ( parents = True, exist_ok = True )

    # Backup existing file if present.

    backup_existing_file ( docx_path )

    # Create document with configured page settings.

    document = create_document ( config )

    # Generate sections.

    print ( "  Generating cover page..." )
    generate_cover_page ( document, config )

    # Start new section for main content.
    # (Separate section allows different header/footer configuration.)

    add_section_break ( document, config )
    print ( "  Generating instructions page..." )
    generate_instructions_page ( document, config )

    # Calendar section (current year + next year).
    # No page break needed - instructions.py already ends on verso.

    print ( "  Generating calendar pages..." )
    generate_calendar_section ( document, config )

    # Table of Contents (non-numbered section).

    add_page_break ( document )
    print ( "  Generating table of contents..." )
    generate_toc ( document, config )

    # Goals page - start new section with page numbering.
    # (Section break creates new page, so no separate page break needed.)

    add_numbered_section_break ( document, config, start_number = 1 )
    print ( "  Generating goals page..." )
    generate_goals_page ( document, config )

    # Backlog section (use minimized page break for precise table fitting).

    add_page_break ( document, minimize_height = True )
    print ( "  Generating backlog..." )
    generate_backlog ( document, config )

    # Week Planner (use minimized page break for precise table fitting).

    add_page_break ( document, minimize_height = True )
    print ( "  Generating week planner..." )
    generate_week_planner ( document, config )

    # Monthly sections (12 months with cover pages).

    add_page_break ( document )
    print ( "  Generating monthly sections..." )
    generate_monthly_sections ( document, config )

    # Terms and Definitions section (use minimized page break for precise table fitting).

    add_page_break ( document, minimize_height = True )
    print ( "  Generating terms and definitions..." )
    generate_terms_definitions ( document, config )

    # Graph paper section (use minimized page break for precise table fitting).

    add_page_break ( document, minimize_height = True )
    print ( "  Generating graph paper..." )
    generate_graph_paper ( document, config )

    # Rear cover (inside: blank, outside: blank).
    # Add non-numbered section break to remove page numbers from rear cover.

    add_non_numbered_section_break ( document, config )
    print ( "  Generating rear cover..." )
    generate_rear_cover ( document, config )
    print ()

    # Save Word document.

    document.save ( docx_path )
    print ( f"Year Planner saved to: {docx_path}" )

    # Convert to PDF only when requested (lazy import keeps Word/pywin32 out of DOCX-only runs).

    if produce_pdf:

        from docx2pdf import convert as convert_to_pdf

        print ( "Converting to PDF..." )
        convert_to_pdf ( str ( docx_path ), str ( pdf_path ) )
        print ( f"PDF saved to: {pdf_path}" )
        print ()


#-----------------------------------------------------------------------------------------------------------------------
# Function: main
#
# Description:
#
#   Main entry point. Parses arguments, validates the year, gates --pdf, loads the bundled configuration, and generates
#   the document.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   Exit code (0 for success, non-zero for error).
#-----------------------------------------------------------------------------------------------------------------------

def main () -> int:

    # Main entry point.

    # Configure stdout for UTF-8 so the banner's box-drawing characters render on Windows.

    if sys.stdout is not None and hasattr ( sys.stdout, "reconfigure" ):
        sys.stdout.reconfigure ( encoding = "utf-8", errors = "replace" )

    args = _build_parser ().parse_args ()

    # Print the program banner.

    print ( _banner_text () )
    print ()

    # Resolve and validate the planner year.

    effective_year = args.year if args.year is not None else date.today ().year

    if not ( 1000 <= effective_year <= 9999 ):

        print ( f"Error: --year must be between 1000 and 9999 (got {effective_year}).", file = sys.stderr )
        return 1

    # Gate the optional --pdf add-on before doing any work (Windows + Microsoft Word only).

    produce_pdf = args.pdf

    if produce_pdf and not word_available ():

        print ( "Error: --pdf requires Microsoft Word, which was not detected. PDF export is supported on Windows "
                "and macOS with Word installed; other platforms are DOCX-only. Omit --pdf to generate the DOCX.",
                file = sys.stderr )

        return 1

    try:

        # Resolve configuration: the bundled resource unless the hidden --config overrides it.

        config_path = args.config if args.config else resource_path ( "config", "config.yaml" )
        config      = load_config ( config_path )

        # Apply the effective year to the loaded configuration. The cover version is taken as-is from the
        # config file's document.version (config/config.yaml) and is intentionally left untouched here.

        config.document.year = effective_year

        print ( f"Loading configuration from: {config_path}" )
        print ( f"  Title:   {config.document.title}" )
        print ( f"  Year:    {config.document.year}" )
        print ( f"  Version: {config.document.version}" )
        print ()

        # Resolve the output paths and generate the document.

        docx_path, pdf_path = resolve_output_paths ( args.filename, effective_year )
        generate_year_planner ( config, docx_path, produce_pdf, pdf_path )

        print ( "Done!" )
        print ()

        # Return data to caller.

        return 0

    except FileNotFoundError as e:

        print ( f"Error: File not found: {e}", file = sys.stderr )

        # Return data to caller.

        return 1

    except Exception as e:

        print ( f"Error: {e}", file = sys.stderr )

        # Return data to caller.

        return 1


if __name__ == "__main__":

    sys.exit ( main () )
