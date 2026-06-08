#-----------------------------------------------------------------------------------------------------------------------
# Module:  paths.py
# Project: Year Planner Generator
# Version: 1.2
# Author:  Rohin Gosling
#
# Description:
#
#   Path-resolution helpers that let the generator run identically from source and from a PyInstaller one-file
#   executable. Bundled resources (config, images) are read via resource_path(); runtime-generated files are written to
#   a writable cache_dir(); user output paths are resolved by resolve_output_paths(); and word_available() gates the
#   optional --pdf add-on. This module imports nothing from the generator so it stays a leaf dependency.
#-----------------------------------------------------------------------------------------------------------------------

import os
import sys
import tempfile

from pathlib import Path

# Global constants.

CACHE_SUBDIR = "YearPlanner"   # Top-level folder name for the writable runtime cache.


#-----------------------------------------------------------------------------------------------------------------------
# Function: resource_path
#
# Description:
#
#   Resolve the absolute path to a bundled resource. In a frozen one-file executable PyInstaller unpacks data files to
#   sys._MEIPASS; when running from source the project root (parent of src/) is used. The same relative parts therefore
#   resolve to byte-identical files in both modes.
#
# Arguments:
#
#   parts : Path components relative to the resource root (e.g. "assets", "images").
#
# Returns:
#
#   Absolute Path to the requested resource.
#-----------------------------------------------------------------------------------------------------------------------

def resource_path ( *parts: str ) -> Path:

    # Resolve a bundled resource against the frozen bundle dir or the source project root.

    if getattr ( sys, "frozen", False ):
        base = Path ( sys._MEIPASS )
    else:
        base = Path ( __file__ ).resolve ().parent.parent

    # Return data to caller.

    return base.joinpath ( *parts )


#-----------------------------------------------------------------------------------------------------------------------
# Function: cache_dir
#
# Description:
#
#   Return a writable directory for runtime-generated files (e.g. a graph-paper grid that is not bundled). A frozen exe
#   may run from a read-only location (Program Files), so generated assets must never be written next to the binary.
#   Uses %LOCALAPPDATA% on Windows, the system temp directory elsewhere. The directory is created if absent.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   Absolute Path to the cache directory (guaranteed to exist).
#-----------------------------------------------------------------------------------------------------------------------

def cache_dir () -> Path:

    # Choose a writable base directory, then ensure the cache folder exists.

    if sys.platform == "win32":
        base = Path ( os.environ.get ( "LOCALAPPDATA", tempfile.gettempdir () ) )
    else:
        base = Path ( tempfile.gettempdir () )

    path = base / CACHE_SUBDIR / "cache"
    path.mkdir ( parents = True, exist_ok = True )

    # Return data to caller.

    return path


#-----------------------------------------------------------------------------------------------------------------------
# Function: resolve_output_paths
#
# Description:
#
#   Resolve the DOCX and PDF output paths from the user-supplied --filename and the planner year. Both paths share a
#   single stem so the optional PDF sits beside the DOCX. Rules:
#     - omitted        : default stem "year-planner-<year>" in the current working directory
#     - bare name      : "<name>.docx" in the current working directory
#     - .docx / .pdf   : trailing extension stripped, then re-appended per format
#     - sub/dir/name   : relative paths resolved against the CWD; missing parents created
#     - existing dir   : default stem placed inside that directory
#     - interior dots  : e.g. "report.v2" -> "report.v2.docx" (appended, not treated as a suffix)
#
# Arguments:
#
#   filename : The raw --filename value, or None for the default.
#   year     : Planner year, used to build the default stem.
#
# Returns:
#
#   Tuple of (docx_path, pdf_path) as absolute Paths sharing a common stem.
#-----------------------------------------------------------------------------------------------------------------------

def resolve_output_paths ( filename: str | None, year: int ) -> tuple [ Path, Path ]:

    # Resolve a shared output stem and directory, then derive the per-format paths.

    default_stem = f"year-planner-{year}"

    if filename is None:

        target_dir = Path.cwd ()
        stem       = default_stem

    else:

        candidate = Path ( filename )

        # Resolve relative paths against the current working directory.

        if not candidate.is_absolute ():

            candidate = Path.cwd () / candidate

        if candidate.is_dir ():

            # An existing directory: place the default stem inside it.

            target_dir = candidate
            stem       = default_stem

        else:

            target_dir = candidate.parent
            name       = candidate.name
            lowered    = name.lower ()

            # Strip a trailing .docx/.pdf if present; otherwise keep the name verbatim (interior dots included).

            if lowered.endswith ( ".docx" ):

                stem = name [ : -len ( ".docx" ) ]

            elif lowered.endswith ( ".pdf" ):

                stem = name [ : -len ( ".pdf" ) ]

            else:

                stem = name

    # Ensure the destination directory exists, then build the per-format paths.

    target_dir.mkdir ( parents = True, exist_ok = True )

    docx_path = target_dir / f"{stem}.docx"
    pdf_path  = target_dir / f"{stem}.pdf"

    # Return data to caller.

    return docx_path, pdf_path


#-----------------------------------------------------------------------------------------------------------------------
# Function: word_available
#
# Description:
#
#   Detect whether Microsoft Word is available for the optional --pdf add-on. PDF conversion (via docx2pdf) drives Word
#   through COM on Windows and through AppleScript on macOS, so both platforms are supported; every other platform
#   (e.g. Linux) returns False. Detection imports are guarded so this is only exercised when --pdf is requested.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   True if Word appears to be installed and usable, False otherwise.
#-----------------------------------------------------------------------------------------------------------------------

def word_available () -> bool:

    # Detect Microsoft Word per platform.

    if sys.platform == "win32":

        return _word_available_windows ()

    if sys.platform == "darwin":

        return _word_available_macos ()

    # Return data to caller. Other platforms (e.g. Linux) cannot run Microsoft Word.

    return False


#-----------------------------------------------------------------------------------------------------------------------
# Function: _word_available_windows
#
# Description:
#
#   Detect Microsoft Word on Windows: probe the registry for the Word.Application COM class, then fall back to
#   instantiating the COM server. win32com (pywin32) is resolved dynamically so this optional, Windows-only dependency
#   stays out of the cross-platform requirements and out of static dependency analysis (scripts/check_deps.py); the
#   frozen exe still bundles it via docx2pdf's own import.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   True if Word appears to be installed and usable, False otherwise.
#-----------------------------------------------------------------------------------------------------------------------

def _word_available_windows () -> bool:

    # Registry probe: a registered Word.Application\CLSID indicates the COM server is present.

    try:

        import winreg

        with winreg.OpenKey ( winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID" ):

            return True

    except ( OSError, ImportError ):

        pass

    # COM fallback: attempt to instantiate Word and immediately quit.

    try:

        import importlib

        win32com_client = importlib.import_module ( "win32com.client" )
        word            = win32com_client.Dispatch ( "Word.Application" )
        word.Quit ()

        return True

    except Exception:

        return False


#-----------------------------------------------------------------------------------------------------------------------
# Function: _word_available_macos
#
# Description:
#
#   Detect Microsoft Word for Mac. Checks the standard application locations first, then falls back to a Launch Services
#   query for a non-standard install. docx2pdf converts via AppleScript on macOS, so a detected Word enables --pdf.
#
#   NOTE: the macOS --pdf path is wired up but has NOT been tested by the author.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   True if Word appears to be installed, False otherwise.
#-----------------------------------------------------------------------------------------------------------------------

def _word_available_macos () -> bool:

    # Standard application locations (no subprocess required).

    candidates = [
        Path ( "/Applications/Microsoft Word.app" ),
        Path.home () / "Applications" / "Microsoft Word.app",
    ]

    if any ( candidate.is_dir () for candidate in candidates ):
        return True

    # Fall back to Launch Services for Word installed in a non-standard location.

    try:

        import subprocess

        result = subprocess.run (
            [ "osascript", "-e", 'id of application "Microsoft Word"' ],
            capture_output = True, text = True, timeout = 10
        )

        return result.returncode == 0

    except Exception:

        return False
