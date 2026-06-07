#-----------------------------------------------------------------------------------------------------------------------
# Program: Year Planner Generator
# Version: 1.1
# Author:  Rohin Gosling
#
# Description:
#
#   Scans Python source files for imports and compares against requirements.txt.
#   Reports any third-party imports that are missing from requirements.txt.
#
# Usage:
#
#   python scripts/check_deps.py
#
#   Exit codes:
#   - 0 : All dependencies accounted for.
#   - 1 : Missing dependencies found.
#-----------------------------------------------------------------------------------------------------------------------

import ast
import sys
from pathlib import Path
from importlib.util import find_spec

# Global constants.

PROJECT_ROOT = Path(__file__).parent.parent       # Project root (parent of scripts/).
SOURCE_DIRS  = [PROJECT_ROOT / "src"]             # Directories to scan for Python files.

# Map package names in requirements.txt to their import names (for packages where they differ).

PACKAGE_TO_IMPORT = {
    "python-docx": "docx",
    "pillow": "PIL",
    "pyyaml": "yaml",
    "docx2pdf": "docx2pdf",
    "typing_extensions": "typing_extensions",
    "lxml": "lxml",
}

# Known standard library modules (Python 3.12+).
# This is a subset — we also use find_spec() for verification.

STDLIB_MODULES = {
    "abc", "argparse", "ast", "asyncio", "base64", "calendar", "collections",
    "contextlib", "copy", "csv", "dataclasses", "datetime", "decimal", "enum",
    "functools", "glob", "gzip", "hashlib", "html", "http", "importlib", "io",
    "itertools", "json", "logging", "math", "mimetypes", "os", "pathlib",
    "pickle", "platform", "pprint", "random", "re", "shutil", "socket", "ssl",
    "string", "struct", "subprocess", "sys", "tempfile", "textwrap", "threading",
    "time", "timeit", "traceback", "typing", "unittest", "urllib", "uuid",
    "warnings", "weakref", "xml", "zipfile", "zlib",
}


#-----------------------------------------------------------------------------------------------------------------------
# Function: parse_requirements
#
# Description:
#
#   Parse requirements.txt and return set of import names.
#
# Arguments:
#
#   req_path : Path to requirements.txt.
#
# Returns:
#
#   Set of import names (converted from package names where needed).
#-----------------------------------------------------------------------------------------------------------------------

def parse_requirements(req_path: Path) -> set[str]:

    # Parse requirements.txt and return set of import names.

    imports = set()

    if not req_path.exists():
        print(f"Warning: {req_path} not found")
        return imports

    for line in req_path.read_text().splitlines():
        line = line.strip()

        # Skip empty lines and comments.

        if not line or line.startswith("#"):
            continue

        # Extract package name (before any version specifier).

        for sep in [">=", "<=", "==", "!=", "~=", ">", "<", "["]:
            if sep in line:
                line = line.split(sep)[0]
                break

        package_name = line.strip().lower()

        # Convert to import name if mapping exists.

        if package_name in PACKAGE_TO_IMPORT:
            imports.add(PACKAGE_TO_IMPORT[package_name])
        else:

            # Default: assume import name matches package name.

            imports.add(package_name.replace("-", "_"))

    # Return data to caller.

    return imports


#-----------------------------------------------------------------------------------------------------------------------
# Function: extract_imports
#
# Description:
#
#   Extract top-level import names from a Python file.
#
# Arguments:
#
#   file_path : Path to Python file.
#
# Returns:
#
#   Set of top-level module names (e.g., 'docx' from 'from docx.shared import Pt').
#-----------------------------------------------------------------------------------------------------------------------

def extract_imports(file_path: Path) -> set[str]:

    # Extract top-level import names from a Python file.

    imports = set()

    try:
        source = file_path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(file_path))
    except (SyntaxError, UnicodeDecodeError) as e:
        print(f"Warning: Could not parse {file_path}: {e}")
        return imports

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:

                # Get top-level module (e.g., 'docx' from 'docx.shared').

                top_level = alias.name.split(".")[0]
                imports.add(top_level)
        elif isinstance(node, ast.ImportFrom):
            if node.module:

                # Get top-level module.

                top_level = node.module.split(".")[0]
                imports.add(top_level)

    # Return data to caller.

    return imports


#-----------------------------------------------------------------------------------------------------------------------
# Function: is_stdlib_module
#
# Description:
#
#   Check if a module is part of the Python standard library.
#
# Arguments:
#
#   module_name : Name of the module.
#
# Returns:
#
#   True if it's a stdlib module.
#-----------------------------------------------------------------------------------------------------------------------

def is_stdlib_module(module_name: str) -> bool:

    # Check if a module is part of the Python standard library.

    if module_name in STDLIB_MODULES:
        return True

    # Try to find the spec and check if it's in stdlib.

    try:
        spec = find_spec(module_name)
        if spec is None:
            return False

        # Check if it's in the stdlib location.

        if spec.origin:
            origin = Path(spec.origin)

            # Stdlib modules are typically in Lib/ not site-packages/.

            return "site-packages" not in str(origin)
    except (ModuleNotFoundError, ValueError):
        return False

    return False


#-----------------------------------------------------------------------------------------------------------------------
# Function: is_local_module
#
# Description:
#
#   Check if a module is a local project module.
#
# Arguments:
#
#   module_name : Name of the module.
#
# Returns:
#
#   True if it's a local module (src, test, etc.).
#-----------------------------------------------------------------------------------------------------------------------

def is_local_module(module_name: str) -> bool:

    # Check if a module is a local project module.

    local_prefixes = {"src", "test", "tests", "scripts"}

    # Return data to caller.

    return module_name in local_prefixes


#-----------------------------------------------------------------------------------------------------------------------
# Function: main
#
# Description:
#
#   Main entry point for the dependency checker.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   Exit code (0 = success, 1 = missing deps found).
#-----------------------------------------------------------------------------------------------------------------------

def main() -> int:

    # Main entry point for the dependency checker.

    print("Dependency Checker")
    print("=" * 40)
    print()

    # Parse requirements.txt.

    req_path = PROJECT_ROOT / "requirements.txt"
    known_packages = parse_requirements(req_path)
    print(f"Packages in requirements.txt: {len(known_packages)}")
    for pkg in sorted(known_packages):
        print(f"  - {pkg}")
    print()

    # Collect all imports from source files.

    all_imports: dict[str, list[Path]] = {}  # module -> list of files using it

    for source_dir in SOURCE_DIRS:
        if not source_dir.exists():
            print(f"Warning: Source directory not found: {source_dir}")
            continue

        for py_file in source_dir.rglob("*.py"):
            imports = extract_imports(py_file)
            for imp in imports:
                if imp not in all_imports:
                    all_imports[imp] = []
                all_imports[imp].append(py_file)

    print(f"Unique imports found: {len(all_imports)}")
    print()

    # Find missing dependencies.

    missing: dict[str, list[Path]] = {}

    for module_name, files in all_imports.items():

        # Skip stdlib.

        if is_stdlib_module(module_name):
            continue

        # Skip local modules.

        if is_local_module(module_name):
            continue

        # Check if it's in requirements.txt.
        # Handle case variations (PIL vs pil).

        if module_name.lower() not in {p.lower() for p in known_packages}:
            missing[module_name] = files

    # Report results.

    if missing:
        print("MISSING DEPENDENCIES")
        print("-" * 40)
        for module_name in sorted(missing.keys()):
            files = missing[module_name]
            print(f"\n  {module_name}")
            print(f"    Used in:")
            for f in files:
                rel_path = f.relative_to(PROJECT_ROOT)
                print(f"      - {rel_path}")
        print()
        print(f"Found {len(missing)} missing package(s).")
        print("Add them to requirements.txt")
        return 1
    else:
        print("All dependencies accounted for.")
        return 0


if __name__ == "__main__":
    sys.exit(main())
