#-----------------------------------------------------------------------------------------------------------------------
# Module:  config.py
# Project: Year Planner Generator
# Version: 1.1
# Author:  Rohin Gosling
#
# Description:
#
#   Configuration loader for the Year Planner generator. Loads and validates YAML configuration files,
#   returning a typed Config object with nested dataclass fields for each configuration section.
#-----------------------------------------------------------------------------------------------------------------------

from dataclasses import dataclass
from pathlib     import Path
from typing      import Any

import yaml


#-----------------------------------------------------------------------------------------------------------------------
# Class: DocumentConfig
#
# Description:
#
#   Document metadata configuration.
#
# Attributes:
#
#   title   : Document title text shown on the cover page.
#   version : Document version number string.
#   year    : The year the planner is generated for.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class DocumentConfig:
    title: str
    version: str
    year: int


#-----------------------------------------------------------------------------------------------------------------------
# Class: PageConfig
#
# Description:
#
#   Page layout configuration.
#
# Attributes:
#
#   width                : Page width in centimeters (A4 = 21.0).
#   height               : Page height in centimeters (A4 = 29.7).
#   margin_top           : Top margin in centimeters.
#   margin_bottom        : Bottom margin in centimeters.
#   margin_left          : Left (outside) margin in centimeters.
#   margin_right         : Right (outside) margin in centimeters.
#   gutter_size          : Binding gutter size in centimeters (mirror margins).
#   page_number_position : Vertical distance of page numbers from the bottom edge in centimeters.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class PageConfig:
    width: float
    height: float
    margin_top: float
    margin_bottom: float
    margin_left: float
    margin_right: float
    gutter_size: float
    page_number_position: float


#-----------------------------------------------------------------------------------------------------------------------
# Class: BorderConfig
#
# Description:
#
#   Table border configuration.
#
# Attributes:
#
#   thickness : Line thickness in points.
#   grayscale : Border color as grayscale percentage (0 = white, 100 = black).
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class BorderConfig:
    thickness: float
    grayscale: int


#-----------------------------------------------------------------------------------------------------------------------
# Class: TitleRowConfig
#
# Description:
#
#   Table title row configuration.
#
# Attributes:
#
#   height               : Row height in points.
#   background_grayscale : Background color as grayscale percentage (0 = white, 100 = black).
#   font_size            : Font size in points.
#   font_grayscale       : Font color as grayscale percentage (0 = white, 100 = black).
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class TitleRowConfig:
    height: float
    background_grayscale: int
    font_size: float
    font_grayscale: int


#-----------------------------------------------------------------------------------------------------------------------
# Class: HeaderRowConfig
#
# Description:
#
#   Table header row configuration.
#
# Attributes:
#
#   height               : Row height in points.
#   background_grayscale : Background color as grayscale percentage (0 = white, 100 = black).
#   font_size            : Font size in points.
#   font_grayscale       : Font color as grayscale percentage (0 = white, 100 = black).
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class HeaderRowConfig:
    height: float
    background_grayscale: int
    font_size: float
    font_grayscale: int


#-----------------------------------------------------------------------------------------------------------------------
# Class: ContentRowConfig
#
# Description:
#
#   Table content row configuration.
#
# Attributes:
#
#   font_size      : Font size in points.
#   font_grayscale : Font color as grayscale percentage (0 = white, 100 = black).
#   font_italic    : Whether content row text is rendered in italic.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class ContentRowConfig:
    font_size: float
    font_grayscale: int
    font_italic: bool


#-----------------------------------------------------------------------------------------------------------------------
# Class: TableConfig
#
# Description:
#
#   Table styling configuration.
#
# Attributes:
#
#   border      : Border thickness and color settings.
#   title_row   : Title row appearance settings.
#   header_row  : Header row appearance settings.
#   content_row : Content row appearance settings.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class TableConfig:
    border: BorderConfig
    title_row: TitleRowConfig
    header_row: HeaderRowConfig
    content_row: ContentRowConfig


#-----------------------------------------------------------------------------------------------------------------------
# Class: DebugConfig
#
# Description:
#
#   Debug mode configuration.
#
# Attributes:
#
#   enabled             : Whether debug visualization (red borders, blue gutter lines) is active.
#   config_info_overlay : Whether the configuration info overlay is shown on each page.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class DebugConfig:
    enabled: bool
    config_info_overlay: bool


#-----------------------------------------------------------------------------------------------------------------------
# Class: ConfigInfoOverlayConfig
#
# Description:
#
#   Configuration info overlay settings.
#
# Attributes:
#
#   bottom          : Distance from the bottom edge of the page in centimeters.
#   right           : Distance from the right edge on recto pages in centimeters.
#   left            : Distance from the left edge on verso pages in centimeters.
#   width           : Text box width in centimeters.
#   title           : Overlay title text string.
#   title_font_size : Title text font size in points.
#   data_font_size  : Field data font size in points.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class ConfigInfoOverlayConfig:

    bottom: float
    right: float
    left: float
    width: float
    title: str
    title_font_size: float
    data_font_size: float


#-----------------------------------------------------------------------------------------------------------------------
# Class: ContactTableConfig
#
# Description:
#
#   Contact table configuration.
#
# Attributes:
#
#   row_height  : Height of each contact table row in points.
#   label_width : Width of the label column in centimeters.
#   value_width : Width of the value column in centimeters.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class ContactTableConfig:

    row_height: float
    label_width: float
    value_width: float


#-----------------------------------------------------------------------------------------------------------------------
# Class: CoverConfig
#
# Description:
#
#   Cover page configuration.
#
# Attributes:
#
#   contact_fields : List of field name strings shown in the contact info table.
#   contact_table  : Dimensional settings for the contact info table.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class CoverConfig:

    contact_fields: list [ str ]
    contact_table: ContactTableConfig


#-----------------------------------------------------------------------------------------------------------------------
# Class: Config
#
# Description:
#
#   Main configuration container holding all typed sub-configurations parsed from config.yaml.
#
# Attributes:
#
#   debug               : Debug visualization settings.
#   config_info_overlay : Config info overlay positioning and font settings.
#   document            : Document metadata (title, version, year).
#   page                : Page size, margins, and gutter settings.
#   table               : Table styling for all row types and borders.
#   cover               : Cover page contact table settings.
#   raw                 : Raw YAML dict preserved for sections that read config directly.
#-----------------------------------------------------------------------------------------------------------------------

@dataclass
class Config:

    debug: DebugConfig
    config_info_overlay: ConfigInfoOverlayConfig
    document: DocumentConfig
    page: PageConfig
    table: TableConfig
    cover: CoverConfig
    raw: dict [ str, Any ]


#-----------------------------------------------------------------------------------------------------------------------
# Function: load_config
#
# Description:
#
#   Load configuration from a YAML file.
#
#   Reads and parses the YAML configuration file at the given path, constructs and returns a fully typed
#   Config object with nested dataclass fields for every configuration section.
#
# Arguments:
#
#   config_path : Path to the YAML configuration file.
#
# Returns:
#
#   Config object with parsed configuration values.
#
# Raises:
#
#   FileNotFoundError : If the configuration file does not exist.
#   yaml.YAMLError    : If the configuration file contains invalid YAML.
#-----------------------------------------------------------------------------------------------------------------------

def load_config ( config_path: str | Path ) -> Config:

    # Load configuration from a YAML file.

    config_path = Path ( config_path )

    with open ( config_path, 'r', encoding = 'utf-8' ) as f:
        raw = yaml.safe_load ( f )

    document = DocumentConfig (
        title   = raw [ 'document' ] [ 'title' ],
        version = raw [ 'document' ] [ 'version' ],
        year    = raw [ 'document' ] [ 'year' ]
    )

    page = PageConfig (
        width                = raw [ 'page' ] [ 'width' ],
        height               = raw [ 'page' ] [ 'height' ],
        margin_top           = raw [ 'page' ] [ 'margin_top' ],
        margin_bottom        = raw [ 'page' ] [ 'margin_bottom' ],
        margin_left          = raw [ 'page' ] [ 'margin_left' ],
        margin_right         = raw [ 'page' ] [ 'margin_right' ],
        gutter_size          = raw [ 'page' ] [ 'gutter_size' ],
        page_number_position = raw [ 'page' ] [ 'page_number_position' ]
    )

    # Parse table config with nested structures

    table_raw = raw [ 'table' ]

    border = BorderConfig (
        thickness = table_raw [ 'border' ] [ 'thickness' ],
        grayscale = table_raw [ 'border' ] [ 'grayscale' ]
    )

    title_row = TitleRowConfig (
        height               = table_raw [ 'title_row' ] [ 'height' ],
        background_grayscale = table_raw [ 'title_row' ] [ 'background_grayscale' ],
        font_size            = table_raw [ 'title_row' ] [ 'font_size' ],
        font_grayscale       = table_raw [ 'title_row' ] [ 'font_grayscale' ]
    )

    header_row = HeaderRowConfig (
        height               = table_raw [ 'header_row' ] [ 'height' ],
        background_grayscale = table_raw [ 'header_row' ] [ 'background_grayscale' ],
        font_size            = table_raw [ 'header_row' ] [ 'font_size' ],
        font_grayscale       = table_raw [ 'header_row' ] [ 'font_grayscale' ]
    )

    content_row = ContentRowConfig (
        font_size      = table_raw [ 'content_row' ] [ 'font_size' ],
        font_grayscale = table_raw [ 'content_row' ] [ 'font_grayscale' ],
        font_italic    = table_raw [ 'content_row' ] [ 'font_italic' ]
    )

    table = TableConfig (
        border      = border,
        title_row   = title_row,
        header_row  = header_row,
        content_row = content_row
    )

    contact_table = ContactTableConfig (
        row_height  = raw [ 'cover' ] [ 'contact_table' ] [ 'row_height' ],
        label_width = raw [ 'cover' ] [ 'contact_table' ] [ 'label_width' ],
        value_width = raw [ 'cover' ] [ 'contact_table' ] [ 'value_width' ]
    )

    cover = CoverConfig (
        contact_fields = raw [ 'cover' ] [ 'contact_fields' ],
        contact_table  = contact_table
    )

    # Parse debug config (handle both old boolean and new dict format)

    debug_raw = raw.get ( 'debug', {} )

    if isinstance ( debug_raw, bool ):

        # Backwards compatibility with old format

        debug = DebugConfig ( enabled = debug_raw, config_info_overlay = False )

    else:

        debug = DebugConfig (
            enabled             = debug_raw.get ( 'enabled', False ),
            config_info_overlay = debug_raw.get ( 'config_info_overlay', False )
        )

    # Parse config info overlay settings

    overlay_raw = raw.get ( 'config_info_overlay', {} )
    config_info_overlay = ConfigInfoOverlayConfig (
        bottom          = overlay_raw.get ( 'bottom', 1.8 ),
        right           = overlay_raw.get ( 'right', 1.2 ),
        left            = overlay_raw.get ( 'left', 1.2 ),
        width           = overlay_raw.get ( 'width', 6.0 ),
        title           = overlay_raw.get ( 'title', 'Config Info' ),
        title_font_size = overlay_raw.get ( 'title_font_size', 8 ),
        data_font_size  = overlay_raw.get ( 'data_font_size', 6 )
    )

    # Return data to caller.

    return Config (
        debug               = debug,
        config_info_overlay = config_info_overlay,
        document            = document,
        page                = page,
        table               = table,
        cover               = cover,
        raw                 = raw
    )
