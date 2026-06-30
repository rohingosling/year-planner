#-----------------------------------------------------------------------------------------------------------------------
# Module:  rear_cover.py
# Project: Year Planner Generator
# Version: 1.1
# Author:  Rohin Gosling
#
# Description:
#
#   Rear cover generator for the Year Planner. Generates blank rear cover pages reserved for future
#   barcode/branding.
#-----------------------------------------------------------------------------------------------------------------------

from docx import Document

from src.config   import Config
from src.document import add_page_break, add_config_info_overlay


#-----------------------------------------------------------------------------------------------------------------------
# Function: generate_rear_cover
#
# Description:
#
#   Generate the rear cover (blank recto and verso).
#
#   Reserved for future barcode/branding.
#
# Arguments:
#
#   document : The Word document to add the rear cover to.
#   config   : Configuration with document settings.
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def generate_rear_cover ( document: Document, config: Config ) -> None:

    # Generate the rear cover (blank recto and verso).

    # Inside rear cover (recto) — blank.
    # The previous section's page break positions us on this recto.
    # Add minimal content to ensure the page exists.

    para                               = document.add_paragraph ()
    para.paragraph_format.space_before = 0
    para.paragraph_format.space_after  = 0
    add_config_info_overlay ( document, config, is_recto = True )

    # Outside rear cover (verso) — blank.

    add_page_break ( document )
    add_config_info_overlay ( document, config, is_recto = False )
