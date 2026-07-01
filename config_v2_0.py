"""
PEPS BOM Automation Tool - Configuration
Version: 2.0
Date: 21 May 2026

This file contains all configuration settings, file paths, and RAMCO constants.
"""

import os

# ============================================================================
# SERVER CONFIGURATION
# ============================================================================
APP_HOST = '0.0.0.0'       # Bind all interfaces (localhost + LAN)
APP_PORT = 5020            # Port 5020
APP_LAN_IP = '192.168.0.133'  # LAN IP for display in startup message
DEBUG_MODE = False         # Set to False in production

# ============================================================================
# FILE PATHS
# ============================================================================
# Root folder containing FG and SFG product xlsm macro files
BOM_FILES_ROOT = r'D:\Hari JR. DATA\BOM\Source Macro -03.06.2026'

# Item master for component lookup in New Product Wizard
ITEM_MASTER_PATH = r'D:\Hari JR. DATA\BOM\Automation\Variety Master\Items & variety details.xlsx'

# SFG Size Variety master — has BOX and ASSEMBLY sheets, each listing valid
# Length/Width/Height combinations. Filenames containing "BOX" use the BOX
# sheet's sizes; filenames containing "AS" (assembly) use the ASSEMBLY sheet's.
SFG_SIZE_VARIETY_PATH = r'D:\Hari JR. DATA\BOM\Automation\Size Variety\SFG Size Variaty -23.06.2026.xlsx'

# Output folder for generated xlsx files
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'output')

# Database file
DATABASE_PATH = os.path.join(os.path.dirname(__file__), 'bom_tool.db')

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ============================================================================
# PRODUCT CONFIGURATION
# ============================================================================
# Valid height values across all product families
VALID_HEIGHTS = ['3', '4', '5', '6', '8', '10', '12', '14', '16']

# Standard length values (in inches)
STANDARD_LENGTHS = [72, 75, 78, 80, 84]

# ============================================================================
# RAMCO CONSTANTS - CreateProductStructure Import Format (43 columns)
# ============================================================================
# These values are hardcoded as per PEPS-BOM-PHASE2-001 v3 responses
RAMCO_CONSTANTS = {
    # Column 7: Maintained By
    'maintained_by': 'PEPS',
    
    # Column 8: PS Usage
    'ps_usage': 'Standard',
    
    # Column 9: PS Category
    'ps_category': 'Production',
    
    # Column 14: Default For
    'default_for': 'MRP&ROLLUP',
    
    # Column 29: I/O
    'io': 'INPUT',
    
    # Column 30: Valid From
    'valid_from': '2019-04-01',
    
    # Column 31: Valid To
    'valid_to': '2065-03-31',
    
    # Column 32: Min. Qty.
    'min_qty': 1,
    
    # Column 33: Max. Qty.
    'max_qty': 100000000,
    
    # Column 36: Stock Status
    'stock_status': 'Accepted',
    
    # Column 37: Critical
    'critical': 'No',
    
    # Column 38: Use-Up Option
    'use_up_option': 'No',
    
    # Column 12: Process Plan No. - blank pending ERP team input
    'process_plan_no': '',
}

# ============================================================================
# PS DESCRIPTION AUTO-GENERATION PATTERN
# ============================================================================
# Pattern: Peps [Type] [Variety] [Category] [Colour] LxWxH
# Examples from PEPS-BOM-PHASE2-001 v3:
#   Peps Springkoil Bonnell Normal Maroon 75X42X05
#   Peps Bonnell Comfort Normal Pink 78X36X06
#   Peps Pocketed Supreme Normal Beige 72X30X06

def generate_ps_description(product_type, variety, category, colour, length, width, height):
    """
    Generate PS Description following PEPS pattern.
    
    Args:
        product_type: e.g., 'Springkoil', 'Bonnell', 'Pocketed'
        variety: e.g., 'Bonnell', 'Comfort', 'Supreme'
        category: e.g., 'Normal', 'Pillowtop', 'Eurotop'
        colour: e.g., 'Maroon', 'Pink', 'Beige'
        length, width, height: dimensions
    
    Returns:
        str: Formatted PS Description
    """
    # Pad dimensions to 2 digits
    l = str(length).zfill(2)
    w = str(width).zfill(2)
    h = str(height).zfill(2)
    
    return f"Peps {product_type} {variety} {category} {colour} {l}X{w}X{h}"

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================
LOG_LEVEL = 'INFO'  # DEBUG, INFO, WARNING, ERROR
LOG_FILE = os.path.join(os.path.dirname(__file__), 'bom_tool.log')

# ============================================================================
# FILE SCANNING CONFIGURATION
# ============================================================================
# File extensions to scan
VALID_EXTENSIONS = ['.xlsm']

# Maximum file size to process (in MB)
MAX_FILE_SIZE_MB = 50

# ============================================================================
# API CONFIGURATION
# ============================================================================
# Maximum products that can be selected for bulk download
MAX_BULK_PRODUCTS = 50

# Maximum rows in a single xlsx output
MAX_OUTPUT_ROWS = 100000

# Timeout for background processes (seconds)
PROCESS_TIMEOUT = 300

# ============================================================================
# VERSION INFO
# ============================================================================
APP_VERSION = '2.0'
APP_NAME = 'PEPS BOM Automation Tool'
APP_DESCRIPTION = 'Phase 2 - Configurator, Global Replace, New Product Wizard, Nearest BOM'
