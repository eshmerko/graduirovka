import os
import re
import argparse
from pathlib import Path
import pypdf
import pandas as pd
import tabula
import numpy as np
from typing import List, Dict, Tuple, Optional

def extract_tables_from_pdf(pdf_path: str) -> List[pd.DataFrame]:
    """Extract tables from a PDF file using tabula."""
    # Extract all tables from the PDF
    tables = tabula.read_pdf(
        pdf_path,
        pages='all',
        multiple_tables=True,
        lattice=True,  # Use lattice mode for tables with grid lines
        guess=False,   # Don't guess table structure
        stream=False   # Don't use stream mode
    )
    
    return tables

def identify_calibration_tables(tables: List[pd.DataFrame]) -> List[Tuple[pd.DataFrame, str, str]]:
    """Identify calibration tables among extracted tables and focus on level and volume columns only."""
    calibration_tables = []
    
    for table in tables:
        # Identify patterns in column headers and data to find the calibration table structure
        level_col = None
        volume_col = None
        
        # Strategy 1: Check column headers directly
        for col in table.columns:
            if isinstance(col, str):
                col_lower = col.lower()
                if "уровень наполнения" in col_lower or "уровень" in col_lower and "см" in col_lower:
                    level_col = col
                elif "вместимость" in col_lower and "м3" in col_lower:
                    volume_col = col
        
        # If found through column headers
        if level_col and volume_col:
            calibration_tables.append((table, level_col, volume_col))
            continue
            
        # Strategy 2: Check first rows for header information
        if not table.empty and len(table.columns) >= 3:  # At least 3 columns expected in the table
            # Check first few rows for header-like text
            for i in range(min(3, len(table))):
                row = table.iloc[i]
                for j, val in enumerate(row):
                    if not isinstance(val, str):
                        continue
                    val_lower = val.lower()
                    
                    # Look for level column (first column in each block)
                    if ("уровень наполнения" in val_lower or "уровень" in val_lower) and "см" in val_lower:
                        level_col = table.columns[j]
                    
                    # Look for volume column (second column in each block)
                    elif "вместимость" in val_lower and "м3" in val_lower:
                        volume_col = table.columns[j]
                
                if level_col and volume_col:
                    # Found both columns, use data below these headers
                    filtered_table = table.iloc[i+1:].copy()
                    filtered_table.columns = table.columns
                    calibration_tables.append((filtered_table, level_col, volume_col))
                    break
        
        # Strategy 3: For tables with consistent structure but no clear headers
        # Look for blocks with numeric values where pattern is: integer, decimal number, decimal ~0.08xx
        if not (level_col and volume_col) and not table.empty:
            numeric_columns = []
            for col_idx in range(len(table.columns)):
                # Check if column contains mostly numeric values
                values = table[table.columns[col_idx]].dropna()
                numeric_count = sum(1 for v in values if isinstance(v, (int, float)) or 
                                   (isinstance(v, str) and re.match(r'^\d+\.?\d*, v.replace(",", .')))
                if numeric_count > len(values) * 0.7:  # More than 70% numeric
                    numeric_columns.append(col_idx)
            
            # Look for groups of three columns, where the first is likely level, second is volume
            for i in range(len(numeric_columns) - 2):  # Need at least 3 consecutive columns
                col1_idx = numeric_columns[i]
                col2_idx = numeric_columns[i+1]
                col3_idx = numeric_columns[i+2]
                
                # Check if third column has values around 0.08xx (coefficients to ignore)
                coef_pattern = False
                for val in table[table.columns[col3_idx]].dropna():
                    if isinstance(val, str):
                        val = val.replace(',', '.').strip()
                        if re.match(r'0\.08\d+', val):
                            coef_pattern = True
                            break
                    elif isinstance(val, float) and 0.08 <= val <= 0.09:
                        coef_pattern = True
                        break
                
                # If this looks like a calibration block, take the first two columns
                if coef_pattern:
                    level_col = table.columns[col1_idx]
                    volume_col = table.columns[col2_idx]
                    calibration_tables.append((table, level_col, volume_col))
                    break
    
    return calibration_tables

def extract_level_volume_pairs(calibration_tables: List[Tuple[pd.DataFrame, str, str]]) -> List[Tuple[int, float]]:
    """Extract level-volume pairs from calibration tables."""
    level_volume_pairs = []
    
    for table, level_col, volume_col in calibration_tables:
        for _, row in table.iterrows():
            level_value = row[level_col]
            volume_value = row[volume_col]
            
            # Skip rows with missing values
            if pd.isna(level_value) or pd.isna(volume_value):
                continue
            
            # Convert to string and clean up
            if isinstance(level_value, str):
                level_str = level_value.strip()
                # Extract only numbers from the level string
                level_match = re.search(r'\d+', level_str)
                if level_match:
                    level = int(level_match.group())
                else:
                    continue
            else:
                try:
                    level = int(level_value)
                except (ValueError, TypeError):
                    continue
            
            if isinstance(volume_value, str):
                volume_str = volume_value.strip().replace(' ', '')
                # Replace comma with dot for decimal point
                volume_str = volume_str.replace(',', '.')
                # Extract floating point number
                volume_match = re.search(r'\d+\.\d+', volume_str)
                if volume_match:
                    volume = float(volume_match.group())
                else:
                    # Try to convert the whole string
                    try:
                        volume = float(volume_str)
                    except (ValueError, TypeError):
                        continue
            else:
                try:
                    volume = float(volume_value)
                except (ValueError, TypeError):
                    continue
            
            level_volume_pairs.append((level, volume))
    
    # Sort by level
    level_volume_pairs.sort(key=lambda x: x[0])
    
    return level_volume_pairs

def fallback_extraction(pdf_path: str) -> List[Tuple[int, float]]:
    """Fallback extraction method using regex patterns on raw text."""
    level_volume_pairs = []
    reader = pypdf.PdfReader(pdf_path)
    
    for page in reader.pages:
        text = page.extract_text()
        
        # Look for patterns like "300 258.217" in the text
        # This regex looks for a pattern of digits followed by space and then digits with optional decimal point
        pattern = r'(\d+)\s+(\d+[\.,]\d+)'
        matches = re.findall(pattern, text)
        
        for match in matches:
            try:
                level = int(match[0])
                volume = float(match[1].replace(',', '.'))
                level_volume_pairs.append((level, volume))
            except (ValueError, IndexError):
                continue
    
    # Sort by level
    level_volume_pairs.sort(key=lambda x: x[0])
    
    return level_volume_pairs

def write_to_txt(pairs: List[Tuple[int, float]], output_path: str):
    """Write level-volume pairs to text file in the desired format."""
    with open(output_path, 'w') as f:
        for level, volume in pairs:
            f.write(f"{level}~{volume:.3f}\n")

def clean_and_filter_pairs(pairs: List[Tuple[int, float]]) -> List[Tuple[int, float]]:
    """Clean and filter the pairs to remove duplicates and ensure proper sorting."""
    # Remove duplicates
    unique_pairs = list(set(pairs))
    
    # Sort by level
    sorted_pairs = sorted(unique_pairs, key=lambda x: x[0])
    
    return sorted_pairs

def main():
    parser = argparse.ArgumentParser(description='Extract calibration data from PDF tables.')
    parser.add_argument('pdf_path', help='Path to the PDF file')
    parser.add_argument('--output', '-o', help='Output text file path', default=None)
    args = parser.parse_args()
    
    pdf_path = args.pdf_path
    output_path = args.output or os.path.splitext(pdf_path)[0] + '.txt'
    
    print(f"Processing PDF: {pdf_path}")
    
    try:
        # Try tabula extraction first
        tables = extract_tables_from_pdf(pdf_path)
        calibration_tables = identify_calibration_tables(tables)
        level_volume_pairs = extract_level_volume_pairs(calibration_tables)
        
        # If tabula doesn't find enough data, try fallback method
        if len(level_volume_pairs) < 10:
            print("Tabula extraction produced insufficient results. Trying fallback method...")
            level_volume_pairs = fallback_extraction(pdf_path)
        
        # Clean and filter the pairs
        final_pairs = clean_and_filter_pairs(level_volume_pairs)
        
        if final_pairs:
            write_to_txt(final_pairs, output_path)
            print(f"Successfully extracted {len(final_pairs)} level-volume pairs.")
            print(f"Results saved to: {output_path}")
        else:
            print("No calibration data found in the PDF.")
            
    except Exception as e:
        print(f"Error processing PDF: {e}")
        # Try fallback method if tabula fails
        try:
            print("Trying fallback extraction method...")
            level_volume_pairs = fallback_extraction(pdf_path)
            final_pairs = clean_and_filter_pairs(level_volume_pairs)
            
            if final_pairs:
                write_to_txt(final_pairs, output_path)
                print(f"Fallback method extracted {len(final_pairs)} level-volume pairs.")
                print(f"Results saved to: {output_path}")
            else:
                print("No calibration data found in the PDF.")
        except Exception as e:
            print(f"Fallback extraction also failed: {e}")

if __name__ == "__main__":
    main()