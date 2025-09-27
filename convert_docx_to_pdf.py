"""
DOCX to PDF Converter
This script combines all DOCX files in the Weekly Reports folder into a single PDF.
"""

import sys
from pathlib import Path
import tempfile

try:
    from docx2pdf import convert
    from PyPDF2 import PdfMerger
except ImportError:
    print("Error: Required libraries not found.")
    print("Please install them using: pip install docx2pdf PyPDF2")
    sys.exit(1)


def combine_docx_to_single_pdf(input_folder="Weekly Reports", output_filename="Combined_Weekly_Reports.pdf"):
    """
    Convert all DOCX files in the input folder to a single combined PDF.
    
    Args:
        input_folder (str): Path to folder containing DOCX files
        output_filename (str): Name of the output PDF file
    """
    
    # Get the current script directory
    script_dir = Path(__file__).parent
    input_path = script_dir / input_folder
    output_path = script_dir / output_filename
    
    # Check if input folder exists
    if not input_path.exists():
        print(f"Error: Input folder '{input_folder}' not found.")
        return False
    
    # Get all DOCX files and sort them alphabetically
    docx_files = sorted(list(input_path.glob("*.docx")))
    
    if not docx_files:
        print(f"No DOCX files found in '{input_folder}'")
        return False
    
    print(f"Found {len(docx_files)} DOCX files to combine into single PDF...")
    
    # Create a temporary directory for individual PDFs
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        pdf_files = []
        successful_conversions = 0
        failed_conversions = 0
        
        # Convert each DOCX to individual PDF first
        for i, docx_file in enumerate(docx_files, 1):
            try:
                # Create temporary PDF filename
                temp_pdf_name = f"temp_{i:03d}_{docx_file.stem}.pdf"
                temp_pdf_path = temp_path / temp_pdf_name
                
                print(f"Converting ({i}/{len(docx_files)}): {docx_file.name}")
                
                # Convert DOCX to PDF
                convert(str(docx_file), str(temp_pdf_path))
                
                if temp_pdf_path.exists():
                    pdf_files.append(temp_pdf_path)
                    file_size = temp_pdf_path.stat().st_size / 1024  # Size in KB
                    print(f"  ✓ Converted successfully ({file_size:.1f} KB)")
                    successful_conversions += 1
                else:
                    print(f"  ✗ Failed to create PDF")
                    failed_conversions += 1
                    
            except Exception as e:
                print(f"  ✗ Error converting {docx_file.name}: {str(e)}")
                failed_conversions += 1
        
        # Combine all PDFs into a single file
        if pdf_files:
            try:
                print(f"\nCombining {len(pdf_files)} PDFs into single file...")
                
                merger = PdfMerger()
                
                for pdf_file in pdf_files:
                    merger.append(str(pdf_file))
                
                # Write the combined PDF
                with open(output_path, 'wb') as output_file:
                    merger.write(output_file)
                
                merger.close()
                
                if output_path.exists():
                    final_size = output_path.stat().st_size / 1024  # Size in KB
                    print(f"  ✓ Combined PDF created successfully ({final_size:.1f} KB)")
                    
                    # Summary
                    print("\n" + "="*50)
                    print("COMBINATION SUMMARY")
                    print("="*50)
                    print(f"Total DOCX files found: {len(docx_files)}")
                    print(f"Successfully converted: {successful_conversions}")
                    print(f"Failed conversions: {failed_conversions}")
                    print(f"Combined PDF: {output_path}")
                    print(f"Final size: {final_size:.1f} KB")
                    
                    return True
                else:
                    print("  ✗ Failed to create combined PDF")
                    return False
                    
            except Exception as e:
                print(f"  ✗ Error combining PDFs: {str(e)}")
                return False
        else:
            print("No PDFs were successfully created to combine.")
            return False


def main():
    """Main function to run the conversion process."""
    
    print("DOCX to Single PDF Combiner")
    print("="*35)
    
    # Check if we're in the right directory
    script_dir = Path(__file__).parent
    weekly_reports_path = script_dir / "Weekly Reports"
    
    if not weekly_reports_path.exists():
        print("Error: 'Weekly Reports' folder not found in the current directory.")
        print(f"Current directory: {script_dir}")
        return
    
    # Run the conversion and combination
    success = combine_docx_to_single_pdf()
    
    if success:
        print("\n🎉 Combination completed successfully!")
    else:
        print("\n❌ Combination failed or no files were processed.")


if __name__ == "__main__":
    main()
