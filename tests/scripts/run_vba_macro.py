#!/usr/bin/env python3
"""
Run VBA Macro via COM Automation

This script uses Windows COM automation to run the VBA timeline generator macro.
Requires Windows with PowerPoint and Excel installed.

Usage:
    python scripts/run_vba_macro.py [--excel-file FILE] [--output FILE] [--visible]
"""
import argparse
import os
import sys
import time
from pathlib import Path


def check_office_installed():
    """Check if Microsoft Office is installed."""
    try:
        import win32com.client
        
        # Try to create Excel application
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
        except Exception as e:
            print(f"Excel not available: {e}")
            return False
        
        # Try to create PowerPoint application
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Quit()
        except Exception as e:
            print(f"PowerPoint not available: {e}")
            return False
        
        return True
    except ImportError:
        print("pywin32 not installed. Install with: pip install pywin32")
        return False


def run_vba_macro(
    excel_file: str,
    vba_file: str,
    output_file: str,
    visible: bool = False
) -> bool:
    """
    Run the VBA timeline generator macro.

    Args:
        excel_file: Path to Excel data file
        vba_file: Path to VBA .bas file
        output_file: Path for output PowerPoint file
        visible: Whether to show Office applications

    Returns:
        True if successful, False otherwise
    """
    import win32com.client
    from win32com.client import constants
    
    excel_app = None
    ppt_app = None
    
    try:
        print(f"Excel file: {os.path.abspath(excel_file)}")
        print(f"VBA file: {os.path.abspath(vba_file)}")
        print(f"Output file: {os.path.abspath(output_file)}")
        print()
        
        # Open Excel with data
        print("Opening Excel...")
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = visible
        excel_app.DisplayAlerts = False
        
        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
        
        workbook = excel_app.Workbooks.Open(os.path.abspath(excel_file))
        print(f"  Opened workbook: {workbook.Name}")
        
        # Verify TimelineData sheet exists
        sheet_names = [sheet.Name for sheet in workbook.Sheets]
        if "TimelineData" not in sheet_names:
            raise ValueError(f"Sheet 'TimelineData' not found. Available: {sheet_names}")
        print(f"  Found TimelineData sheet")
        
        # Open PowerPoint
        print("Opening PowerPoint...")
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True  # PowerPoint must be visible to run VBA
        
        # Create new presentation
        presentation = ppt_app.Presentations.Add()
        print(f"  Created new presentation")
        
        # Import VBA module
        print("Importing VBA module...")
        if not os.path.exists(vba_file):
            raise FileNotFoundError(f"VBA file not found: {vba_file}")
        
        # Access VBA project (requires Trust access to VBA project object model)
        try:
            vba_project = presentation.VBProject
            vba_project.VBComponents.Import(os.path.abspath(vba_file))
            print(f"  Imported: {vba_file}")
        except Exception as e:
            print(f"  Warning: Could not import VBA module: {e}")
            print("  Make sure 'Trust access to VBA project object model' is enabled")
            print("  in File > Options > Trust Center > Trust Center Settings > Macro Settings")
            raise
        
        # Run the macro
        print("Running CreateTimelineFromData macro...")
        try:
            ppt_app.Run("CreateTimelineFromData")
            print("  Macro executed successfully")
        except Exception as e:
            print(f"  Error running macro: {e}")
            raise
        
        # Wait for macro to complete
        time.sleep(2)
        
        # Save the presentation
        print("Saving presentation...")
        output_dir = os.path.dirname(os.path.abspath(output_file))
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        # Save as .pptx format (24 = ppSaveAsOpenXMLPresentation)
        presentation.SaveAs(os.path.abspath(output_file), 24)
        print(f"  Saved to: {output_file}")
        
        # Get slide count
        slide_count = presentation.Slides.Count
        print(f"  Generated {slide_count} slide(s)")
        
        # Close presentation
        presentation.Close()
        workbook.Close(SaveChanges=False)
        
        print()
        print("Timeline generation completed successfully!")
        return True
        
    except Exception as e:
        print(f"\nError: {e}")
        return False
        
    finally:
        # Clean up
        if excel_app:
            try:
                excel_app.Quit()
            except Exception:
                pass
        if ppt_app:
            try:
                ppt_app.Quit()
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(
        description="Run VBA timeline generator macro via COM automation"
    )
    parser.add_argument(
        "--excel-file",
        default="timeline.xlsx",
        help="Path to Excel data file (default: timeline.xlsx)"
    )
    parser.add_argument(
        "--vba-file",
        default="timeline.bas",
        help="Path to VBA .bas file (default: timeline.bas)"
    )
    parser.add_argument(
        "--output",
        default="output/generated_timeline.pptx",
        help="Output path for generated presentation"
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Office applications during execution"
    )
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Only check if Office is installed"
    )
    
    args = parser.parse_args()
    
    # Check platform
    import platform
    if platform.system() != "Windows":
        print("Error: This script requires Windows with Microsoft Office installed")
        return 1
    
    # Check Office installation
    if not check_office_installed():
        print("\nMicrosoft Office (Excel and PowerPoint) is required")
        return 1
    
    if args.check_only:
        print("Office installation check passed!")
        return 0
    
    # Run the macro
    success = run_vba_macro(
        excel_file=args.excel_file,
        vba_file=args.vba_file,
        output_file=args.output,
        visible=args.visible
    )
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())

