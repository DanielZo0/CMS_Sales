#!/usr/bin/env python3
"""
CMZ Data Processing Suite - Main Entry Point

This is the main entry point for all CMZ data processing operations.
Provides a unified interface to extract, process, and combine sales and purchases data.
"""

import sys
import os
from pathlib import Path

def print_usage():
    """Print usage information"""
    print("""
CMZ Data Processing Suite
========================

Usage: python main.py [command]

Commands:
  gui         - Launch GUI interface for interactive processing
  extract     - Extract data from PDF and Excel files (default)
  combine     - Generate combined sales and purchases CSV
  help        - Show this help message

Examples:
  python main.py           # Extract all data (command line)
  python main.py gui       # Launch GUI interface
  python main.py extract   # Extract all data (command line)
  python main.py combine   # Generate combined CSV only
  python main.py help      # Show help
""")

def run_extraction():
    """Run the main data extraction process"""
    try:
        from combined_data_extractor import main as extract_main
        extract_main()
    except ImportError as e:
        print(f"Error importing extraction module: {e}")
        return False
    except Exception as e:
        print(f"Error during extraction: {e}")
        return False
    return True

def run_combine():
    """Run only the CSV combination process"""
    try:
        from generate_combined_csv import generate_combined_csv
        return generate_combined_csv()
    except ImportError as e:
        print(f"Error importing combine module: {e}")
        return False
    except Exception as e:
        print(f"Error during combination: {e}")
        return False

def run_gui():
    """Launch the GUI interface"""
    try:
        from gui_extractor import main as gui_main
        gui_main()
        return True
    except ImportError as e:
        print(f"Error importing GUI module: {e}")
        print("Make sure tkinter is installed: pip install tk")
        return False
    except Exception as e:
        print(f"Error launching GUI: {e}")
        return False

def main():
    """Main entry point"""
    # Get command line argument
    command = sys.argv[1] if len(sys.argv) > 1 else "extract"
    
    if command in ["help", "-h", "--help"]:
        print_usage()
        return
    
    print("CMZ Data Processing Suite")
    print("=" * 50)
    
    if command == "gui":
        print("Launching GUI interface...")
        success = run_gui()
        if not success:
            print("\nGUI launch failed!")
            sys.exit(1)
    
    elif command == "extract":
        print("Running full data extraction process...")
        success = run_extraction()
        if success:
            print("\nData extraction completed successfully!")
        else:
            print("\nData extraction failed!")
            sys.exit(1)
    
    elif command == "combine":
        print("Running CSV combination process...")
        success = run_combine()
        if success:
            print("\nCSV combination completed successfully!")
        else:
            print("\nCSV combination failed!")
            sys.exit(1)
    
    else:
        print(f"Unknown command: {command}")
        print_usage()
        sys.exit(1)

if __name__ == "__main__":
    main()
