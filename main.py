"""
Updated Main Processing Code
Now uses autosaver.py module for reliable Book1 capture
"""

import os
import time
import pandas as pd
from autosaver import capture_book1, is_book1_available

# Configuration
SAVE_FOLDER = r"C:\Users\sasuk\Documents\CapturedExports"
PROCESSED_FOLDER = r"C:\Users\sasuk\Documents\ProcessedExports"

# Ensure directories exist
os.makedirs(SAVE_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def transform_excel_file(filepath):
    """
    Transform captured Excel file into processed reports.
    (Your existing processing logic - unchanged)
    """
    try:
        captured_df = pd.read_excel(filepath, dtype={"Item ID": str})
        brand_map_df = pd.read_csv("brand_map.csv", dtype={"Item ID": str})
        df = captured_df.merge(brand_map_df, on="Item ID", how="left")
        df_by_cat = df.dropna(subset=["CATEGORY"]).copy()
        df_by_cat["Brand : Category"] = df["Brand"].astype(str) + " : " + df["CATEGORY"].astype(str)
        df_by_cat = df_by_cat.sort_values(by="Item ID")

        # Create processed reports
        grouped_df_id = calc_profit_percentage_accname(df, 0)
        grouped_df_br = calc_profit_percentage_brand(df, 0)
        grouped_df_brcat = calc_profit_percentage_brand(df_by_cat, 1)

        # Generate output filenames
        processed_filename_id = "processed_ver-id_" + os.path.basename(filepath) 
        processed_filename_br = "processed_ver-br_" + os.path.basename(filepath) 
        processed_filename_brcat = "processed_ver-brcat_" + os.path.basename(filepath) 
        
        processed_path_id = os.path.join(PROCESSED_FOLDER, processed_filename_id)
        processed_path_br = os.path.join(PROCESSED_FOLDER, processed_filename_br)
        processed_path_brcat = os.path.join(PROCESSED_FOLDER, processed_filename_brcat)
        
        # Save processed files
        grouped_df_id.to_excel(processed_path_id, index=False)
        grouped_df_br.to_excel(processed_path_br, index=False)
        grouped_df_brcat.to_excel(processed_path_brcat, index=False)

        print(f"✅ Transformed and saved:")
        print(f"   📊 ID Report: {os.path.basename(processed_path_id)}")
        print(f"   📊 Brand Report: {os.path.basename(processed_path_br)}")
        print(f"   📊 Brand-Category Report: {os.path.basename(processed_path_brcat)}")
        
        # Open processed files
        os.startfile(processed_path_id)
        os.startfile(processed_path_br)
        os.startfile(processed_path_brcat)

        return True
        
    except Exception as e:
        print(f"⚠️ Error during processing: {e}")
        return False

def calc_profit_percentage_accname(df, vernum):
    """Calculate profit percentage by account name."""
    if vernum == 0:
        grouped_df = df.groupby("Account Name", as_index=False).agg({
            "Sale Price": "sum",
            "Unit Cost": "sum",
        })
        grouped_df["Profit %"] = ((grouped_df["Sale Price"] - grouped_df["Unit Cost"]) / grouped_df["Sale Price"]) * 100
        grouped_df["Profit %"] = grouped_df["Profit %"].round(2).astype(str) + "%"
        grouped_df.rename(columns={"Sale Price": "Agg Sale Price", "Unit Cost": "Agg Unit Cost"}, inplace=True)
        return grouped_df

def calc_profit_percentage_brand(df, vernum):
    """Calculate profit percentage by brand or brand-category."""
    if vernum == 0:
        grouped_df = df.groupby("Brand", as_index=False).agg({
            "Sale Price": "sum",
            "Unit Cost": "sum",
        })
        grouped_df["Profit %"] = ((grouped_df["Sale Price"] - grouped_df["Unit Cost"]) / grouped_df["Sale Price"]) * 100
        grouped_df["Profit %"] = grouped_df["Profit %"].round(2).astype(str) + "%"
        grouped_df.rename(columns={"Sale Price": "Agg Sale Price", "Unit Cost": "Agg Unit Cost"}, inplace=True)
        return grouped_df
    
    if vernum == 1:
        grouped_df = df.groupby("Brand : Category", as_index=False).agg({
            "Sale Price": "sum",
            "Unit Cost": "sum",
        })
        grouped_df["Profit %"] = ((grouped_df["Sale Price"] - grouped_df["Unit Cost"]) / grouped_df["Sale Price"]) * 100
        grouped_df["Profit %"] = grouped_df["Profit %"].round(2).astype(str) + "%"
        grouped_df.rename(columns={"Sale Price": "Agg Sale Price", "Unit Cost": "Agg Unit Cost"}, inplace=True)
        return grouped_df

def auto_capture_and_transform():
    """
    Main automation loop - continuously monitor for Book1 and process it.
    Now uses reliable autosaver.py module instead of problematic COM approach.
    """
    print("🚀 Excel Automation with Reliable Auto-Saver")
    print("👀 Monitoring for Book1 exports...")
    print("   (Press Ctrl+C to stop)")
    
    last_check_failed = False
    
    while True:
        try:
            # Check if Book1 is available using autosaver module
            if is_book1_available():
                print("\n📄 Book1 detected! Starting capture...")
                
                # Capture using autosaver module
                saved_file = capture_book1(SAVE_FOLDER, verbose=True)
                
                if saved_file:
                    print(f"📁 File captured: {os.path.basename(saved_file)}")
                    
                    # Process the captured file
                    print("🔄 Starting data transformation...")
                    success = transform_excel_file(saved_file)
                    
                    if success:
                        print("✅ Processing completed successfully!")
                        print("📊 Processed reports opened automatically")
                        
                        # Optional: Clean up captured file after processing
                        # os.remove(saved_file)
                        # print(f"🗑️ Cleaned up captured file")
                        
                    else:
                        print("❌ Processing failed - check error messages above")
                    
                    print("\n" + "─" * 50)
                    print("👀 Monitoring for next Book1 export...")
                    
                else:
                    print("❌ Failed to capture Book1 - will retry in 5 seconds")
                
                last_check_failed = False
                
            else:
                # Only print "waiting" message occasionally to avoid spam
                if not last_check_failed:
                    print("⏳ No Book1 detected, waiting for export...")
                last_check_failed = True
            
            # Wait before next check
            time.sleep(5)  # Check every 5 seconds
            
        except KeyboardInterrupt:
            print("\n🛑 Automation stopped by user")
            break
            
        except Exception as e:
            print(f"⚠️ Unexpected error in automation loop: {e}")
            print("🔄 Continuing monitoring...")
            time.sleep(5)  # Wait longer after errors

def capture_once():
    """
    One-time capture and processing (alternative to continuous monitoring).
    Perfect for manual triggers or GUI integration.
    """
    print("🔍 Looking for Book1 to capture...")
    
    # Use autosaver module for capture
    saved_file = capture_book1(SAVE_FOLDER, verbose=True)
    
    if saved_file:
        print("🔄 Processing captured file...")
        success = transform_excel_file(saved_file)
        
        if success:
            print("✅ Capture and processing completed!")
            return saved_file
        else:
            print("❌ Processing failed")
            return None
    else:
        print("❌ No Book1 available for capture")
        return None

if __name__ == "__main__":
    # Choose your preferred mode:
    
    # Option 1: Continuous monitoring (your original behavior)
    auto_capture_and_transform()
    
    # Option 2: One-time capture (uncomment to use instead)
    # capture_once()