# **Excel VBA Automation – Data Sync & Conditional Formatting**  

## **Overview**  
This VBA script automates data synchronization and conditional formatting between two sheets: **"EFC Master"** and **"Device Information."** It ensures efficient data transfer and enforces structured user interactions.  

## **Features**  

- **Auto Data Sync:** Selecting any cell in **B2:Q14** on **"Device Information"** copies values and formatting (color, font, bold) to **"EFC Master."**  
- **Conditional Formatting:**  
  - If **Q9** changes to `"No"`, it turns **red** and prompts the user about the **Go-Back** sheet.  
  - If `"No"` is selected, a hidden hyperlink (Q273) automatically opens the **Go-Back** sheet.  
  - If `"Yes"`, Q9 turns **green**.  
  - Q9 is always set to **"Non-Managed CC Switch Resolved."**  
- **Optimized Performance:** Uses **screen updating and event disabling** to prevent unnecessary recalculations.  

## **How to Use**  

1. **Enable Macros:** Ensure macros are enabled in Excel (`File > Options > Trust Center > Macro Settings`).  
2. **Selection-Based Sync:** Click on any cell in **B2:Q14** on **"Device Information"** to transfer data to **"EFC Master."**  
3. **Monitor Q9:** If Q9 changes, follow prompts for the **Go-Back** sheet when needed.  

## **Customization**  

- Modify **B2:Q14** if different data ranges need syncing.  
- Change **Q9** if a different cell should control formatting.  
- Adjust **Q273** if the Go-Back hyperlink is stored elsewhere.  

## **Notes**  

- Ensure the **Go-Back sheet hyperlink exists** in **Q273** or update the reference.  
- Undo (`CTRL + Z`) won’t work, as VBA directly modifies cells.  
- If macros don’t run, check **macro security settings** and enable VBA.  

## **License**  

This script is provided as-is. Modify and use it as needed.
