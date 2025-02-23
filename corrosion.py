import datetime
import pandas as pd

# Sample data structure
areas = {
    "1-HC-P": {"color": "Green", "message": "Good Condition"},
    "2-HC-P": {"color": "Green", "message": "Good Condition"},
    # Add all other areas...
}

# Function to update an area
def update_area(area, new_color, new_message):
    if area not in areas:
        print(f"Area {area} not found!")
        return
    
    previous_state = areas[area].copy()
    areas[area]["color"] = new_color
    areas[area]["message"] = new_message
    log_change(area, previous_state, new_color, new_message)

# Function to log changes
def log_change(area, previous_state, new_color, new_message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    report = {
        "Area": area,
        "Previous Color": previous_state["color"],
        "Previous Message": previous_state["message"],
        "New Color": new_color,
        "New Message": new_message,
        "Timestamp": timestamp
    }
    save_to_excel(report)  # Save to Excel or print

# Function to save changes to Excel
def save_to_excel(report):
    df = pd.DataFrame([report])
    try:
        # Load existing report file
        existing_df = pd.read_excel("ship_report.xlsx")
        updated_df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        # Create new file if it doesn't exist
        updated_df = df
    
    # Save to Excel
    updated_df.to_excel("ship_report.xlsx", index=False)
    print(f"Report saved for area {report['Area']}.")

# Example usage
update_area("1-HC-P", "Yellow", "Moderate Corrosion")
update_area("CH-1-MD-P", "Red", "Heavy Corrosion")
update_area("FWD", "Light Blue", "Maintenance Work in Progress")