from docxtpl import DocxTemplate

def get_valid_number(prompt):
    """
    A helper function that safely gets an integer from the user.
    If they type a string (like 'five'), it catches the error and asks again.
    """
    while True:
        user_input = input(prompt)
        try:
            # Try to convert the input into an integer
            number = int(user_input)
            if number < 0:
                print("⚠️ Please enter a positive number.")
                continue
            return number
        except ValueError:
            # If it fails, show this error and restart the loop
            print("❌ Invalid input! Please type a number (e.g., 1, 2, 3).")

def collect_hackathon_data():
    """
    Collects all team and member data and structures it for docxtpl.
    """
    teams_data = []
    
    print("\n=== Hackathon Report Generator ===")
    num_teams = get_valid_number("Enter the total number of teams: ")

    for i in range(num_teams):
        print(f"\n--- Team {i + 1} Details ---")
        
        # We leave these as strings so users can type "2A" or "Team Alpha"
        team_no = str(i+1)
        team_name = input("Enter Team Name: ").strip()
        ps_no = input("Enter Problem Statement No 🅿🆂: ").strip()
        
        members_list = []
        num_members = get_valid_number(f"How many members are in team '{team_name}'? ")
        
        for j in range(num_members):
            print(f"  [ Member {j + 1} ]")
            member_name = input("    Name: ").strip()
            # Using 'id' instead of 'roll_no' to keep the Word template tag short!
            member_id = input("    Roll No: ").strip() 
            
            members_list.append({
                "name": member_name,
                "id": member_id
            })
            
        teams_data.append({
            "team_no": team_no,
            "team_name": team_name,
            "ps_no": ps_no,
            "members": members_list,
            "length" : num_members
        })
        
    return {"teams": teams_data}

# ==========================================
# Main Execution Pipeline
# ==========================================
if __name__ == "__main__":

    doc = DocxTemplate("template.docx")

    # 1. Get the safe data from the user
    context = collect_hackathon_data()

    #collecting the front page data
    code = input("Enter the Hackathon Code: ")
    duration = input("Enter the Number of Hours: ") + "Hours"
    date = input("Enter Date: ")
    venue = input("Enter the Venue :")
    if venue == "":
        venue =  "ideaPad, Einstein Block, ITM Gwalior"
         
    event_details = {
        "code": code,
        "duration": duration,
        "date": date,
        "venue": venue
    }
    
#     # Merge the event details into the main context dictionary
    context.update(event_details)
    
    print("\n✅ Data collection complete. Generating Word Document...")
    
    # 2. Pass it directly into your template
    doc.render(context)
    doc.save("Final_Hackathon_Report.docx")


# Save docx first
doc.save("Final_Hackathon_Report.docx")


print("🎉 Success! 'Final_Hackathon_Report.docx' has been created.")