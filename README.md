# EVERYTHING-DATA

def create_construction_plan(file_name="Construction_Action_Plan.xlsx"):
    # Define project tasks
    data = {
        "Task ID": list(range(1, 13)),
        "Task Name": [
            "Project Planning", "Site Survey & Soil Testing", "Permits & Approvals", "Site Preparation & Excavation",
            "Foundation Work", "Structural Framing", "Roofing Installation", "Plumbing & Electrical Installation",
            "HVAC & Insulation", "Interior & Exterior Finishing", "Final Inspection & Quality Check", "Project Handover & Documentation"
        ],
        "Description": [
            "Define project scope, budget, and timeline.", "Conduct site analysis and soil testing.",
            "Obtain necessary permits and approvals from authorities.", "Clear the site, level ground, and excavate as required.",
            "Lay foundation (footings, slab, reinforcement, curing).", "Construct columns, beams, and floors.",
            "Install roof structure and waterproofing.", "Install plumbing pipes, electrical wiring, and fixtures.",
            "Set up heating, ventilation, and insulation systems.", "Apply flooring, painting, doors, windows, and exterior finishes.",
            "Conduct inspections, address defects, and ensure compliance.", "Complete documentation, client handover, and final reporting."
        ],
        "Responsible Party": [
            "Project Manager", "Geotechnical Engineer", "Legal Consultant", "Site Engineer", "Civil Engineer",
            "Structural Engineer", "Roofing Contractor", "MEP Contractor", "HVAC Specialist", "Interior Designer",
            "Quality Inspector", "Project Manager"
        ],
        "Start Date": [
            "2025-03-01", "2025-03-05", "2025-03-10", "2025-03-15", "2025-03-25", "2025-04-05", "2025-04-20",
            "2025-05-01", "2025-05-10", "2025-05-20", "2025-06-01", "2025-06-10"
        ],
        "End Date": [
            "2025-03-04", "2025-03-09", "2025-03-14", "2025-03-24", "2025-04-04", "2025-04-19", "2025-04-30",
            "2025-05-09", "2025-05-19", "2025-05-31", "2025-06-09", "2025-06-15"
        ],
        "Status": ["Not Started"] * 12,
        "Remarks": [""] * 12
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save DataFrame to an Excel file
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Action Plan')
        
        # Access workbook and worksheet for formatting
        workbook = writer.book
        worksheet = writer.sheets['Action Plan']
        
        # Format header row
        for col in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Adjust column widths
        for col in worksheet.columns:
            max_length = max((len(str(cell.value)) for cell in col), default=10)
            worksheet.column_dimensions[col[0].column_letter].width = max_length + 2
    
    print(f"Excel file '{file_name}' created successfully.")

