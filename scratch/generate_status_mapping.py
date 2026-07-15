import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Define the combined dataset with Status Codes and Responsible Store mappings
data = [
    # Status Code | English Status | Khmer Operation | Meaning / Real-World Action | Responsible Store
    ("100 / -99", "Registered / Unscanned", "បានចុះឈ្មោះ / មិនទាន់ស្កេន", 
     "Order details have been created in the system, but the package has not yet been physically picked up or scanned.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("110", "Confirmed", "ភ្ញៀវហាង បង្កើតបុង", 
     "The customer/shop creates the booking or invoice in the system.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("120", "Pickup in progress", "កំពុងប្រមូលអីវ៉ាន់", 
     "The courier is currently picking up the parcel from the sender's location.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("200", "Received", "ហាងទទួលអីវ៉ាន់", 
     "The origin store/hub physically receives the package from the sender.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("311", "Assign order to sack", "ហាងទី១ បែងចែក ដឹកបន្តទៅហាងទទួល", 
     "The origin store (Store 1) sorts and assigns the package into a shipping sack/transit bag to go to the next store.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("302 (Outbound)", "Sack completed", "ហាងទី១ បែងចែកជោគជ័យ", 
     "The packaging of the transit sack at the origin store is completed successfully.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("210 / 300", "In Transit", "កំពុងដឹកជញ្ជូនលើផ្លូវ", 
     "The package is loaded onto a truck/vehicle and is traveling between offices or hubs.", 
     "ឡានដឹក (Transit Vehicle)"),
     
    ("306 (Hub)", "Hub Transit", "ដល់មជ្ឈមណ្ឌលចែកចាយ (Hub)", 
     "The package has arrived at the central sorting hub.", 
     "មជ្ឈមណ្ឌល (Hub / Transit Center)"),
     
    ("302 (Inbound) / 310", "Accept handover", "ហាងទី ២ រឺហាងទទួល បានទទួលអីវ៉ាន់", 
     "The receiving/destination store (Store 2) receives the handover of the sack/parcel from the truck driver.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("309 / 306", "Exploited-Forwarding", "ស្កេនបុងក្នុងចង់អីវ៉ាន់ជោគជ័យ", 
     "The individual invoice/barcode of the package inside the sack is scanned successfully (unpacked).", 
     "ហាងទទួល (Receiving Store)"),
     
    ("400", "Assignment confirmation", "ស្កេនបុងម្តងទៀត និងប្រគល់អោយភ្ញៀវ ផ្ទាល់ រឺម៉ូតូ", 
     "The barcode is scanned again to assign it for final delivery (to the motorcycle rider) or hand it directly to the customer.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("401 / 402", "Delivering", "បុគ្គលិកកំពុងដឹកជញ្ជូនជូនភ្ញៀវ", 
     "The delivery rider is actively on the road delivering the package to the receiver.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("420", "At Store for Pickup", "អីវ៉ាន់រង់ចាំភ្ញៀវមកយកនៅហាង", 
     "The package is kept at the branch waiting for the customer to pick it up themselves.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("410", "Shipped / Delivered", "ហាងទី២ បានទទួល / ប្រគល់ជោគជ័យ", 
     "Store 2 has successfully received the shipment or successfully delivered to the customer.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("430 / 460", "Redelivery", "រៀបចំដឹកជញ្ជូនម្តងទៀត", 
     "The first delivery attempt failed, so a second delivery attempt is scheduled.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("472", "Stored / Hold", "រក្សាទុក / ផ្អាកបណ្តោះអាសន្ន", 
     "The parcel is put on hold in the store's warehouse (held due to incorrect phone or receiver request).", 
     "ហាងទទួល (Receiving Store)"),
     
    ("480", "Change Address", "កែប្រែអាសយដ្ឋានថ្មី", 
     "The receiver requested a different delivery location, and the address is updated in the system.", 
     "ហាងទទួល (Receiving Store)"),
     
    ("201", "Cancelled / Closed", "បានលុបចោលបុង", 
     "The order was cancelled by the sender or closed by the system.", 
     "ហាងផ្ញើ (Sending Store)"),
     
    ("500 / 520 / 540", "Return Processing / Returned", "ផ្ញើត្រឡប់ទៅអ្នកផ្ញើវិញ", 
     "The delivery failed, and the package is returned (or being returned) back to the sender.", 
     "ហាងផ្ញើ (Sending Store)")
]

# Create workbook and sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Status Mapping"

# Enable grid lines
ws.views.sheetView[0].showGridLines = True

# Add headers
headers = ["Status Code", "English Status", "Khmer Operation", "Meaning / Real-World Action", "Responsible Store (ហាងទទួល/ផ្ញើ)"]
ws.append(headers)

# Add data rows
for row in data:
    ws.append(row)

# Style variables
font_family = "Segoe UI"
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Elegant Navy
header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
data_font = Font(name=font_family, size=11)
khmer_font = Font(name="Khmer OS Battambang", size=11)

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# Format headers
for col_idx in range(1, 6):
    cell = ws.cell(row=1, column=col_idx)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center" if col_idx == 1 else "left", vertical="center", wrap_text=True)
ws.row_dimensions[1].height = 30

# Format data rows
for r_idx in range(2, len(data) + 2):
    ws.row_dimensions[r_idx].height = 24
    
    # Status Code
    cell_code = ws.cell(row=r_idx, column=1)
    cell_code.font = Font(name=font_family, size=11, bold=True)
    cell_code.border = thin_border
    cell_code.alignment = Alignment(horizontal="center", vertical="center")
    
    # English Status
    cell_en = ws.cell(row=r_idx, column=2)
    cell_en.font = data_font
    cell_en.border = thin_border
    cell_en.alignment = Alignment(vertical="center")
    
    # Khmer Operation
    cell_kh = ws.cell(row=r_idx, column=3)
    cell_kh.font = khmer_font
    cell_kh.border = thin_border
    cell_kh.alignment = Alignment(vertical="center")
    
    # Meaning
    cell_desc = ws.cell(row=r_idx, column=4)
    cell_desc.font = data_font
    cell_desc.border = thin_border
    cell_desc.alignment = Alignment(vertical="center")
    
    # Responsible Store
    cell_store = ws.cell(row=r_idx, column=5)
    cell_store.font = khmer_font
    cell_store.border = thin_border
    cell_store.alignment = Alignment(vertical="center")

# Set manual column widths
ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 30
ws.column_dimensions['C'].width = 45
ws.column_dimensions['D'].width = 75
ws.column_dimensions['E'].width = 35

# Save the workbook
output_path = "c:/Users/DELL/Desktop/daily_push/status_mapping_v2.xlsx"
wb.save(output_path)
print(f"Excel successfully created/updated at: {output_path}")
