"""
Extract relevant data from Daily Planning Template into a simple CSV file
"""
import openpyxl
import csv
from datetime import datetime

def extract_data_to_csv(template_path="Daily Planning Template.xlsm", output_path="daily_plan_data.csv"):
    """Extract orders data from template to CSV."""
    wb = openpyxl.load_workbook(template_path, data_only=True)
    main_sheet = wb['Main']
    
    # Extract limits (row 2)
    headers = [cell.value for cell in main_sheet[1]]
    limits_row = [cell.value for cell in main_sheet[2]]
    
    limits = {}
    try:
        qty_idx = headers.index('Qty')
        picks_idx = headers.index('Picks')
        hours_idx = headers.index('Hours')
        limits = {
            'Qty': limits_row[qty_idx],
            'Picks': limits_row[picks_idx],
            'Hours': limits_row[hours_idx]
        }
    except ValueError:
        print("Warning: Could not extract limits")
    
    # Extract order headers (row 11)
    order_headers = [cell.value for cell in main_sheet[11]]
    
    # Extract orders (starting row 12)
    orders = []
    for row_idx in range(12, main_sheet.max_row + 1):
        row = [cell.value for cell in main_sheet[row_idx]]
        if not row[0] or row[0] == 'Order No':
            continue
        
        order = {}
        for idx, header in enumerate(order_headers):
            if idx < len(row) and header:
                value = row[idx]
                # Convert datetime to string
                if isinstance(value, datetime):
                    value = value.strftime('%Y-%m-%d')
                order[header] = value
        
        if order.get('Order No') and order.get('Part No'):
            orders.append(order)
    
    # Write to CSV
    if orders:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=order_headers)
            writer.writeheader()
            writer.writerows(orders)
        
        print(f"Extracted {len(orders)} orders to {output_path}")
        print(f"Limits: {limits}")
        print(f"\nColumns: {order_headers}")
        print(f"\nFirst 3 orders:")
        for i, order in enumerate(orders[:3], 1):
            print(f"\nOrder {i}:")
            for key, value in list(order.items())[:10]:  # First 10 fields
                print(f"  {key}: {value}")
    
    # Extract brand-specific limits (BVI and Malosa)
    # Based on explore_limits.py output, BVI is in first section, Malosa starts around col 29
    bvi_limits = {}
    malosa_limits = {}
    
    # Find BVI section (starts around col 16-17 based on headers)
    try:
        # BVI limits are in the main section (cols 2-4 for Qty, Picks, Hours)
        bvi_limits = {
            'Qty': limits.get('Qty', 0),
            'Picks': limits.get('Picks', 0),
            'Hours': limits.get('Hours', 0),
            'Low Picks': limits_row[headers.index('Low Picks')] if 'Low Picks' in headers else None,
            'Big Picks': limits_row[headers.index('Big Picks')] if 'Big Picks' in headers else None,
            'Large Orders': limits_row[headers.index('Large Orders')] if 'Large Orders' in headers else None,
            'Offline Jobs': limits_row[headers.index('Offline Jobs')] if 'Offline Jobs' in headers else None,
        }
        
        # Find Malosa section - look for "Malosa" in headers
        if 'Malosa' in headers:
            malosa_start = headers.index('Malosa')
            # Malosa Qty, Picks, Hours should be in next few columns
            for i in range(malosa_start, len(headers)):
                if headers[i] == 'Qty' and i + 2 < len(limits_row):
                    malosa_limits = {
                        'Qty': limits_row[i] if limits_row[i] is not None else 0,
                        'Picks': limits_row[i+1] if limits_row[i+1] is not None else 0,
                        'Hours': limits_row[i+2] if limits_row[i+2] is not None else 0,
                    }
                    break
    except Exception as e:
        print(f"Warning: Could not extract brand-specific limits: {e}")
    
    # Also save limits to a separate file
    with open('daily_plan_limits.txt', 'w') as f:
        f.write(f"BVI Qty Limit: {bvi_limits.get('Qty', limits.get('Qty', 0))}\n")
        f.write(f"BVI Picks Limit: {bvi_limits.get('Picks', limits.get('Picks', 0))}\n")
        f.write(f"BVI Hours Limit: {bvi_limits.get('Hours', limits.get('Hours', 0))}\n")
        if bvi_limits.get('Low Picks'):
            f.write(f"BVI Low Picks Limit: {bvi_limits['Low Picks']}\n")
        if bvi_limits.get('Big Picks'):
            f.write(f"BVI Big Picks Limit: {bvi_limits['Big Picks']}\n")
        if bvi_limits.get('Large Orders'):
            f.write(f"BVI Large Orders Limit: {bvi_limits['Large Orders']}\n")
        if bvi_limits.get('Offline Jobs'):
            f.write(f"BVI Offline Jobs Limit: {bvi_limits['Offline Jobs']}\n")
        
        f.write(f"\nMalosa Qty Limit: {malosa_limits.get('Qty', 0)}\n")
        f.write(f"Malosa Picks Limit: {malosa_limits.get('Picks', 0)}\n")
        f.write(f"Malosa Hours Limit: {malosa_limits.get('Hours', 0)}\n")
        
        f.write(f"\nNo Duplicate Parts\n")
        f.write(f"Limit Low Picks Orders\n")
        f.write(f"Limit High Picks Orders\n")
        f.write(f"Limit High Qty Orders\n")
        f.write(f"Limit Low Qty Orders\n")
        f.write(f"Limit High Hours Orders\n")
        f.write(f"Limit Low Hours orders\n")
    
    print(f"\nLimits saved to daily_plan_limits.txt")
    print(f"BVI Limits: {bvi_limits}")
    print(f"Malosa Limits: {malosa_limits}")
    
    return orders, limits

if __name__ == "__main__":
    extract_data_to_csv()
