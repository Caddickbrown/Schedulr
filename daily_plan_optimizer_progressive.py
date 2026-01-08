"""
Daily Planning Optimizer - Progressive Fill
Fills each day to 100% hours before moving to the next day.
Balances lines proportionally within each day.
"""
import csv
import openpyxl
import os
from datetime import datetime
from typing import List, Dict, Tuple
import json


class DailyPlanOptimizerProgressive:
    def __init__(self, template_path: str = None):
        """
        Initialize optimizer with Excel template file.
        
        Args:
            template_path: Path to Excel template file (default: 'Daily Planning Template.xlsm')
        """
        self.template_path = template_path or 'Daily Planning Template.xlsm'
        self.limits = {}
        self.orders = []
        self.brand_limits = {}
        
    def load_data(self):
        """Load orders and limits directly from Excel template."""
        self._load_limits()
        self._load_orders()
        
    def _load_limits(self):
        """Load brand-specific limits directly from Excel template."""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template file not found: {self.template_path}")
        
        wb = openpyxl.load_workbook(self.template_path, data_only=True)
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
                'Qty': self._parse_float(limits_row[qty_idx]),
                'Picks': self._parse_float(limits_row[picks_idx]),
                'Hours': self._parse_float(limits_row[hours_idx])
            }
        except ValueError:
            print("Warning: Could not extract basic limits")
        
        # Extract brand-specific limits (BVI and Malosa)
        bvi_limits = {}
        malosa_limits = {}
        
        try:
            bvi_limits = {
                'Qty': limits.get('Qty', 0),
                'Picks': limits.get('Picks', 0),
                'Hours': limits.get('Hours', 0),
            }
            
            if 'Low Picks' in headers:
                try:
                    bvi_limits['Low Picks'] = self._parse_float(limits_row[headers.index('Low Picks')])
                except (ValueError, IndexError):
                    pass
            if 'Big Picks' in headers:
                try:
                    bvi_limits['Big Picks'] = self._parse_float(limits_row[headers.index('Big Picks')])
                except (ValueError, IndexError):
                    pass
            if 'Large Orders' in headers:
                try:
                    bvi_limits['Large Orders'] = self._parse_float(limits_row[headers.index('Large Orders')])
                except (ValueError, IndexError):
                    pass
            if 'Offline Jobs' in headers:
                try:
                    bvi_limits['Offline Jobs'] = self._parse_float(limits_row[headers.index('Offline Jobs')])
                except (ValueError, IndexError):
                    pass
            
            if 'Malosa' in headers:
                malosa_start = headers.index('Malosa')
                for i in range(malosa_start, len(headers)):
                    if headers[i] == 'Qty' and i + 2 < len(limits_row):
                        malosa_limits = {
                            'Qty': self._parse_float(limits_row[i]),
                            'Picks': self._parse_float(limits_row[i+1]),
                            'Hours': self._parse_float(limits_row[i+2]),
                        }
                        break
        except Exception as e:
            print(f"Warning: Could not extract brand-specific limits: {e}")
        
        if not bvi_limits.get('Qty'):
            bvi_limits = {'Qty': 10544, 'Picks': 750, 'Hours': 390}
        if not malosa_limits.get('Qty'):
            malosa_limits = {'Qty': 3335, 'Picks': 130, 'Hours': 90}
        
        self.brand_limits = {
            'BVI': bvi_limits,
            'Malosa': malosa_limits
        }
        
        print(f"Loaded brand limits: {self.brand_limits}")
        self.limits = self.brand_limits.get('BVI', {})
    
    def _load_orders(self):
        """Load orders directly from Excel template."""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template file not found: {self.template_path}")
        
        wb = openpyxl.load_workbook(self.template_path, data_only=True)
        main_sheet = wb['Main']
        
        order_headers = [cell.value for cell in main_sheet[11]]
        
        for row_idx in range(12, main_sheet.max_row + 1):
            row = [cell.value for cell in main_sheet[row_idx]]
            if not row[0] or row[0] == 'Order No':
                continue
            
            try:
                row_dict = {}
                for idx, header in enumerate(order_headers):
                    if idx < len(row) and header:
                        value = row[idx]
                        row_dict[header] = value
                
                suggested_line = str(row_dict.get('Suggested Line', '')).strip()
                if suggested_line in ['C3/4', 'C3&4']:
                    suggested_line = 'C3/4'
                
                qty = self._parse_float(row_dict.get('Lot Size', 0))
                picks = self._parse_float(row_dict.get('Picks', 0))
                hours = self._parse_float(row_dict.get('Hours', 0))
                
                start_date = row_dict.get('Start Date')
                if isinstance(start_date, datetime):
                    parsed_date = start_date
                elif start_date:
                    parsed_date = self._parse_date(str(start_date))
                else:
                    parsed_date = None
                
                order = {
                    'Order No': str(row_dict.get('Order No', '')).strip(),
                    'Part No': str(row_dict.get('Part No', '')).strip(),
                    'Brand': str(row_dict.get('Brand', '')).strip(),
                    'Start Date': parsed_date,
                    'Lot Size': qty,
                    'Picks': picks,
                    'Hours': hours,
                    'Country': str(row_dict.get('Country', '')).strip(),
                    'Wrap Type': str(row_dict.get('Wrap Type', '')).strip(),
                    'CPU': self._parse_float(row_dict.get('CPU', 0)),
                    'Suggested Line': suggested_line,
                }
                
                if order['Order No'] and order['Part No'] and order['Lot Size'] > 0:
                    self.orders.append(order)
            except Exception as e:
                order_no = row[0] if row else 'unknown'
                print(f"Error processing order {order_no}: {e}")
                continue
        
        print(f"Loaded {len(self.orders)} orders")
    
    def _parse_date(self, date_str):
        """Parse date string to datetime object."""
        if not date_str or date_str.strip() == '':
            return None
        try:
            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d %H:%M:%S']:
                try:
                    return datetime.strptime(date_str.strip(), fmt)
                except ValueError:
                    continue
            return None
        except:
            return None
    
    def _parse_float(self, value):
        """Parse float value, handling empty strings and None."""
        if value is None or value == '' or value == '-':
            return 0.0
        try:
            return float(str(value).replace(',', ''))
        except:
            return 0.0
    
    def _get_line_category(self, line: str) -> str:
        """Normalize line name to category (C1, C2, C3/4, or Other)."""
        if not line:
            return 'Other'
        line_upper = line.upper().strip()
        if line_upper in ['C1']:
            return 'C1'
        elif line_upper in ['C2']:
            return 'C2'
        elif line_upper in ['C3/4', 'C3&4', 'C3/4']:
            return 'C3/4'
        else:
            return 'Other'
    
    def generate_multi_day_plans(self, num_days: int, brand: str = None) -> List[Dict]:
        """
        Generate multi-day plans with 100% hours utilization per day.
        
        PROGRESSIVE FILL APPROACH:
        1. Fill each day to EXACTLY 100% hours before moving to next day
        2. Balance lines proportionally within each day
        3. Prioritize earlier start dates
        4. Last day gets remainder (may be < 100%)
        
        Args:
            num_days: Maximum number of days to plan
            brand: Brand to plan (BVI, Malosa, etc.)
        
        Returns:
            List of day plans with 'day', 'orders', 'totals', 'utilization', 'num_orders'
        """
        # Get limits for brand
        if brand:
            limits = self.brand_limits.get(brand, self.limits)
            brand_orders = [o for o in self.orders if o.get('Brand', '').upper() == brand.upper()]
        else:
            limits = self.limits
            brand = 'BVI'
            brand_orders = [o for o in self.orders if o.get('Brand', '').upper() == 'BVI']
        
        if not brand_orders:
            return []
        
        hours_limit = limits['Hours']
        picks_limit = limits['Picks']
        qty_limit = limits['Qty']
        offline_limit = limits.get('Offline Jobs', float('inf'))
        
        # Prepare orders with metrics
        orders_with_metrics = []
        for order in brand_orders:
            qty = order.get('Lot Size', 0) or 0
            picks = order.get('Picks', 0) or 0
            hours = order.get('Hours', 0) or 0
            
            if qty > 0 and hours > 0:
                orders_with_metrics.append({
                    'order': order,
                    'qty': qty,
                    'picks': picks,
                    'hours': hours,
                    'start_date': order.get('Start Date') or datetime.max,
                    'line': self._get_line_category(order.get('Suggested Line', ''))
                })
        
        # Calculate totals
        total_hours = sum(item['hours'] for item in orders_with_metrics)
        total_picks = sum(item['picks'] for item in orders_with_metrics)
        total_qty = sum(item['qty'] for item in orders_with_metrics)
        total_orders = len(orders_with_metrics)
        
        # Calculate line distribution in source data
        line_counts_source = {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0}
        for item in orders_with_metrics:
            line_counts_source[item['line']] += 1
        
        total_c1_c2_c34 = line_counts_source['C1'] + line_counts_source['C2'] + line_counts_source['C3/4']
        if total_c1_c2_c34 > 0:
            line_ratios = {
                'C1': line_counts_source['C1'] / total_c1_c2_c34,
                'C2': line_counts_source['C2'] / total_c1_c2_c34,
                'C3/4': line_counts_source['C3/4'] / total_c1_c2_c34,
            }
        else:
            line_ratios = {'C1': 0.33, 'C2': 0.33, 'C3/4': 0.34}
        
        # Calculate actual number of full days possible
        full_days_possible = int(total_hours / hours_limit)
        remainder_hours = total_hours - (full_days_possible * hours_limit)
        actual_days = full_days_possible + (1 if remainder_hours > 0 else 0)
        num_days = min(num_days, actual_days)
        
        print(f"\n{'='*60}")
        print(f"PROGRESSIVE FILL PLANNING: {brand}")
        print(f"{'='*60}")
        print(f"Total orders: {total_orders}")
        print(f"Total hours: {total_hours:.1f}")
        print(f"Hours limit per day: {hours_limit}")
        print(f"Full days possible: {full_days_possible} at 100%")
        print(f"Remainder: {remainder_hours:.1f} hours ({remainder_hours/hours_limit*100:.1f}%)")
        print(f"\nSource line distribution:")
        print(f"  C1: {line_counts_source['C1']} ({line_ratios['C1']*100:.1f}%)")
        print(f"  C2: {line_counts_source['C2']} ({line_ratios['C2']*100:.1f}%)")
        print(f"  C3/4: {line_counts_source['C3/4']} ({line_ratios['C3/4']*100:.1f}%)")
        print(f"  Other: {line_counts_source['Other']}")
        
        # Sort by start date (earlier first), then by hours (larger first for better packing)
        orders_with_metrics.sort(key=lambda x: (x['start_date'], -x['hours']))
        
        remaining_orders = orders_with_metrics.copy()
        days = []
        
        # Fill each day to 100% before moving to next
        for day_num in range(1, num_days + 1):
            if not remaining_orders:
                break
            
            # Determine target for this day
            remaining_hours_total = sum(item['hours'] for item in remaining_orders)
            
            # Target 100% if there's enough hours, otherwise take what's available
            if remaining_hours_total >= hours_limit:
                target_hours = hours_limit
            else:
                target_hours = remaining_hours_total
            
            day = self._fill_day_progressive(
                day_num, remaining_orders, target_hours, hours_limit,
                line_ratios, offline_limit, brand, is_last_day=False
            )
            
            # Remove selected orders from remaining
            selected_order_nos = {o['Order No'] for o in day['orders']}
            remaining_orders = [item for item in remaining_orders 
                              if item['order']['Order No'] not in selected_order_nos]
            
            # Finalize day
            day['utilization'] = {
                'Qty': day['totals']['Qty'] / qty_limit * 100 if qty_limit > 0 else 0,
                'Picks': day['totals']['Picks'] / picks_limit * 100 if picks_limit > 0 else 0,
                'Hours': day['totals']['Hours'] / hours_limit * 100 if hours_limit > 0 else 0
            }
            day['brand'] = brand
            day['day_label'] = f"Day {day_num}"
            day['offline_limit'] = offline_limit
            
            print(f"\n--- Day {day_num} ---")
            print(f"  Orders: {day['num_orders']}")
            print(f"  Hours: {day['totals']['Hours']:.1f} ({day['utilization']['Hours']:.1f}%)")
            print(f"  Lines: C1={day['line_counts']['C1']}, C2={day['line_counts']['C2']}, "
                  f"C3/4={day['line_counts']['C3/4']}, Other={day['line_counts']['Other']}")
            
            if day['num_orders'] > 0:
                days.append(day)
        
        # If there are remaining orders, add them to the last day
        if remaining_orders and days:
            last_day = days[-1]
            print(f"\n  Adding {len(remaining_orders)} remaining orders to Day {last_day['day']}...")
            for item in remaining_orders:
                self._add_order_to_day(last_day, item)
            
            # Update last day's metrics
            last_day['utilization'] = {
                'Qty': last_day['totals']['Qty'] / qty_limit * 100 if qty_limit > 0 else 0,
                'Picks': last_day['totals']['Picks'] / picks_limit * 100 if picks_limit > 0 else 0,
                'Hours': last_day['totals']['Hours'] / hours_limit * 100 if hours_limit > 0 else 0
            }
            last_day['line_distribution'] = {
                'C1': {'count': last_day['line_counts']['C1'], 'hours': last_day['line_hours']['C1']},
                'C2': {'count': last_day['line_counts']['C2'], 'hours': last_day['line_hours']['C2']},
                'C3/4': {'count': last_day['line_counts']['C3/4'], 'hours': last_day['line_hours']['C3/4']},
                'Other': {'count': last_day['line_counts']['Other'], 'hours': last_day['line_hours']['Other']}
            }
            print(f"  Day {last_day['day']} final: {last_day['num_orders']} orders, "
                  f"{last_day['totals']['Hours']:.1f} hours ({last_day['utilization']['Hours']:.1f}%)")
        
        return days
    
    def _fill_day_progressive(self, day_num: int, available_orders: List[Dict], 
                              target_hours: float, hours_limit: float,
                              line_ratios: Dict, offline_limit: float, brand: str,
                              is_last_day: bool = False) -> Dict:
        """
        Fill a single day to target hours with proportional line balance.
        
        Strategy:
        1. Calculate how many orders from each line we want (proportional)
        2. Greedily select orders to hit hours target while respecting proportions
        3. Prioritize earlier start dates
        """
        day = {
            'day': day_num,
            'orders': [],
            'totals': {'Qty': 0, 'Picks': 0, 'Hours': 0},
            'num_orders': 0,
            'line_counts': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
            'line_hours': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
            'offline_count': 0,
        }
        
        # Group available orders by line
        orders_by_line = {'C1': [], 'C2': [], 'C3/4': [], 'Other': []}
        for item in available_orders:
            orders_by_line[item['line']].append(item)
        
        # Calculate target counts per line based on proportions and expected order count
        # Estimate order count based on average hours per order
        avg_hours_per_order = sum(item['hours'] for item in available_orders) / len(available_orders) if available_orders else 10
        estimated_orders = int(target_hours / avg_hours_per_order)
        
        # Calculate target line counts (proportional)
        c1_c2_c34_target = int(estimated_orders * 0.7)  # ~70% from C1/C2/C3/4
        target_line_counts = {
            'C1': int(c1_c2_c34_target * line_ratios['C1']),
            'C2': int(c1_c2_c34_target * line_ratios['C2']),
            'C3/4': int(c1_c2_c34_target * line_ratios['C3/4']),
        }
        
        # Phase 1: Ensure at least one from each target line (for balance)
        target_lines = ['C1', 'C2', 'C3/4']
        for line in target_lines:
            if orders_by_line[line] and day['line_counts'][line] == 0:
                # Pick earliest-dated order from this line that fits
                for item in orders_by_line[line]:
                    if day['totals']['Hours'] + item['hours'] <= target_hours * 1.1:
                        self._add_order_to_day(day, item)
                        orders_by_line[line].remove(item)
                        break
        
        # Phase 2: Fill to target hours using proportional selection
        max_iterations = len(available_orders) * 2
        iteration = 0
        
        while iteration < max_iterations:
            iteration += 1
            
            current_hours = day['totals']['Hours']
            hours_remaining = target_hours - current_hours
            
            # Stop if we've hit target (within 1%)
            if current_hours >= target_hours * 0.99:
                break
            
            # Stop if we're close enough and can't find good fits
            if current_hours >= target_hours * 0.95 and hours_remaining < 5:
                break
            
            # Find best order to add
            best_order = None
            best_score = -float('inf')
            best_line = None
            
            for line in ['C1', 'C2', 'C3/4', 'Other']:
                for item in orders_by_line[line]:
                    hours = item['hours']
                    new_hours = current_hours + hours
                    
                    # Don't exceed target by more than 2%
                    if new_hours > target_hours * 1.02:
                        continue
                    
                    # Don't exceed hard limit by more than 5%
                    if new_hours > hours_limit * 1.05:
                        continue
                    
                    # Check offline limit
                    order_line_raw = item['order'].get('Suggested Line', '').strip()
                    is_offline = order_line_raw.upper() == 'OFFLINE'
                    if is_offline and day['offline_count'] >= offline_limit:
                        continue
                    
                    # SCORING
                    score = 0
                    
                    # 1. Hours fit score (how well does this fill remaining hours?)
                    if hours <= hours_remaining:
                        # Fits within remaining - prefer orders that fill more
                        fill_ratio = hours / hours_remaining if hours_remaining > 0 else 0
                        score += fill_ratio * 50
                        # Bonus for good fits (50-100% of remaining)
                        if 0.5 <= fill_ratio <= 1.0:
                            score += 30
                    else:
                        # Would overshoot - penalty based on overage
                        overage = hours - hours_remaining
                        score -= overage * 5
                    
                    # 2. Date priority (earlier = better)
                    # Calculate days from earliest available
                    earliest = min(i['start_date'] for i in available_orders)
                    days_from_earliest = (item['start_date'] - earliest).days if earliest != datetime.max else 0
                    date_score = max(0, 30 - days_from_earliest)  # Up to 30 points for early dates
                    score += date_score
                    
                    # 3. Line balance score
                    if line in ['C1', 'C2', 'C3/4']:
                        current_line_count = day['line_counts'][line]
                        target_count = target_line_counts.get(line, 0)
                        
                        if current_line_count < target_count:
                            # Under target for this line - bonus
                            score += 20
                        elif current_line_count >= target_count * 1.5:
                            # Over target - penalty
                            score -= 15
                    
                    if score > best_score:
                        best_score = score
                        best_order = item
                        best_line = line
            
            if best_order:
                self._add_order_to_day(day, best_order)
                orders_by_line[best_line].remove(best_order)
            else:
                # No suitable order found - try to find ANY order that fits
                found = False
                for line in ['C1', 'C2', 'C3/4', 'Other']:
                    for item in orders_by_line[line]:
                        if day['totals']['Hours'] + item['hours'] <= target_hours * 1.05:
                            order_line_raw = item['order'].get('Suggested Line', '').strip()
                            is_offline = order_line_raw.upper() == 'OFFLINE'
                            if not (is_offline and day['offline_count'] >= offline_limit):
                                self._add_order_to_day(day, item)
                                orders_by_line[line].remove(item)
                                found = True
                                break
                    if found:
                        break
                
                if not found:
                    break  # Can't add any more orders
        
        # Phase 3: Fine-tune to get closer to exactly 100%
        if day['totals']['Hours'] < target_hours * 0.98:
            # Try adding small orders to fill the gap
            all_remaining = []
            for line in orders_by_line:
                all_remaining.extend(orders_by_line[line])
            
            for item in sorted(all_remaining, key=lambda x: x['hours']):
                if day['totals']['Hours'] + item['hours'] <= target_hours * 1.02:
                    order_line_raw = item['order'].get('Suggested Line', '').strip()
                    is_offline = order_line_raw.upper() == 'OFFLINE'
                    if not (is_offline and day['offline_count'] >= offline_limit):
                        self._add_order_to_day(day, item)
                        orders_by_line[item['line']].remove(item)
                        if day['totals']['Hours'] >= target_hours * 0.99:
                            break
        
        # Set line distribution
        day['line_distribution'] = {
            'C1': {'count': day['line_counts']['C1'], 'hours': day['line_hours']['C1']},
            'C2': {'count': day['line_counts']['C2'], 'hours': day['line_hours']['C2']},
            'C3/4': {'count': day['line_counts']['C3/4'], 'hours': day['line_hours']['C3/4']},
            'Other': {'count': day['line_counts']['Other'], 'hours': day['line_hours']['Other']}
        }
        
        return day
    
    def _add_order_to_day(self, day: Dict, item: Dict):
        """Helper to add an order to a day and update all tracking."""
        order = item['order']
        day['orders'].append(order)
        day['totals']['Qty'] += item['qty']
        day['totals']['Picks'] += item['picks']
        day['totals']['Hours'] += item['hours']
        day['num_orders'] += 1
        
        line = item['line']
        day['line_counts'][line] += 1
        day['line_hours'][line] += item['hours']
        
        order_line_raw = order.get('Suggested Line', '').strip()
        if order_line_raw.upper() == 'OFFLINE':
            day['offline_count'] += 1
    
    def _create_remainder(self, remaining_orders: List[Dict], hours_limit: float,
                         qty_limit: float, picks_limit: float, 
                         offline_limit: float, brand: str) -> Dict:
        """Create a remainder day with leftover orders."""
        remainder = {
            'day': 'Remainder',
            'orders': [item['order'] for item in remaining_orders],
            'totals': {
                'Qty': sum(item['qty'] for item in remaining_orders),
                'Picks': sum(item['picks'] for item in remaining_orders),
                'Hours': sum(item['hours'] for item in remaining_orders)
            },
            'utilization': {
                'Qty': sum(item['qty'] for item in remaining_orders) / qty_limit * 100 if qty_limit > 0 else 0,
                'Picks': sum(item['picks'] for item in remaining_orders) / picks_limit * 100 if picks_limit > 0 else 0,
                'Hours': sum(item['hours'] for item in remaining_orders) / hours_limit * 100 if hours_limit > 0 else 0
            },
            'num_orders': len(remaining_orders),
            'line_counts': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
            'line_hours': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
            'offline_count': 0,
        }
        
        for item in remaining_orders:
            line = item['line']
            remainder['line_counts'][line] += 1
            remainder['line_hours'][line] += item['hours']
            order_line_raw = item['order'].get('Suggested Line', '').strip()
            if order_line_raw.upper() == 'OFFLINE':
                remainder['offline_count'] += 1
        
        remainder['line_distribution'] = {
            'C1': {'count': remainder['line_counts']['C1'], 'hours': remainder['line_hours']['C1']},
            'C2': {'count': remainder['line_counts']['C2'], 'hours': remainder['line_hours']['C2']},
            'C3/4': {'count': remainder['line_counts']['C3/4'], 'hours': remainder['line_hours']['C3/4']},
            'Other': {'count': remainder['line_counts']['Other'], 'hours': remainder['line_hours']['Other']}
        }
        remainder['offline_limit'] = offline_limit
        remainder['brand'] = brand
        remainder['day_label'] = 'Remainder'
        
        return remainder
    
    def _calculate_std(self, values: list) -> float:
        """Calculate standard deviation."""
        if len(values) < 2:
            return 0
        mean = sum(values) / len(values)
        variance = sum((x - mean) ** 2 for x in values) / len(values)
        return variance ** 0.5
    
    def export_to_excel(self, suggestions: List[Dict], output_path: str = 'daily_plan_suggestions.xlsx'):
        """Export suggestions to Excel file."""
        wb = openpyxl.Workbook()
        
        # Check if this is multi-day
        is_multi_day = len(suggestions) > 1 and 'day' in suggestions[0]
        
        if is_multi_day:
            ws = wb.active
            ws.title = "Multi-Day Plan"
            
            # Write summary for each day
            row = 1
            ws['A1'] = 'Day Summary'
            row += 1
            
            for suggestion in suggestions:
                day_label = suggestion.get('day_label', f"Day {suggestion.get('day', '?')}")
                ws[f'A{row}'] = day_label
                ws[f'B{row}'] = 'Orders'
                ws[f'C{row}'] = suggestion['num_orders']
                ws[f'D{row}'] = 'Qty'
                ws[f'E{row}'] = suggestion['totals']['Qty']
                ws[f'F{row}'] = 'Picks'
                ws[f'G{row}'] = suggestion['totals']['Picks']
                ws[f'H{row}'] = 'Hours'
                ws[f'I{row}'] = suggestion['totals']['Hours']
                ws[f'J{row}'] = 'Hours %'
                ws[f'K{row}'] = f"{suggestion['utilization']['Hours']:.1f}%"
                row += 1
            
            row += 1
            ws[f'A{row}'] = 'Orders:'
            row += 1
            
            # Collect all orders with day labels
            all_orders = []
            for suggestion in suggestions:
                day_label = suggestion.get('day_label', f"Day {suggestion.get('day', '?')}")
                for order in suggestion['orders']:
                    order_with_day = order.copy()
                    order_with_day['Day'] = day_label
                    all_orders.append(order_with_day)
            
            if all_orders:
                headers = ['Day'] + [k for k in all_orders[0].keys() if k != 'Day']
                for col_idx, header in enumerate(headers, 1):
                    ws.cell(row=row, column=col_idx, value=header)
                row += 1
                
                for order in all_orders:
                    col_idx = 1
                    ws.cell(row=row, column=col_idx, value=order.get('Day', ''))
                    col_idx += 1
                    for header in headers[1:]:
                        value = order.get(header, '')
                        if isinstance(value, datetime):
                            value = value.strftime('%Y-%m-%d')
                        ws.cell(row=row, column=col_idx, value=value)
                        col_idx += 1
                    row += 1
        
        wb.save(output_path)
        print(f"Exported to {output_path}")


def main():
    """Main function to run the progressive optimizer."""
    template_path = "Daily Planning Template.xlsm"
    
    optimizer = DailyPlanOptimizerProgressive(template_path=template_path)
    optimizer.load_data()
    
    print("\n" + "="*60)
    print("PROGRESSIVE FILL OPTIMIZER")
    print("="*60)
    print("\nThis optimizer will:")
    print("  1. Fill each day to 100% hours before moving to next")
    print("  2. Balance lines proportionally within each day")
    print("  3. Prioritize earlier start dates")
    print("  4. Last day gets remainder (may be < 100%)")
    
    for brand in ['BVI', 'Malosa']:
        if brand in optimizer.brand_limits:
            limits = optimizer.brand_limits[brand]
            brand_orders = [o for o in optimizer.orders if o.get('Brand', '').upper() == brand.upper()]
            
            if not brand_orders:
                print(f"\nNo orders found for {brand}")
                continue
            
            total_hours = sum(o.get('Hours', 0) or 0 for o in brand_orders)
            hours_per_day = limits['Hours']
            estimated_max_days = max(1, int(total_hours / hours_per_day) + 2)
            
            # Run the progressive multi-day planning
            day_plans = optimizer.generate_multi_day_plans(estimated_max_days, brand=brand)
            
            if day_plans:
                complete_days = [d for d in day_plans if d.get('day') != 'Remainder']
                remainder_days = [d for d in day_plans if d.get('day') == 'Remainder']
                
                print(f"\n{'='*60}")
                print(f"RESULTS: {brand}")
                print(f"{'='*60}")
                
                # Summary table
                print(f"\n{'Day':<12} {'Orders':<8} {'Qty':<10} {'Picks':<10} {'Hours':<10} {'Hours %':<10}")
                print("-" * 70)
                
                for day_plan in day_plans:
                    day_label = day_plan.get('day_label', f"Day {day_plan.get('day', '?')}")
                    print(f"{day_label:<12} {day_plan['num_orders']:<8} "
                          f"{day_plan['totals']['Qty']:<10.0f} "
                          f"{day_plan['totals']['Picks']:<10.0f} "
                          f"{day_plan['totals']['Hours']:<10.1f} "
                          f"{day_plan['utilization']['Hours']:<10.1f}%")
                
                # Calculate totals
                total_orders = sum(d['num_orders'] for d in day_plans)
                total_qty = sum(d['totals']['Qty'] for d in day_plans)
                total_picks = sum(d['totals']['Picks'] for d in day_plans)
                total_hours_planned = sum(d['totals']['Hours'] for d in day_plans)
                
                print("-" * 70)
                print(f"{'TOTAL':<12} {total_orders:<8} "
                      f"{total_qty:<10.0f} "
                      f"{total_picks:<10.0f} "
                      f"{total_hours_planned:<10.1f}")
                
                # Balance metrics for complete days only
                if len(complete_days) > 1:
                    hours_utils = [d['utilization']['Hours'] for d in complete_days]
                    orders_counts = [d['num_orders'] for d in complete_days]
                    
                    print(f"\nBalance Metrics (Complete Days, excluding Remainder):")
                    print(f"  Hours: avg={sum(hours_utils)/len(hours_utils):.1f}%, "
                          f"min={min(hours_utils):.1f}%, max={max(hours_utils):.1f}%")
                    print(f"  Orders: avg={sum(orders_counts)/len(orders_counts):.1f}, "
                          f"min={min(orders_counts)}, max={max(orders_counts)}")
                
                # Line distribution summary
                print(f"\nLine Distribution per Day:")
                for day_plan in day_plans:
                    day_label = day_plan.get('day_label', f"Day {day_plan.get('day', '?')}")
                    line_dist = day_plan.get('line_distribution', {})
                    c1 = line_dist.get('C1', {}).get('count', 0)
                    c2 = line_dist.get('C2', {}).get('count', 0)
                    c34 = line_dist.get('C3/4', {}).get('count', 0)
                    other = line_dist.get('Other', {}).get('count', 0)
                    print(f"  {day_label}: C1={c1}, C2={c2}, C3/4={c34}, Other={other}")
                
                # Start date analysis
                print(f"\nStart Date Range per Day:")
                for day_plan in day_plans:
                    day_label = day_plan.get('day_label', f"Day {day_plan.get('day', '?')}")
                    dates = [o.get('Start Date') for o in day_plan['orders'] if o.get('Start Date')]
                    if dates:
                        min_date = min(dates)
                        max_date = max(dates)
                        print(f"  {day_label}: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")
                
                # Create output directory
                output_dir = 'output'
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                # Generate timestamp
                timestamp = datetime.now().strftime('%Y%m%d%H%M')
                
                # Export
                brand_lower = brand.lower()
                excel_filename = f'{timestamp}-{brand_lower}-progressive-plan.xlsx'
                excel_path = os.path.join(output_dir, excel_filename)
                
                optimizer.export_to_excel(day_plans, excel_path)
                print(f"\nExported to: {excel_path}")
            else:
                print(f"No plans generated for {brand}")
    
    print("\n" + "="*60)
    print("Done! Check the generated Excel files.")
    print("="*60)


if __name__ == "__main__":
    main()
