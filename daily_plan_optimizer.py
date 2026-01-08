"""
Daily Planning Optimizer
Generates optimal daily plans balancing Qty, Picks, and Hours based on limits.
Reads directly from Excel template file.
"""
import csv
import openpyxl
import os
from datetime import datetime
from typing import List, Dict, Tuple
import json


class DailyPlanOptimizer:
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
            # BVI limits are in the main section
            bvi_limits = {
                'Qty': limits.get('Qty', 0),
                'Picks': limits.get('Picks', 0),
                'Hours': limits.get('Hours', 0),
            }
            
            # Try to get additional BVI limits
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
            
            # Find Malosa section - look for "Malosa" in headers
            if 'Malosa' in headers:
                malosa_start = headers.index('Malosa')
                # Malosa Qty, Picks, Hours should be in next few columns
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
        
        # Set defaults if not found
        if not bvi_limits.get('Qty'):
            bvi_limits = {'Qty': 10544, 'Picks': 750, 'Hours': 390}
        if not malosa_limits.get('Qty'):
            malosa_limits = {'Qty': 3335, 'Picks': 130, 'Hours': 90}
        
        self.brand_limits = {
            'BVI': bvi_limits,
            'Malosa': malosa_limits
        }
        
        print(f"Loaded brand limits: {self.brand_limits}")
        # Keep legacy self.limits for backward compatibility (defaults to BVI)
        self.limits = self.brand_limits.get('BVI', {})
    
    def _load_orders(self):
        """Load orders directly from Excel template."""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template file not found: {self.template_path}")
        
        wb = openpyxl.load_workbook(self.template_path, data_only=True)
        main_sheet = wb['Main']
        
        # Extract order headers (row 11)
        order_headers = [cell.value for cell in main_sheet[11]]
        
        # Extract orders (starting row 12)
        for row_idx in range(12, main_sheet.max_row + 1):
            row = [cell.value for cell in main_sheet[row_idx]]
            if not row[0] or row[0] == 'Order No':
                continue
            
            try:
                # Build order dict from row
                row_dict = {}
                for idx, header in enumerate(order_headers):
                    if idx < len(row) and header:
                        value = row[idx]
                        # Keep datetime as datetime object (don't convert to string)
                        row_dict[header] = value
                
                # Normalize line name (C3/4, C3&4 -> C3/4)
                suggested_line = str(row_dict.get('Suggested Line', '')).strip()
                if suggested_line in ['C3/4', 'C3&4']:
                    suggested_line = 'C3/4'
                
                # Calculate efficiency metrics
                qty = self._parse_float(row_dict.get('Lot Size', 0))
                picks = self._parse_float(row_dict.get('Picks', 0))
                hours = self._parse_float(row_dict.get('Hours', 0))
                
                # Try to get from Excel, otherwise calculate
                qty_per_hr = self._parse_float(row_dict.get('Qty/Hr', 0))
                picks_per_hr = self._parse_float(row_dict.get('Picks/Hr', 0))
                picks_per_qty = self._parse_float(row_dict.get('Picks/Qty', 0))
                
                # Calculate if missing
                if qty_per_hr == 0 and hours > 0:
                    qty_per_hr = qty / hours
                if picks_per_hr == 0 and hours > 0:
                    picks_per_hr = picks / hours
                if picks_per_qty == 0 and qty > 0:
                    picks_per_qty = picks / qty
                
                # Parse date - handle both datetime objects and strings
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
                    'Qty/Hr': qty_per_hr,
                    'Picks/Hr': picks_per_hr,
                    'Picks/Qty': picks_per_qty,
                }
                
                # Only add if we have essential data
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
            # Try different date formats
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
    
    def _categorize_order_difficulty(self, order: Dict, brand_orders: List[Dict] = None) -> str:
        """
        Categorize order as Easy, Medium, or Difficult based on efficiency metrics.
        
        Easy: High Qty/Hr, Low Picks/Qty (efficient, low pick density)
        Medium: Moderate efficiency metrics
        Difficult: Low Qty/Hr, High Picks/Qty (inefficient, high pick density)
        """
        qty_per_hr = order.get('Qty/Hr', 0)
        picks_per_qty = order.get('Picks/Qty', 0)
        
        if qty_per_hr == 0 or picks_per_qty == 0:
            return 'Medium'
        
        # Use provided brand_orders or all orders for this brand
        if brand_orders is None:
            brand = order.get('Brand', 'BVI')
            brand_orders = [o for o in self.orders if o.get('Brand', '').upper() == brand.upper()]
        
        if len(brand_orders) < 3:
            return 'Medium'
        
        # Get efficiency metrics for all brand orders
        qty_hr_values = [o.get('Qty/Hr', 0) for o in brand_orders if o.get('Qty/Hr', 0) > 0]
        picks_qty_values = [o.get('Picks/Qty', 0) for o in brand_orders if o.get('Picks/Qty', 0) > 0]
        
        if not qty_hr_values or not picks_qty_values:
            return 'Medium'
        
        # Calculate percentiles (33rd and 67th)
        qty_hr_sorted = sorted(qty_hr_values)
        picks_qty_sorted = sorted(picks_qty_values)
        
        qty_hr_33_idx = len(qty_hr_sorted) // 3
        qty_hr_67_idx = (len(qty_hr_sorted) * 2) // 3
        picks_qty_33_idx = len(picks_qty_sorted) // 3
        picks_qty_67_idx = (len(picks_qty_sorted) * 2) // 3
        
        qty_hr_33 = qty_hr_sorted[qty_hr_33_idx] if qty_hr_33_idx < len(qty_hr_sorted) else 0
        qty_hr_67 = qty_hr_sorted[qty_hr_67_idx] if qty_hr_67_idx < len(qty_hr_sorted) else 0
        picks_qty_33 = picks_qty_sorted[picks_qty_33_idx] if picks_qty_33_idx < len(picks_qty_sorted) else 0
        picks_qty_67 = picks_qty_sorted[picks_qty_67_idx] if picks_qty_67_idx < len(picks_qty_sorted) else 0
        
        # Categorize: Easy = high Qty/Hr AND low Picks/Qty
        # Difficult = low Qty/Hr OR high Picks/Qty
        is_easy = qty_per_hr >= qty_hr_67 and picks_per_qty <= picks_qty_67
        is_difficult = qty_per_hr <= qty_hr_33 or picks_per_qty >= picks_qty_67
        
        if is_easy:
            return 'Easy'
        elif is_difficult:
            return 'Difficult'
        else:
            return 'Medium'
    
    def _calculate_line_balance_score(self, selected_orders: List[Dict]) -> float:
        """
        Calculate how balanced the line distribution is (1:1:1 ratio).
        Returns a score where higher = more balanced.
        """
        line_counts = {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0}
        line_hours = {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0}
        
        for order in selected_orders:
            line = self._get_line_category(order.get('Suggested Line', ''))
            line_counts[line] += 1
            line_hours[line] += order.get('Hours', 0) or 0
        
        # Focus on C1, C2, C3/4 (ignore Other)
        target_lines = ['C1', 'C2', 'C3/4']
        counts = [line_counts[line] for line in target_lines]
        hours = [line_hours[line] for line in target_lines]
        
        if sum(counts) == 0:
            return 0.0
        
        # Calculate balance based on both counts and hours
        # For 1:1:1 ratio, we want equal distribution
        count_mean = sum(counts) / len(counts) if len(counts) > 0 else 0
        hours_mean = sum(hours) / len(hours) if len(hours) > 0 else 0
        
        # Calculate variance (lower = more balanced)
        count_variance = sum((c - count_mean) ** 2 for c in counts) / len(counts) if len(counts) > 0 else 0
        hours_variance = sum((h - hours_mean) ** 2 for h in hours) / len(hours) if len(hours) > 0 else 0
        
        # Score: prefer lower variance (more balanced)
        # Normalize and invert so higher score = more balanced
        count_balance = 100 - (count_variance ** 0.5) * 10
        hours_balance = 100 - (hours_variance ** 0.5) * 10
        
        return (count_balance + hours_balance) / 2
    
    def optimize_plan_balanced(self, brand: str = None, limits: Dict = None) -> Tuple[List[Dict], Dict]:
        """
        Optimize daily plan with Hours as the pivot.
        - Hours MUST be hit (target the limit)
        - Target ~40 orders
        - Balance Qty and Picks around Hours constraint
        - Prioritize by due date (earlier dates first)
        
        Args:
            brand: Brand to filter orders (BVI, Malosa, etc.)
            limits: Limits dictionary to use (defaults to self.limits)
        """
        # Filter orders by brand if specified
        orders_to_optimize = self.orders
        if brand:
            orders_to_optimize = [o for o in self.orders if o.get('Brand', '').upper() == brand.upper()]
        
        if not orders_to_optimize:
            return [], {
                'totals': {'Qty': 0, 'Picks': 0, 'Hours': 0},
                'utilization': {'Qty': 0, 'Picks': 0, 'Hours': 0},
                'num_orders': 0
            }
        
        # Use provided limits or default to self.limits
        if limits is None:
            limits = self.limits
        
        target_orders = 40
        hours_target = limits['Hours']
        
        # Filter and prepare orders with metrics
        orders_with_metrics = []
        for order in orders_to_optimize:
            qty = order.get('Lot Size', 0) or 0
            picks = order.get('Picks', 0) or 0
            hours = order.get('Hours', 0) or 0
            
            if qty > 0 and hours > 0:  # Must have quantity and hours
                # Get date priority (earlier = higher priority)
                date_priority = 0.0
                if order.get('Start Date'):
                    try:
                        earliest_date = min(o.get('Start Date') for o in orders_to_optimize if o.get('Start Date'))
                        if earliest_date:
                            days_diff = (order['Start Date'] - earliest_date).days
                            # Earlier dates get higher priority (0-1 range, normalized)
                            date_priority = 1.0 - min(days_diff / 60.0, 1.0)  # 60 day window
                    except:
                        pass
                
                # Categorize order difficulty
                difficulty = self._categorize_order_difficulty(order, orders_to_optimize)
                
                orders_with_metrics.append({
                    'order': order,
                    'qty': qty,
                    'picks': picks,
                    'hours': hours,
                    'date_priority': date_priority,
                    'start_date': order.get('Start Date') or datetime.max,
                    'difficulty': difficulty
                })
        
        # Sort primarily by due date (earlier first), then by hours contribution
        orders_with_metrics.sort(key=lambda x: (x['start_date'], -x['hours']))
        
        selected = []
        totals = {'Qty': 0, 'Picks': 0, 'Hours': 0}
        line_counts = {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0}
        line_hours = {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0}
        offline_count = 0  # Track number of offline orders
        difficulty_counts = {'Easy': 0, 'Medium': 0, 'Difficult': 0}  # Track difficulty mix
        
        # Phase 1: Fill to Hours target, prioritizing by due date and line balance
        # Hours is the PIVOT - must hit this target
        # Qty and Picks are flexible (can go over if needed to hit hours)
        remaining_orders = orders_with_metrics.copy()
        hours_tolerance = 0.02  # 2% tolerance for hours (very tight)
        qty_picks_flexibility = 1.5  # Allow Qty/Picks to go 50% over if needed (very flexible)
        
        # Phase 0: Ensure at least one order from each target line early
        target_lines = ['C1', 'C2', 'C3/4']
        for target_line in target_lines:
            if line_counts[target_line] == 0:
                # Find best order for this line
                best_for_line = None
                best_score = -float('inf')
                
                for item in remaining_orders:
                    item_line = self._get_line_category(item['order'].get('Suggested Line', ''))
                    if item_line != target_line:
                        continue
                    
                    qty = item['qty']
                    picks = item['picks']
                    hours = item['hours']
                    
                    # Check basic constraints
                    if (totals['Hours'] + hours > hours_target * 1.1 or
                        totals['Qty'] + qty > limits['Qty'] * 2.0 or
                        totals['Picks'] + picks > limits['Picks'] * 2.0):
                        continue
                    
                    # Score: prefer smaller hours to leave room for more orders
                    score = -hours + item['date_priority'] * 10
                    
                    if score > best_score:
                        best_score = score
                        best_for_line = item
                
                if best_for_line:
                    selected.append(best_for_line['order'])
                    totals['Qty'] += best_for_line['qty']
                    totals['Picks'] += best_for_line['picks']
                    totals['Hours'] += best_for_line['hours']
                    order_line = self._get_line_category(best_for_line['order'].get('Suggested Line', ''))
                    line_counts[order_line] += 1
                    line_hours[order_line] += best_for_line['hours']
                    # Update difficulty tracking
                    order_difficulty = best_for_line.get('difficulty', 'Medium')
                    difficulty_counts[order_difficulty] += 1
                    remaining_orders.remove(best_for_line)
        
        iterations = 0
        max_iterations = len(orders_with_metrics) * 3  # Safety limit
        
        while remaining_orders and iterations < max_iterations:
            iterations += 1
            
            # Check if we have at least one order from each target line
            target_lines = ['C1', 'C2', 'C3/4']
            has_all_lines = all(line_counts[line] > 0 for line in target_lines)
            
            # Stop conditions: MUST reach at least 99.5% of hours target
            # 1. We've hit hours target (at least 99.5%, max 102%) AND have enough orders AND have all lines represented
            if (totals['Hours'] >= hours_target * 0.995 and 
                totals['Hours'] <= hours_target * 1.02 and
                len(selected) >= target_orders * 0.7 and  # At least 70% of target
                has_all_lines):  # Must have at least one order from each line
                break
            
            # 2. We're at or over 100% of hours target - can stop if we have reasonable orders
            # But don't stop if we're missing a line (unless we're way over hours)
            if (totals['Hours'] >= hours_target and
                len(selected) >= target_orders * 0.6):
                if has_all_lines or totals['Hours'] > hours_target * 1.05:
                    break
            
            best_order = None
            best_score = -float('inf')
            
            for item in remaining_orders:
                qty = item['qty']
                picks = item['picks']
                hours = item['hours']
                
                # Hours is the hard constraint - must stay within tolerance
                # But allow slightly more flexibility if we need to balance lines
                hours_after = totals['Hours'] + hours
                order_line = self._get_line_category(item['order'].get('Suggested Line', ''))
                max_hours_tolerance = hours_tolerance
                
                # Allow more hours flexibility if this order helps balance lines
                if order_line in ['C1', 'C2', 'C3/4']:
                    target_lines = ['C1', 'C2', 'C3/4']
                    if line_counts[order_line] == 0:
                        # Allow up to 5% over if adding first order to a line
                        max_hours_tolerance = 0.05
                    elif sum(line_counts[line] for line in target_lines) > 0:
                        # Check if this line is underrepresented
                        line_ratios = [line_counts[line] / sum(line_counts[line] for line in target_lines) 
                                      for line in target_lines]
                        target_ratio = 1.0 / 3.0
                        if line_ratios[target_lines.index(order_line)] < target_ratio * 0.7:
                            max_hours_tolerance = 0.04  # Allow 4% over for underrepresented lines
                
                if hours_after > hours_target * (1 + max_hours_tolerance):
                    continue  # Would exceed hours limit
                
                # Qty and Picks are flexible - allow significant overage to hit hours target
                qty_would_exceed = totals['Qty'] + qty > self.limits['Qty'] * qty_picks_flexibility
                picks_would_exceed = totals['Picks'] + picks > self.limits['Picks'] * qty_picks_flexibility
                
                # Only block if BOTH would exceed AND we're not close to hours target
                if qty_would_exceed and picks_would_exceed:
                    # Both would exceed - only allow if we're close to hours target or need orders
                    if totals['Hours'] < hours_target * 0.85 and len(selected) >= target_orders * 0.9:
                        continue  # Too far from target and have enough orders
                
                # Score based on:
                # 1. How close we get to hours target (CRITICAL - highest weight)
                hours_distance_from_target = abs(hours_target - hours_after)
                current_distance = abs(hours_target - totals['Hours'])
                
                if hours_after <= hours_target:
                    # We're still under target - STRONGLY prefer getting closer
                    hours_score = (current_distance - hours_distance_from_target) * 100
                    # Extra bonus if this gets us very close
                    if hours_distance_from_target < hours_target * 0.05:
                        hours_score += 50
                else:
                    # We're over target - prefer staying as close as possible
                    hours_score = -hours_distance_from_target * 200  # Strong penalty for going over
                
                # 2. Order count bonus (STRONG preference for ~40 orders)
                order_count_bonus = 0
                if len(selected) < target_orders * 0.6:
                    order_count_bonus = 40  # Very strong bonus when well below target
                elif len(selected) < target_orders * 0.8:
                    order_count_bonus = 30  # Strong bonus when below target
                elif len(selected) < target_orders:
                    order_count_bonus = 20  # Moderate bonus when approaching target
                elif len(selected) > target_orders * 1.3:
                    order_count_bonus = -20  # Penalty for too many orders
                
                # Bonus for smaller orders when we need more order count (helps reach 40)
                if len(selected) < target_orders:
                    # Prefer orders with lower qty (helps avoid hitting qty limit too early)
                    avg_order_qty = totals['Qty'] / len(selected) if len(selected) > 0 else 0
                    if qty < avg_order_qty * 0.5:  # Order is much smaller than average
                        order_count_bonus += 15  # Extra bonus for small orders
            
                # 3. Qty/Picks constraint penalty (prefer staying within limits)
                constraint_penalty = 0
                if qty_would_exceed:
                    constraint_penalty -= 5  # Small penalty for exceeding Qty
                if picks_would_exceed:
                    constraint_penalty -= 5  # Small penalty for exceeding Picks
                
                # 4. Balance score: prefer orders that help balance qty and picks
                current_qty_util = totals['Qty'] / limits['Qty'] if limits['Qty'] > 0 else 0
                current_picks_util = totals['Picks'] / limits['Picks'] if limits['Picks'] > 0 else 0
                new_qty_util = (totals['Qty'] + qty) / limits['Qty'] if limits['Qty'] > 0 else 0
                new_picks_util = (totals['Picks'] + picks) / limits['Picks'] if limits['Picks'] > 0 else 0
                
                # Calculate how balanced qty and picks are
                current_balance = abs(current_qty_util - current_picks_util)
                new_balance = abs(new_qty_util - new_picks_util)
                balance_improvement = (current_balance - new_balance) * 2
                
                # 5. Date priority bonus (earlier dates preferred)
                date_bonus = item['date_priority'] * 5
                
                # 6. Difficulty blending bonus (prefer mixing Easy/Medium/Difficult orders)
                difficulty_bonus = 0
                order_difficulty = item.get('difficulty', 'Medium')
                total_selected = len(selected)
                
                if total_selected > 0:
                    # Calculate current difficulty distribution
                    easy_ratio = difficulty_counts['Easy'] / total_selected
                    medium_ratio = difficulty_counts['Medium'] / total_selected
                    difficult_ratio = difficulty_counts['Difficult'] / total_selected
                    
                    # Target: ~30% Easy, ~40% Medium, ~30% Difficult (balanced mix)
                    target_easy = 0.30
                    target_medium = 0.40
                    target_difficult = 0.30
                    
                    # Calculate how much this order would improve the mix
                    if order_difficulty == 'Easy':
                        # Bonus if we have too few easy orders
                        if easy_ratio < target_easy * 0.8:
                            difficulty_bonus = 15 * (target_easy - easy_ratio) / target_easy
                        # Small penalty if we have too many easy orders
                        elif easy_ratio > target_easy * 1.2:
                            difficulty_bonus = -5
                    elif order_difficulty == 'Medium':
                        # Always prefer medium orders (they're the "sweet spot")
                        if medium_ratio < target_medium:
                            difficulty_bonus = 20 * (target_medium - medium_ratio) / target_medium
                    elif order_difficulty == 'Difficult':
                        # Strong bonus if we have too few difficult orders (need to blend them in)
                        if difficult_ratio < target_difficult * 0.7:
                            difficulty_bonus = 25 * (target_difficult - difficult_ratio) / target_difficult
                        # Small penalty if we have too many difficult orders
                        elif difficult_ratio > target_difficult * 1.3:
                            difficulty_bonus = -3
                else:
                    # First order - prefer medium difficulty to start balanced
                    if order_difficulty == 'Medium':
                        difficulty_bonus = 10
                
                # 7. Line balance bonus (prefer orders that help achieve 1:1:1 ratio for C1:C2:C3/4)
                # Note: order_line already calculated above
                line_balance_bonus = 0
                
                if order_line in ['C1', 'C2', 'C3/4']:
                    target_lines = ['C1', 'C2', 'C3/4']
                    current_counts = [line_counts[line] for line in target_lines]
                    
                    # CRITICAL: If a line has zero orders, strongly prefer adding to it
                    if line_counts[order_line] == 0:
                        # Very strong bonus for adding first order to a line
                        line_balance_bonus += 50
                    
                    # Calculate what the balance would be if we add this order
                    temp_line_counts = line_counts.copy()
                    temp_line_hours = line_hours.copy()
                    temp_line_counts[order_line] += 1
                    temp_line_hours[order_line] += hours
                    
                    current_hours = [line_hours[line] for line in target_lines]
                    new_counts = [temp_line_counts[line] for line in target_lines]
                    new_hours = [temp_line_hours[line] for line in target_lines]
                    
                    # Calculate how balanced we are (1:1:1 ratio)
                    # Lower variance = more balanced
                    if sum(current_counts) > 0:
                        current_mean = sum(current_counts) / len(current_counts)
                        current_variance = sum((c - current_mean) ** 2 for c in current_counts) / len(current_counts)
                    else:
                        current_variance = float('inf')
                    
                    if sum(new_counts) > 0:
                        new_mean = sum(new_counts) / len(new_counts)
                        new_variance = sum((c - new_mean) ** 2 for c in new_counts) / len(new_counts)
                    else:
                        new_variance = 0
                    
                    # Prefer orders that reduce variance (improve balance)
                    variance_improvement = current_variance - new_variance
                    
                    # Also consider hours balance
                    if sum(current_hours) > 0:
                        current_hours_mean = sum(current_hours) / len(current_hours)
                        current_hours_variance = sum((h - current_hours_mean) ** 2 for h in current_hours) / len(current_hours)
                    else:
                        current_hours_variance = float('inf')
                    
                    if sum(new_hours) > 0:
                        new_hours_mean = sum(new_hours) / len(new_hours)
                        new_hours_variance = sum((h - new_hours_mean) ** 2 for h in new_hours) / len(new_hours)
                    else:
                        new_hours_variance = 0
                    
                    hours_variance_improvement = current_hours_variance - new_hours_variance
                    
                    # Strong bonus for improving balance (weighted by how far we are from target)
                    line_balance_bonus += (variance_improvement * 40 + hours_variance_improvement * 20)
                    
                    # Extra bonus if this line is underrepresented (stronger preference)
                    if sum(new_counts) > 0:
                        line_ratio = new_counts[target_lines.index(order_line)] / sum(new_counts)
                        target_ratio = 1.0 / 3.0  # 1:1:1 ratio
                        if line_ratio < target_ratio * 0.8:  # Line is underrepresented
                            # Stronger bonus the more underrepresented
                            underrepresentation = (target_ratio - line_ratio) / target_ratio
                            line_balance_bonus += 30 * underrepresentation
                
                score = hours_score + balance_improvement + date_bonus + order_count_bonus + constraint_penalty + difficulty_bonus + line_balance_bonus
                
                if score > best_score:
                    best_score = score
                    best_order = item
            
            if best_order:
                selected.append(best_order['order'])
                totals['Qty'] += best_order['qty']
                totals['Picks'] += best_order['picks']
                totals['Hours'] += best_order['hours']
                
                # Update line tracking
                order_line = self._get_line_category(best_order['order'].get('Suggested Line', ''))
                line_counts[order_line] += 1
                line_hours[order_line] += best_order['hours']
                
                # Update difficulty tracking
                order_difficulty = best_order.get('difficulty', 'Medium')
                difficulty_counts[order_difficulty] += 1
                
                # Update offline count
                order_line_raw = best_order['order'].get('Suggested Line', '').strip()
                if order_line_raw.upper() == 'OFFLINE':
                    offline_count += 1
                
                remaining_orders.remove(best_order)
            else:
                # No more orders can fit - try relaxing constraints slightly
                if totals['Hours'] < hours_target * 0.995:  # Must reach at least 99.5%
                    # Still far from target - allow more flexibility
                    qty_picks_flexibility = min(qty_picks_flexibility * 1.1, 1.5)  # Allow up to 50% over
                    hours_tolerance = min(hours_tolerance * 1.2, 0.05)  # Allow up to 5% over
                elif totals['Hours'] >= hours_target * 0.995:
                    # We've reached at least 99.5% - this is acceptable
                    break
                else:
                    # Can't add more and we're below 99.5% - keep trying with relaxed constraints
                    qty_picks_flexibility = min(qty_picks_flexibility * 1.1, 1.5)
                    hours_tolerance = min(hours_tolerance * 1.2, 0.05)
        
        # Phase 2: Fine-tune to get closer to hours target AND order count target
        # ALWAYS run Phase 2 to try to reach 100% hours target
        # Continue adding orders if we need more orders OR need to reach 100% hours
        if len(selected) < target_orders * 1.1 or totals['Hours'] < hours_target * 0.995:
            # Sort remaining by: 1) date priority, 2) how close they get us to hours target
            remaining_orders.sort(key=lambda x: (
                x['start_date'],
                abs(hours_target - totals['Hours'] - x['hours'])
            ))
            
            for item in remaining_orders[:300]:  # Check top 300 candidates
                # Stop if we've hit hours target (at least 99.5%) and have enough orders
                if (totals['Hours'] >= hours_target * 0.995 and
                    len(selected) >= target_orders * 0.8):
                    break
                
                # Don't go too far over on orders
                if len(selected) >= target_orders * 1.5:
                    break
                    
                qty = item['qty']
                picks = item['picks']
                hours = item['hours']
                
                # Hours must stay within tolerance (max 2% over, but prefer reaching 100%)
                new_hours = totals['Hours'] + hours
                if new_hours > hours_target * 1.02:
                    continue
                
                # If we're still under target, strongly prefer orders that get us closer
                if totals['Hours'] < hours_target * 0.995:
                    # Must get closer to target
                    current_hours_distance = abs(hours_target - totals['Hours'])
                    new_hours_distance = abs(hours_target - new_hours)
                    if new_hours_distance >= current_hours_distance:
                        continue  # Skip if it doesn't improve
                
                # Qty/Picks can be flexible (up to 50% over)
                if (totals['Qty'] + qty > limits['Qty'] * 1.5 or
                    totals['Picks'] + picks > limits['Picks'] * 1.5):
                    continue
                
                # Decision logic:
                # Priority 1: Get hours as close to target as possible
                # Priority 2: Get order count to ~40
                
                current_hours_distance = abs(hours_target - totals['Hours'])
                new_hours_distance = abs(hours_target - new_hours)
                hours_improves = new_hours_distance < current_hours_distance
                
                if totals['Hours'] < hours_target * 0.995:
                    # Under 99.5% target - MUST add orders that get us closer
                    if hours_improves:
                        selected.append(item['order'])
                        totals['Qty'] += qty
                        totals['Picks'] += picks
                        totals['Hours'] += hours
                        order_difficulty = item.get('difficulty', 'Medium')
                        difficulty_counts[order_difficulty] += 1
                        remaining_orders.remove(item)
                elif totals['Hours'] >= hours_target * 0.995 and totals['Hours'] <= hours_target * 1.01:
                    # Very close to target (within 1%) - can add small orders for order count
                    if len(selected) < target_orders:
                        # Prefer small orders that don't push us too far over
                        if hours < hours_target * 0.03:  # Orders < 3% of target
                            selected.append(item['order'])
                            totals['Qty'] += qty
                            totals['Picks'] += picks
                            totals['Hours'] += hours
                            order_difficulty = item.get('difficulty', 'Medium')
                            difficulty_counts[order_difficulty] += 1
                            remaining_orders.remove(item)
                        elif hours_improves:  # Or if it actually improves hours
                            selected.append(item['order'])
                            totals['Qty'] += qty
                            totals['Picks'] += picks
                            totals['Hours'] += hours
                            order_difficulty = item.get('difficulty', 'Medium')
                            difficulty_counts[order_difficulty] += 1
                            remaining_orders.remove(item)
                else:
                    # Over target (1-2%) - only add if it improves hours AND we need orders
                    if hours_improves and len(selected) < target_orders * 0.9:
                        selected.append(item['order'])
                        totals['Qty'] += qty
                        totals['Picks'] += picks
                        totals['Hours'] += hours
                        order_difficulty = item.get('difficulty', 'Medium')
                        difficulty_counts[order_difficulty] += 1
                        remaining_orders.remove(item)
        
        # Calculate final utilization
        utilization = {
            'Qty': totals['Qty'] / limits['Qty'] * 100 if limits['Qty'] > 0 else 0,
            'Picks': totals['Picks'] / limits['Picks'] * 100 if limits['Picks'] > 0 else 0,
            'Hours': totals['Hours'] / limits['Hours'] * 100 if limits['Hours'] > 0 else 0
        }
        
        # Calculate line distribution
        line_distribution = {
            'C1': {'count': line_counts['C1'], 'hours': line_hours['C1']},
            'C2': {'count': line_counts['C2'], 'hours': line_hours['C2']},
            'C3/4': {'count': line_counts['C3/4'], 'hours': line_hours['C3/4']},
            'Other': {'count': line_counts['Other'], 'hours': line_hours['Other']}
        }
        
        return selected, {
            'totals': totals,
            'utilization': utilization,
            'num_orders': len(selected),
            'line_distribution': line_distribution,
            'offline_count': offline_count,
            'offline_limit': limits.get('Offline Jobs', None)
        }
    
    def _calculate_balance_score(self, utilization: Dict) -> float:
        """Calculate how balanced the utilization is across all three metrics."""
        # Lower standard deviation = more balanced
        values = [utilization['Qty'], utilization['Picks'], utilization['Hours']]
        mean = sum(values) / len(values)
        variance = sum((x - mean) ** 2 for x in values) / len(values)
        std_dev = variance ** 0.5
        
        # Return negative std_dev (lower is better, so we maximize negative)
        # Also factor in total utilization
        total_util = sum(values) / len(values)
        return total_util - std_dev  # Higher is better
    
    def generate_suggestions(self, num_suggestions: int = 1, brand: str = None) -> List[Dict]:
        """Generate plan suggestions for a specific brand."""
        suggestions = []
        
        # Get limits for brand
        if brand:
            limits = self.brand_limits.get(brand, self.limits)
        else:
            limits = self.limits
            brand = 'BVI'  # Default
        
        # Generate balanced plan
        selected_orders, stats = self.optimize_plan_balanced(brand=brand, limits=limits)
        
        suggestions.append({
            'strategy': 'balanced',
            'brand': brand,
            'orders': selected_orders,
            'totals': stats['totals'],
            'utilization': stats['utilization'],
            'num_orders': stats['num_orders']
        })
        
        return suggestions
    
    def generate_all_brand_suggestions(self) -> Dict[str, List[Dict]]:
        """Generate suggestions for all brands (BVI and Malosa)."""
        all_suggestions = {}
        
        for brand in ['BVI', 'Malosa']:
            if brand in self.brand_limits:
                suggestions = self.generate_suggestions(brand=brand)
                all_suggestions[brand] = suggestions
        
        return all_suggestions
    
    def generate_multi_day_plans(self, num_days: int, brand: str = None) -> List[Dict]:
        """
        Generate multi-day plans with balanced utilization across days.
        
        Args:
            num_days: Number of days to plan
            brand: Brand to plan (BVI, Malosa, etc.)
        
        Returns:
            List of day plans, each with 'day', 'orders', 'totals', 'utilization', 'num_orders'
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
        
        hours_target = limits['Hours']
        target_orders_per_day = 40
        
        # Prepare orders with metrics
        orders_with_metrics = []
        for order in brand_orders:
            qty = order.get('Lot Size', 0) or 0
            picks = order.get('Picks', 0) or 0
            hours = order.get('Hours', 0) or 0
            
            if qty > 0 and hours > 0:
                # Get date priority
                date_priority = 0.0
                if order.get('Start Date'):
                    try:
                        earliest_date = min(o.get('Start Date') for o in brand_orders if o.get('Start Date'))
                        if earliest_date:
                            days_diff = (order['Start Date'] - earliest_date).days
                            date_priority = 1.0 - min(days_diff / 60.0, 1.0)
                    except:
                        pass
                
                # Categorize order difficulty
                difficulty = self._categorize_order_difficulty(order, brand_orders)
                
                orders_with_metrics.append({
                    'order': order,
                    'qty': qty,
                    'picks': picks,
                    'hours': hours,
                    'date_priority': date_priority,
                    'start_date': order.get('Start Date') or datetime.max,
                    'difficulty': difficulty
                })
        
        # Sort by date priority
        orders_with_metrics.sort(key=lambda x: (x['start_date'], -x['hours']))
        
        # Initialize days with line tracking
        days = []
        offline_limit = limits.get('Offline Jobs', float('inf'))
        for day_num in range(1, num_days + 1):
            days.append({
                'day': day_num,
                'orders': [],
                'totals': {'Qty': 0, 'Picks': 0, 'Hours': 0},
                'utilization': {'Qty': 0, 'Picks': 0, 'Hours': 0},
                'num_orders': 0,
                'line_counts': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
                'line_hours': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
                'offline_count': 0,
                'difficulty_counts': {'Easy': 0, 'Medium': 0, 'Difficult': 0}
            })
        
        # Phase 0: Ensure each day gets at least one order from each target line
        target_lines = ['C1', 'C2', 'C3/4']
        remaining_orders = orders_with_metrics.copy()
        
        for day in days:
            for target_line in target_lines:
                if day['line_counts'][target_line] == 0:
                    # Find best order for this line for this day
                    best_for_line = None
                    best_score = -float('inf')
                    
                    for item in remaining_orders:
                        item_line = self._get_line_category(item['order'].get('Suggested Line', ''))
                        if item_line != target_line:
                            continue
                        
                        qty = item['qty']
                        picks = item['picks']
                        hours = item['hours']
                        
                        # Check basic constraints
                        if (day['totals']['Hours'] + hours > hours_target * 1.1 or
                            day['totals']['Qty'] + qty > limits['Qty'] * 2.0 or
                            day['totals']['Picks'] + picks > limits['Picks'] * 2.0):
                            continue
                        
                        # Score: prefer smaller hours to leave room for more orders
                        score = -hours + item['date_priority'] * 10
                        
                        if score > best_score:
                            best_score = score
                            best_for_line = item
                    
                    if best_for_line:
                        day['orders'].append(best_for_line['order'])
                        day['totals']['Qty'] += best_for_line['qty']
                        day['totals']['Picks'] += best_for_line['picks']
                        day['totals']['Hours'] += best_for_line['hours']
                        day['num_orders'] += 1
                        order_line = self._get_line_category(best_for_line['order'].get('Suggested Line', ''))
                        day['line_counts'][order_line] += 1
                        day['line_hours'][order_line] += best_for_line['hours']
                        # Update difficulty tracking
                        order_difficulty = best_for_line.get('difficulty', 'Medium')
                        day['difficulty_counts'][order_difficulty] += 1
                        remaining_orders.remove(best_for_line)
        
        # Fill each day sequentially to 100% before moving to next day
        # This ensures each day reaches the hours target
        completed_days = []
        
        for day_idx, day in enumerate(days):
            if not remaining_orders:
                break
            
            # Use the single-day optimizer to fill this day to 100%
            # Create a temporary optimizer with only remaining orders
            temp_orders = [item['order'] for item in remaining_orders]
            
            # Use optimize_plan_balanced to fill this day
            # We'll manually track and add orders to this day
            day_hours_target = hours_target
            day_selected = []
            day_totals = {'Qty': day['totals']['Qty'], 'Picks': day['totals']['Picks'], 'Hours': day['totals']['Hours']}
            day_line_counts = day['line_counts'].copy()
            day_line_hours = day['line_hours'].copy()
            day_offline_count = day['offline_count']
            day_difficulty_counts = day['difficulty_counts'].copy()
            
            # Continue adding orders until we reach at least 99.5% of hours target
            max_iterations = len(remaining_orders) * 3
            iteration = 0
            
            while remaining_orders and day_totals['Hours'] < day_hours_target * 0.995 and iteration < max_iterations:
                iteration += 1
                
                best_order = None
                best_score = -float('inf')
                
                for item in remaining_orders:
                    qty = item['qty']
                    picks = item['picks']
                    hours = item['hours']
                    
                    # Check if adding this order would exceed limits
                    new_hours = day_totals['Hours'] + hours
                    if new_hours > day_hours_target * 1.02:  # Max 2% over
                        continue
                    
                    # Check Offline Jobs limit
                    order_line_raw = item['order'].get('Suggested Line', '').strip()
                    is_offline = order_line_raw.upper() == 'OFFLINE'
                    if is_offline and day_offline_count >= offline_limit:
                        continue
                    
                    # Check Qty/Picks limits (flexible up to 50% over)
                    if (day_totals['Qty'] + qty > limits['Qty'] * 1.5 or
                        day_totals['Picks'] + picks > limits['Picks'] * 1.5):
                        continue
                    
                    # Score: prioritize getting closer to hours target
                    current_distance = abs(day_hours_target - day_totals['Hours'])
                    new_distance = abs(day_hours_target - new_hours)
                    hours_improvement = current_distance - new_distance
                    
                    # Strongly prefer orders that get us closer to 100%
                    score = hours_improvement * 100
                    
                    # Bonus for line balance
                    order_line = self._get_line_category(item['order'].get('Suggested Line', ''))
                    if order_line in ['C1', 'C2', 'C3/4']:
                        if day_line_counts[order_line] == 0:
                            score += 50  # Strong bonus for first order in a line
                    
                    # Bonus for difficulty blending
                    order_difficulty = item.get('difficulty', 'Medium')
                    if day_totals['Hours'] > 0:
                        total_selected = len(day_selected)
                        if total_selected > 0:
                            easy_ratio = day_difficulty_counts['Easy'] / total_selected
                            difficult_ratio = day_difficulty_counts['Difficult'] / total_selected
                            if order_difficulty == 'Difficult' and difficult_ratio < 0.3:
                                score += 20  # Bonus for difficult orders if underrepresented
                    
                    if score > best_score:
                        best_score = score
                        best_order = item
                
                if best_order:
                    # Add order to day
                    day_selected.append(best_order['order'])
                    day_totals['Qty'] += best_order['qty']
                    day_totals['Picks'] += best_order['picks']
                    day_totals['Hours'] += best_order['hours']
                    
                    order_line = self._get_line_category(best_order['order'].get('Suggested Line', ''))
                    day_line_counts[order_line] += 1
                    day_line_hours[order_line] += best_order['hours']
                    
                    order_difficulty = best_order.get('difficulty', 'Medium')
                    day_difficulty_counts[order_difficulty] += 1
                    
                    order_line_raw = best_order['order'].get('Suggested Line', '').strip()
                    if order_line_raw.upper() == 'OFFLINE':
                        day_offline_count += 1
                    
                    remaining_orders.remove(best_order)
                else:
                    # Can't add more orders to this day
                    break
            
            # Update day with final totals
            day['orders'].extend(day_selected)
            day['totals'] = day_totals
            day['num_orders'] = len(day['orders'])
            day['line_counts'] = day_line_counts
            day['line_hours'] = day_line_hours
            day['offline_count'] = day_offline_count
            day['difficulty_counts'] = day_difficulty_counts
            day['utilization'] = {
                'Qty': day_totals['Qty'] / limits['Qty'] * 100 if limits['Qty'] > 0 else 0,
                'Picks': day_totals['Picks'] / limits['Picks'] * 100 if limits['Picks'] > 0 else 0,
                'Hours': day_totals['Hours'] / limits['Hours'] * 100 if limits['Hours'] > 0 else 0
            }
            
            # Only keep days that reached at least 50% of hours target
            if day_totals['Hours'] >= day_hours_target * 0.5:
                completed_days.append(day)
            else:
                # Put orders back and stop
                for order in day_selected:
                    # Find the original item in remaining_orders
                    for orig_item in orders_with_metrics:
                        if orig_item['order']['Order No'] == order.get('Order No'):
                            if orig_item not in remaining_orders:
                                remaining_orders.append(orig_item)
                            break
                break
        
        # Replace days with completed days
        days = completed_days
        
        # Create Remainder with all leftover orders
        if remaining_orders:
            remainder = {
                'day': 'Remainder',
                'orders': [item['order'] for item in remaining_orders],
                'totals': {
                    'Qty': sum(item['qty'] for item in remaining_orders),
                    'Picks': sum(item['picks'] for item in remaining_orders),
                    'Hours': sum(item['hours'] for item in remaining_orders)
                },
                'utilization': {
                    'Qty': sum(item['qty'] for item in remaining_orders) / limits['Qty'] * 100 if limits['Qty'] > 0 else 0,
                    'Picks': sum(item['picks'] for item in remaining_orders) / limits['Picks'] * 100 if limits['Picks'] > 0 else 0,
                    'Hours': sum(item['hours'] for item in remaining_orders) / limits['Hours'] * 100 if limits['Hours'] > 0 else 0
                },
                'num_orders': len(remaining_orders),
                'line_counts': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
                'line_hours': {'C1': 0, 'C2': 0, 'C3/4': 0, 'Other': 0},
                'offline_count': 0,
                'difficulty_counts': {'Easy': 0, 'Medium': 0, 'Difficult': 0}
            }
            
            # Calculate line distribution for remainder
            for item in remaining_orders:
                order_line = self._get_line_category(item['order'].get('Suggested Line', ''))
                remainder['line_counts'][order_line] += 1
                remainder['line_hours'][order_line] += item['hours']
                order_line_raw = item['order'].get('Suggested Line', '').strip()
                if order_line_raw.upper() == 'OFFLINE':
                    remainder['offline_count'] += 1
                order_difficulty = item.get('difficulty', 'Medium')
                remainder['difficulty_counts'][order_difficulty] += 1
            
            remainder['line_distribution'] = {
                'C1': {'count': remainder['line_counts']['C1'], 'hours': remainder['line_hours']['C1']},
                'C2': {'count': remainder['line_counts']['C2'], 'hours': remainder['line_hours']['C2']},
                'C3/4': {'count': remainder['line_counts']['C3/4'], 'hours': remainder['line_hours']['C3/4']},
                'Other': {'count': remainder['line_counts']['Other'], 'hours': remainder['line_hours']['Other']}
            }
            remainder['offline_limit'] = offline_limit
            remainder['brand'] = brand
            remainder['day_label'] = 'Remainder'
            days.append(remainder)
        
        # Finalize day plans
        for day in days:
            if day.get('day') != 'Remainder':
                day['brand'] = brand
                day['day_label'] = f"Day {day['day']}"
                if 'line_distribution' not in day:
                    day['line_distribution'] = {
                        'C1': {'count': day['line_counts']['C1'], 'hours': day['line_hours']['C1']},
                        'C2': {'count': day['line_counts']['C2'], 'hours': day['line_hours']['C2']},
                        'C3/4': {'count': day['line_counts']['C3/4'], 'hours': day['line_hours']['C3/4']},
                        'Other': {'count': day['line_counts']['Other'], 'hours': day['line_hours']['Other']}
                    }
                day['offline_count'] = day.get('offline_count', 0)
                day['offline_limit'] = offline_limit
        
        return days
    
    def export_to_excel(self, suggestions: List[Dict], output_path: str = 'daily_plan_suggestions.xlsx'):
        """Export suggestions to Excel file."""
        wb = openpyxl.Workbook()
        
        # Check if this is multi-day (all suggestions have 'day' key)
        is_multi_day = len(suggestions) > 1 and 'day' in suggestions[0]
        
        if is_multi_day:
            # Combine all days into a single sheet with Day column
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
                # Write headers with Day column first
                headers = ['Day'] + [k for k in all_orders[0].keys() if k != 'Day']
                for col_idx, header in enumerate(headers, 1):
                    ws.cell(row=row, column=col_idx, value=header)
                row += 1
                
                # Write orders
                for order in all_orders:
                    col_idx = 1
                    # Write Day first
                    ws.cell(row=row, column=col_idx, value=order.get('Day', ''))
                    col_idx += 1
                    # Write other fields
                    for header in headers[1:]:
                        value = order.get(header, '')
                        if isinstance(value, datetime):
                            value = value.strftime('%Y-%m-%d')
                        ws.cell(row=row, column=col_idx, value=value)
                        col_idx += 1
                    row += 1
        else:
            # Original single-day behavior
            for suggestion in suggestions:
                # Determine sheet name - use day label if available, otherwise strategy
                if 'day_label' in suggestion:
                    sheet_name = suggestion['day_label']
                elif 'day' in suggestion:
                    sheet_name = f"Day {suggestion['day']}"
                else:
                    sheet_name = suggestion.get('strategy', 'Plan').title()
                
                # Limit sheet name to 31 characters (Excel limit)
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                
                # Create sheet for this suggestion
                ws = wb.create_sheet(title=sheet_name)
                
                # Write summary
                if 'day_label' in suggestion:
                    ws['A1'] = 'Day'
                    ws['B1'] = suggestion['day_label']
                else:
                    ws['A1'] = 'Strategy'
                    ws['B1'] = suggestion.get('strategy', 'N/A')
                
                if 'brand' in suggestion:
                    ws['A2'] = 'Brand'
                    ws['B2'] = suggestion['brand']
                    row_offset = 1
                else:
                    row_offset = 0
                
                ws[f'A{3+row_offset}'] = 'Number of Orders'
                ws[f'B{3+row_offset}'] = suggestion['num_orders']
                ws[f'A{4+row_offset}'] = 'Total Qty'
                ws[f'B{4+row_offset}'] = suggestion['totals']['Qty']
                ws[f'A{5+row_offset}'] = 'Total Picks'
                ws[f'B{5+row_offset}'] = suggestion['totals']['Picks']
                ws[f'A{6+row_offset}'] = 'Total Hours'
                ws[f'B{6+row_offset}'] = suggestion['totals']['Hours']
                ws[f'A{7+row_offset}'] = 'Qty Utilization %'
                ws[f'B{7+row_offset}'] = f"{suggestion['utilization']['Qty']:.1f}%"
                ws[f'A{8+row_offset}'] = 'Picks Utilization %'
                ws[f'B{8+row_offset}'] = f"{suggestion['utilization']['Picks']:.1f}%"
                ws[f'A{9+row_offset}'] = 'Hours Utilization %'
                ws[f'B{9+row_offset}'] = f"{suggestion['utilization']['Hours']:.1f}%"
                
                # Write orders
                if suggestion['orders']:
                    headers = list(suggestion['orders'][0].keys())
                    ws[f'A{11+row_offset}'] = 'Orders:'
                    for col_idx, header in enumerate(headers, 1):
                        ws.cell(row=12+row_offset, column=col_idx, value=header)
                    
                    for row_idx, order in enumerate(suggestion['orders'], 13+row_offset):
                        for col_idx, header in enumerate(headers, 1):
                            value = order.get(header, '')
                            if isinstance(value, datetime):
                                value = value.strftime('%Y-%m-%d')
                            ws.cell(row=row_idx, column=col_idx, value=value)
        
        # Remove default sheet if we created new ones
        if not is_multi_day and 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        wb.save(output_path)
        print(f"Exported to {output_path}")
    
    def export_to_csv(self, suggestions: List[Dict], output_path: str = 'daily_plan_suggestions.csv'):
        """Export suggestions to CSV file."""
        if not suggestions:
            return
        
        # Check if this is multi-day (all suggestions have 'day' key)
        is_multi_day = len(suggestions) > 1 and 'day' in suggestions[0]
        
        if is_multi_day:
            # Combine all days into a single CSV with Day column
            all_orders = []
            for suggestion in suggestions:
                day_label = suggestion.get('day_label', f"Day {suggestion.get('day', '?')}")
                for order in suggestion['orders']:
                    order_with_day = order.copy()
                    order_with_day['Day'] = day_label
                    all_orders.append(order_with_day)
            
            if all_orders:
                # Get fieldnames with Day first
                fieldnames = ['Day'] + [k for k in all_orders[0].keys() if k != 'Day']
                
                with open(output_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    
                    for order in all_orders:
                        row = {}
                        for key in fieldnames:
                            value = order.get(key, '')
                            if isinstance(value, datetime):
                                row[key] = value.strftime('%Y-%m-%d')
                            else:
                                row[key] = value
                        writer.writerow(row)
                print(f"Exported to {output_path}")
        else:
            # Single suggestion - export normally
            suggestion = suggestions[0]
            
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                if suggestion['orders']:
                    fieldnames = list(suggestion['orders'][0].keys())
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    
                    for order in suggestion['orders']:
                        # Convert datetime to string
                        row = {}
                        for key, value in order.items():
                            if isinstance(value, datetime):
                                row[key] = value.strftime('%Y-%m-%d')
                            else:
                                row[key] = value
                        writer.writerow(row)
            print(f"Exported to {output_path}")


def main():
    """Main function to run the optimizer."""
    import sys
    
    template_path = "Daily Planning Template.xlsm"
    
    # Load data directly from Excel template
    optimizer = DailyPlanOptimizer(template_path=template_path)
    optimizer.load_data()
    
    # Automatically determine maximum possible days for each brand
    # Calculate based on total hours available vs hours per day
    print("\n" + "="*60)
    print("AUTOMATIC MULTI-DAY PLANNING BY BRAND")
    print("="*60)
    
    for brand in ['BVI', 'Malosa']:
        if brand in optimizer.brand_limits:
            print(f"\n{'='*60}")
            print(f"BRAND: {brand}")
            print(f"{'='*60}")
            
            # Calculate estimated max days based on total hours
            limits = optimizer.brand_limits[brand]
            brand_orders = [o for o in optimizer.orders if o.get('Brand', '').upper() == brand.upper()]
            
            if brand_orders:
                total_hours = sum(o.get('Hours', 0) or 0 for o in brand_orders)
                hours_per_day = limits['Hours']
                # Estimate max days (add 50% buffer to ensure we try enough)
                estimated_max_days = max(1, int((total_hours / hours_per_day) * 1.5) + 5)
                print(f"Estimated maximum days: {estimated_max_days} (based on {total_hours:.0f} total hours / {hours_per_day:.0f} hours per day)")
            else:
                estimated_max_days = 10  # Default if no orders
            
            # Generate plans - function will automatically stop when no more complete days can be created
            day_plans = optimizer.generate_multi_day_plans(estimated_max_days, brand=brand)
            
            if day_plans:
                # Count actual complete days (exclude Remainder)
                complete_days = [d for d in day_plans if d.get('day') != 'Remainder']
                num_complete_days = len(complete_days)
                
                print(f"\nGenerated {num_complete_days} complete day(s) + Remainder")
                print(f"\nSummary across all days:")
                total_orders = sum(d['num_orders'] for d in day_plans)
                total_hours_planned = sum(d['totals']['Hours'] for d in complete_days)
                if num_complete_days > 0:
                    avg_orders = sum(d['num_orders'] for d in complete_days) / num_complete_days
                    avg_hours = total_hours_planned / num_complete_days
                    print(f"  Complete days: {num_complete_days}")
                    print(f"  Total orders in complete days: {sum(d['num_orders'] for d in complete_days)} (avg {avg_orders:.1f} per day)")
                    print(f"  Total hours in complete days: {total_hours_planned:.2f} (avg {avg_hours:.2f} per day)")
                
                remainder = [d for d in day_plans if d.get('day') == 'Remainder']
                if remainder:
                    remainder_plan = remainder[0]
                    print(f"  Remainder: {remainder_plan['num_orders']} orders, {remainder_plan['totals']['Hours']:.2f} hours")
                
                # Print each day
                for day_plan in day_plans:
                    day_label = day_plan.get('day_label', f"Day {day_plan.get('day', '?')}")
                    print(f"\n--- {day_label} ---")
                    print(f"Number of orders: {day_plan['num_orders']}")
                    print(f"Totals: Qty={day_plan['totals']['Qty']:.0f}, "
                          f"Picks={day_plan['totals']['Picks']:.0f}, "
                          f"Hours={day_plan['totals']['Hours']:.2f}")
                    print(f"Utilization: Qty={day_plan['utilization']['Qty']:.1f}%, "
                          f"Picks={day_plan['utilization']['Picks']:.1f}%, "
                          f"Hours={day_plan['utilization']['Hours']:.1f}%")
                    
                    # Show line distribution
                    if 'line_distribution' in day_plan:
                        line_dist = day_plan['line_distribution']
                        print(f"Line Distribution (target 1:1:1):")
                        print(f"  C1: {line_dist['C1']['count']} orders, {line_dist['C1']['hours']:.2f} hours")
                        print(f"  C2: {line_dist['C2']['count']} orders, {line_dist['C2']['hours']:.2f} hours")
                        print(f"  C3/4: {line_dist['C3/4']['count']} orders, {line_dist['C3/4']['hours']:.2f} hours")
                    
                    # Show offline jobs count
                    if 'offline_count' in day_plan and 'offline_limit' in day_plan:
                        offline_count = day_plan.get('offline_count', 0)
                        offline_limit = day_plan.get('offline_limit')
                        if offline_limit and offline_limit != float('inf'):
                            print(f"Offline Jobs: {offline_count} / {offline_limit:.0f}")
                
                # Create output directory if it doesn't exist
                output_dir = 'output'
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                # Generate timestamp in format YYYYMMDDHHMM
                timestamp = datetime.now().strftime('%Y%m%d%H%M')
                
                # Export with timestamped filename
                brand_lower = brand.lower()
                excel_filename = f'{timestamp}-{brand_lower}-plan-suggestion.xlsx'
                
                excel_path = os.path.join(output_dir, excel_filename)
                
                optimizer.export_to_excel(day_plans, excel_path)
            else:
                print(f"No plans generated for {brand}")
    
    print("\n" + "="*60)
    print("Done! Check the generated Excel files for each brand.")
    print("="*60)


if __name__ == "__main__":
    main()