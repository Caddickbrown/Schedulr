# Daily Planning Optimizer - Assumptions and Criteria

## Core Assumptions

1. **Hours is the Pivot Metric**
   - Hours MUST be hit (target: 390 hours for BVI, 90 hours for Malosa)
   - Hours has tight tolerance: 2% (can go slightly over to 5% if needed for line balancing)
   - Qty and Picks are flexible and can exceed limits (up to 50% over) to achieve hours target

2. **Brand Separation**
   - BVI and Malosa are planned separately on the same day
   - Each brand has its own limits:
     - BVI: Qty=10544, Picks=750, Hours=390
     - Malosa: Qty=3335, Picks=130, Hours=90

3. **Order Count Target**
   - Target approximately 40 orders per day
   - This is a soft target (preference, not hard constraint)

4. **Multi-Day Planning**
   - Creates as many complete day plans as possible
   - Each day must hit hours target (within tolerance)
   - Remaining orders that don't fit into complete days go into "Remainder"
   - Days are labeled "Day 1", "Day 2", etc., with "Remainder" for overflow

## Constraints and Limits

### Hard Constraints (Must Not Exceed)

1. **Hours Limit**
   - BVI: 390 hours (can go up to 5% over if needed for line balancing)
   - Malosa: 90 hours (can go up to 5% over if needed for line balancing)
   - Primary constraint - must be hit as closely as possible

2. **Offline Jobs Limit**
   - BVI: 28 offline jobs maximum per day
   - Orders with "Suggested Line" = "Offline" count toward this limit
   - Hard limit - cannot exceed

### Soft Constraints (Flexible)

1. **Qty Limit**
   - BVI: 10544 (can exceed up to 50% if needed to hit hours)
   - Malosa: 3335 (can exceed up to 50% if needed to hit hours)
   - Flexible to allow hours target to be met

2. **Picks Limit**
   - BVI: 750 (can exceed up to 50% if needed to hit hours)
   - Malosa: 130 (can exceed up to 50% if needed to hit hours)
   - Flexible to allow hours target to be met

### Balance Requirements

1. **Line Distribution (1:1:1 Ratio)**
   - C1, C2, and C3/4 work should be balanced in a 1:1:1 ratio
   - Measured by both order count and hours
   - Algorithm ensures at least one order from each line per day
   - Prefers orders that improve balance toward 1:1:1

2. **Utilization Balance**
   - Attempts to balance Qty, Picks, and Hours utilization
   - Hours is prioritized (must hit target)
   - Qty and Picks balance around Hours constraint

## Prioritization Criteria

1. **Due Date Priority**
   - Orders with earlier Start Dates are prioritized
   - Earlier dates get higher priority scores
   - Normalized over a 60-day window

2. **Hours Target Proximity**
   - Orders that get closer to hours target are preferred
   - Strong preference for orders that don't push hours too far over target

3. **Line Balance**
   - Orders that improve 1:1:1 line balance are preferred
   - Strong bonus for adding first order to an underrepresented line
   - Penalty for orders that worsen balance

4. **Order Count**
   - Preference for reaching ~40 orders per day
   - Bonus for smaller orders when below target count
   - Helps avoid hitting Qty limit too early

## Data Requirements

1. **Order Data (from CSV)**
   - Order No: Unique identifier
   - Part No: Part number
   - Brand: BVI or Malosa
   - Start Date: Due date for prioritization
   - Lot Size: Quantity
   - Picks: Number of picks
   - Hours: Total standard hours (from Main sheet)
   - Suggested Line: C1, C2, C3/4, Offline, etc.

2. **Limits Data (from limits file)**
   - Brand-specific limits for Qty, Picks, Hours
   - Offline Jobs limit
   - Other limits (Low Picks, Big Picks, Large Orders) - extracted but not yet enforced

## Algorithm Behavior

1. **Single Day Planning**
   - Fills to hours target
   - Balances line distribution (1:1:1)
   - Prioritizes by due date
   - Targets ~40 orders

2. **Multi-Day Planning**
   - Creates days sequentially
   - Each day must be "complete" (at least 50% of hours target)
   - Stops creating new days when can't fill a complete day
   - Puts all remaining orders in "Remainder"

3. **Line Balancing**
   - Ensures at least one order from C1, C2, C3/4 early in each day
   - Scores orders based on how they improve balance
   - Strong preference for underrepresented lines

4. **Offline Jobs**
   - Tracks count of offline orders per day
   - Blocks selection if limit would be exceeded
   - Limit applies per day (not cumulative across days)

## Output Format

1. **Excel Output**
   - Single file with all days
   - Separate sheets for each day (Day 1, Day 2, etc.) and Remainder
   - Summary statistics at top of each sheet
   - All orders with Day column

2. **CSV Output**
   - Single file with all days combined
   - "Day" column indicates which day or "Remainder"
   - All order data included

## Known Limitations

1. **Not Yet Enforced (Placeholders in limits file)**
   - No Duplicate Parts
   - Limit Low Picks Orders
   - Limit High Picks Orders
   - Limit High Qty Orders
   - Limit Low Qty Orders
   - Limit High Hours Orders
   - Limit Low Hours orders

2. **Remainder Category**
   - Remainder orders don't need to follow all constraints
   - Offline limit may be exceeded in Remainder (it's overflow)
   - Hours target doesn't apply to Remainder

3. **Line Balance**
   - 1:1:1 ratio is a target, not always perfectly achievable
   - Algorithm tries to balance but prioritizes hours target first

## Usage

```bash
# Extract data from template
python extract_data.py

# Single day planning
python daily_plan_optimizer.py

# Multi-day planning (creates as many complete days as possible)
python daily_plan_optimizer.py 5  # Try up to 5 days
```

## Notes

- The optimizer uses a greedy algorithm with scoring
- It prioritizes hours target above all else
- Line balance and order count are secondary priorities
- Due dates are used for tie-breaking and initial ordering
- The system is designed to be flexible with Qty/Picks to ensure hours target is met
