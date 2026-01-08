# Daily Planning Optimizer - Assumptions and Criteria

## Core Assumptions

1. **Hours is the Pivot Metric**
   - Hours MUST be hit (target: 390 hours for BVI, 90 hours for Malosa)
   - **Minimum requirement: 99.5% of hours target** (algorithm will not stop until at least 99.5% is reached)
   - Hours tolerance: 99.5% to 102% (can go up to 2% over if needed for line balancing)
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
   - **Automatic**: Always generates plans for all possible days (no command-line argument needed)
   - **Sequential filling**: Each day is filled to 100% hours before moving to the next day
   - Each day must reach at least 99.5% of hours target to be considered complete
   - Days are filled sequentially (Day 1 to 100%, then Day 2 to 100%, etc.)
   - Remaining orders that don't fit into complete days go into "Remainder"
   - Days are labeled "Day 1", "Day 2", etc., with "Remainder" for overflow

## Constraints and Limits

### Hard Constraints (Must Not Exceed)

1. **Hours Limit**
   - BVI: 390 hours (must reach at least 99.5%, can go up to 102% if needed for line balancing)
   - Malosa: 90 hours (must reach at least 99.5%, can go up to 102% if needed for line balancing)
   - Primary constraint - **must reach at least 99.5%** before algorithm stops

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
   - Hours is prioritized (must hit at least 99.5% of target)
   - Qty and Picks balance around Hours constraint

3. **Difficulty Blending**
   - Orders are categorized as Easy, Medium, or Difficult based on efficiency metrics
   - Efficiency metrics: Qty/Hr, Picks/Hr, Picks/Qty (loaded from CSV or calculated)
   - Target mix: ~30% Easy, ~40% Medium, ~30% Difficult orders
   - Algorithm blends difficult orders with easy ones to process challenging work
   - Prevents leaving all difficult orders for the Remainder

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

5. **Difficulty Blending**
   - Orders categorized by efficiency: Easy (high Qty/Hr, low Picks/Qty), Medium, Difficult (low Qty/Hr or high Picks/Qty)
   - Strong bonus for difficult orders when underrepresented (helps process challenging work)
   - Prefers medium orders as the "sweet spot"
   - Ensures a balanced mix rather than leaving all difficult orders for Remainder

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
   - **Efficiency Metrics**: Qty/Hr, Picks/Hr, Picks/Qty (loaded from CSV or calculated automatically)

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
   - **Automatic**: Always runs in multi-day mode, determines maximum possible days automatically
   - **Sequential filling**: Fills each day to 100% hours before moving to the next
   - Each day must reach at least 99.5% of hours target to be considered complete
   - Algorithm continues until no more complete days can be created
   - Stops creating new days when can't fill a day to at least 99.5% of hours target
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

# Run optimizer (automatically generates all possible days)
python daily_plan_optimizer.py
```

**Note**: The optimizer now automatically determines and generates plans for all possible days. No command-line arguments needed.

## Notes

- The optimizer uses a greedy algorithm with scoring
- **Hours target is the absolute priority**: Algorithm will not stop until at least 99.5% of hours target is reached
- Multi-day planning fills days sequentially to 100% before moving to the next day
- Difficulty blending ensures challenging orders are processed throughout the plan, not just left for Remainder
- Line balance and order count are secondary priorities
- Due dates are used for tie-breaking and initial ordering
- The system is designed to be flexible with Qty/Picks to ensure hours target is met
- Efficiency metrics (Qty/Hr, Picks/Hr, Picks/Qty) are used to categorize orders and create balanced difficulty mixes
