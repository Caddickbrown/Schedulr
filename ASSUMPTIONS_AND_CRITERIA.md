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
   - Orders are categorized as Easy, Medium, or Hard based on efficiency metrics
   - Efficiency metrics: Qty/Hr, Picks/Hr, Picks/Qty (loaded from template or calculated)
   - **Goal: Each day should have a balanced AVERAGE difficulty**
   - When a Hard order is included, pair it with Easy orders to balance out
   - The exact percentage mix doesn't matter - what matters is the average
   - Prevents leaving all difficult orders for the Remainder
   - Remainder day must also maintain balanced difficulty

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
   - Orders categorized by efficiency: Easy (high Qty/Hr, low Picks/Qty), Medium, Hard (low Qty/Hr or high Picks/Qty)
   - **Goal: Each day should have the same AVERAGE difficulty**
   - When Hard orders are included, balance with Easy orders
   - The exact percentage mix (e.g., 30/40/30) doesn't matter - the average does
   - Ensures a balanced mix across ALL days including Remainder

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

2. **Multi-Day Planning (Multi-Round Leveling)**
   - **Phase 1: Round-Robin Distribution**
     - Orders are sorted by start date (earlier first)
     - Distributed across days like dealing cards (order 1→Day1, order 2→Day2, etc.)
     - This ensures approximately equal order counts, picks, difficulty, and lines per day
   - **Phase 2: Hours Balancing**
     - Maximizes hours on Days 1 to N-1 (target 100% each)
     - Day N is designated as "Remainder" and can have less hours
     - Uses swap-based balancing to preserve order count balance
     - When swapping with remainder, prefers swaps that keep remainder balanced
   - **Phase 3: Order Count Check**
     - Verifies order counts are balanced (spread ≤ 3)
     - Does not modify if already balanced
   - **Phase 4: Difficulty Balancing**
     - Ensures each day (including remainder) has similar average difficulty
     - Swaps Hard orders for Easy orders between days to balance
     - Target: all days within 15% of target average difficulty
   - **Result**: Days 1 to N-1 at ~100% hours, Remainder at lower hours but still balanced
   - Days are labeled "Day 1", "Day 2", etc.

3. **Balance Targets**
   - **Hours**: Each day targets equal share of total hours (max hours limit per day)
   - **Order counts**: Each day gets roughly equal number of orders
   - **Picks/Qty**: Naturally balanced by order count balancing

4. **Line Distribution**
   - Line distribution tracked per day
   - Balanced across days by the round-robin distribution

5. **Offline Jobs**
   - Tracks count of offline orders per day
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

2. **Remainder Day (Last Day)**
   - Hours target doesn't apply (can be less than 100%)
   - **Order count is PROPORTIONAL to hours:**
     - If remainder is at 80% hours, it gets ~80% of the orders
     - This leaves room for future orders to top it up
     - When topped up to 100%, order count will match other days
   - **Must remain BALANCED in all other dimensions:**
     - Balanced difficulty mix (Easy/Medium/Hard) - same average as other days
     - Balanced picks (proportional, not all high-pick or all low-pick orders)
     - Balanced line distribution
   - This ensures when new orders arrive to fill the remainder, the day is not unbalanced
   - Offline limit still applies

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
