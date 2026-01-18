# ConnectRetail Profitability Analysis - Complete Documentation

## Table of Contents
1. [Project Overview](#project-overview)
2. [Data Preparation](#data-preparation)
3. [Analysis 1: Profitability](#analysis-1-profitability)
4. [Analysis 2: Discount Impact](#analysis-2-discount-impact)
5. [Analysis 3: Customer Segmentation](#analysis-3-customer-segmentation)
6. [Dashboard Development](#dashboard-development)
7. [Key Findings](#key-findings)
8. [Recommendations](#recommendations)
9. [Technical Challenges](#technical-challenges)
10. [Lessons Learned](#lessons-learned)

---

## Project Overview

### Business Context
ConnectRetail is a US-based retailer specializing in office supplies, furniture, and technology products. Despite 12% revenue growth, profit grew only 3%, indicating significant margin leakage requiring investigation.

### Objective
Conduct end-to-end profitability analysis to:
- Identify root causes of margin decline
- Quantify improvement opportunities
- Provide actionable recommendations with ROI projections
- Build executive dashboard for ongoing monitoring

### Constraints
**Primary Constraint:** Excel Web only (browser-based)
- Linux operating system (no Excel Desktop access)
- Dataset size limitations (optimal: 5K-15K rows)
- No Power Query, VBA, or advanced features
- Slower performance vs. desktop version

### Dataset Specifications
- **Source:** Kaggle Superstore Sales Dataset
- **Size:** 9,994 transactions
- **Time Period:** 2014-2017 (4 years)
- **Columns:** 21 fields including Order ID, Date, Customer, Product, Category, Sales, Profit, Discount
- **Grain:** One row per order line item

---

## Data Preparation

### Step 1: Dataset Selection Process

**Initial Attempt:**
- Brazilian E-Commerce dataset selected
- 90,200 rows after filtering
- **Issue:** Excel Web copy/paste failed at this volume
- **Decision:** Pivot to smaller dataset

**Final Selection:**
- Superstore Sales dataset
- 9,994 rows (perfect size for Excel Web)
- Included profit data (critical for margin analysis)
- Well-structured, minimal cleaning required

### Step 2: Workbook Structure

**Created workbook:** `ConnectRetail_Analysis.xlsx`

**Sheet structure:**
```
├── RawData (9,994 rows × 27 columns after enhancements)
├── 1_Profitability
├── 2_Discount_Analysis
├── 3_Customer_Segmentation
└── Executive_Dashboard
```

### Step 3: Data Transformation

**Converted to Table (Ctrl+T):**
- Enabled structured references: `[@Column_Name]`
- Auto-fill formulas capability
- Dynamic range management

**Created 6 Calculated Columns:**

**1. Profit_Margin_Pct**
```excel
=[@Profit]/[@Sales]
```
Purpose: Convert profit to percentage for analysis

**2. Order_Month**
```excel
=TEXT([@[Order Date]],"MMM-YYYY")
```
Purpose: Enable monthly trend analysis

**3. Order_Year**
```excel
=YEAR([@[Order Date]])
```
Purpose: Annual comparisons

**4. Discount_Band**
```excel
=IF([@Discount]=0,"No Discount",
  IF([@Discount]<=0.15,"Low (1-15%)",
  IF([@Discount]<=0.30,"Medium (16-30%)","High (31%+)")))
```
Purpose: Group discounts for pattern analysis

**5. Days_to_Ship**
```excel
=[@[Ship Date]]-[@[Order Date]]
```
Purpose: Operational efficiency metric

**6. Profit_per_Order**
```excel
=[@Profit]
```
Purpose: Facilitate averaging in pivot tables

### Step 4: Data Quality Verification

**Checks Performed:**
- ✅ No duplicate Order IDs (9,994 unique)
- ✅ No missing critical fields (0% null in key columns)
- ✅ Date range validated (2014-01-03 to 2017-12-30)
- ✅ Profit margins calculated correctly (spot-checked 50 rows)
- ✅ All categories present (Furniture, Office Supplies, Technology)

**Issues Found & Resolved:**
- None - dataset was remarkably clean

---

## Analysis 1: Profitability

### Objective
Identify which products/categories are profitable vs. unprofitable to guide portfolio decisions.

### Approach

**Step 1: Category-Level Analysis**

Created Pivot Table:
```
Rows: Category
Values:
  - Sum of Sales
  - Sum of Profit
  - Average of Profit_Margin_Pct
  - Count of Order ID (renamed "Order Count")
```

**Step 2: Manual Calculated Columns**

Added next to pivot:
```excel
Profit per Order = Total Profit / Order Count
Revenue per Order = Total Sales / Order Count
```

Excel Web Limitation: Calculated fields in pivots unreliable
Workaround: Manual formulas in adjacent columns

**Step 3: Sub-Category Deep-Dive**

Created hierarchical pivot:
```
Rows:
  - Category
  └─ Sub-Category (nested)
Values: Same as category analysis
```

Sorted by Profit (Smallest to Largest) to surface loss leaders

**Step 4: Performance Classification**

Created matrix:
```excel
Classification = 
IF(AND(Profit>10000, Margin>0.15), "Star",
IF(AND(Profit>5000, Margin>0.10), "Solid,
IF(Profit<0, "Exit", "Monitor")))
```

### Key Findings

**Category-Level Results:**

| Category | Revenue | Profit | Margin | Profit/Order | Classification |
|----------|---------|--------|--------|--------------|----------------|
| Technology | $836K | $145K | 15.6% | $78.76 | Star |
| Office Supplies | $719K | $122K | 13.8% | $20.33 | Solid |
| Furniture | $742K | $18K | 3.9% | $8.70 |  Problem |

**Sub-Category Discoveries:**

**Worst Performers:**
1. Tables: -$17,725 profit (-14.7% margin) - **ACTIVELY DESTROYING VALUE**
2. Bookcases: -$3472.556 profit (-12.66% margin) - essentially break-even
3. Supplies: -$1189 profit (11.20% margin) - below company average

**Best Performers:**
1. Copiers: $55,618 profit (31.72% margin) - premium product
2. Phones: $44,515 profit (11.92% margin) - volume driver
3. Accessories: $41,936 profit (21.82% margin) - high-margin add-ons

### Insight Development

**Key Insight:**
> "Furniture generates $742K in revenue (32% of total) but delivers only $18.5K in profit, resulting in a 3.9% margin."

**Business Impact Calculation:**
```
If Furniture merely matched Office Supplies’ 13.8% margin, incremental profit would be substantial:
. Current Furniture profit: $18.5K
. Potential profit at 13.8% margin: $102.4K
($742K × 13.8%)
. Missed opportunity: +$83.9K
That represents a +29% increase in total company profit ($83.9K ÷ $286.4K).

```

### Recommendations

**Immediate (Week 1-2):**
- EXIT Binders completely (-$30K profit drain)
- Increase Tables pricing 30%+ or discontinue
- Phase out Bookcases and Machines

**REDIRECT: Shift sales/marketing investment to:**
  - Copiers (31.7% margin, highest profit density)
  - Paper (42.6% margin, volume + margin)
  - Accessories (21.8% margin, scalable)

**EXPECTED IMPACT:**
  - Eliminating 4 unprofitable categories: +$55K
  - Redirecting to high-margin products: +$70K
  - TOTAL PROFIT INCREASE: +$125K (+44% profit growth)
  - Profit growth: ~+40% without increasing revenue
  **Total Profitability Opportunity: $125K annual**

### Visualizations Created

1. **Revenue vs Profit Bar Chart** (Clustered Column)
   - X-axis: Categories
   - Y-axis: Dollar amounts
   - Legend: Sales (Blue), Profit (Green)

2. **Profit Margin Column Chart**
   - X-axis: Categories
   - Y-axis: Margin %
   - Color-coded: Red (Furniture), Green (Others)

3. **Sub-Category Heatmap**
   - Applied conditional formatting (color scale)
   - Red = Negative, Yellow = Low positive, Green = High

---

## Analysis 2: Discount Impact

### Objective
Determine if discounting drives profitable growth or destroys margins.

### Approach

**Step 1: Discount Band Performance**

Created pivot table:
```
Rows: Discount_Band
Values:
  - Count of Orders
  - Sum of Sales
  - Sum of Profit
  - Average of Profit_Margin_Pct
```

Added calculated columns:
```excel
Avg Order Value = Total Sales / Order Count
Profit per Order = Total Profit / Order Count
```

**Step 2: Category × Discount Analysis**

Created cross-tabulation pivot:
```
Rows: Category, Discount_Band (nested)
Values: Same as Step 1
```

**Step 3: Damage Assessment**

Manual calculation table:
```
Medium Discount Orders: 4,038
Current Loss: -$57,313
If converted to No Discount (at 29-37% margins):
  Potential Profit: $219,491
  Opportunity Cost: $276,804
```

**Step 4: Scenario Modeling**

Created 4 scenarios:
1. Current State (baseline)
2. Eliminate all medium+ discounts
3. Convert medium → low discounts
4. Zero discounts (aggressive)

Each with conversion rate assumptions and profit projections

### Key Findings

**Discount Band Results:**

| Band | Orders | Revenue | Profit | Avg Margin | Profit/Order | 
|------|--------|---------|--------|------------|--------------|
| No Discount | 5,897 | $1.35M | $216K | 16.0% | $36.57 |
| Low (1-15%) | 1,852 | $427K | $67K | 15.8% | $36.33 | 
| Medium (16-30%) | 1,664 | $407K | $48K | 11.8% | $28.73 | 

**SHOCKING Discovery:**
> "We literally paid customers an average of $76.39 per order to take heavily discounted products. Total loss from 581 orders: $44,385."

**Category-Specific Findings:**

**Furniture + Medium Discount:**
- 1,157 orders
- -14.86% margin
- Lost $48,211
- **Cannot afford ANY significant discounting**

**Office Supplies + Medium Discount:**
- 2,881 orders
- -11.14% margin  
- Lost $9,101
- **Discounting kills otherwise healthy margins**

**Technology + Medium Discount:**
- 1,012 orders
- +4.62% margin
- Made $12,274
- **Only category that survives discounting**

**Anomaly Identified:**
Low discounts (1-15%) barely used:
- Furniture: Only 128 orders (6% of category)
- Office: Only 16 orders (0.3%!)
- Technology: Only 2 orders (0.1%!)

**Hypothesis:** Heavy discounts may generate revenue but kill profit

### Insight Development

**Key Insight:**
> "The 'sweet spot' discount range (16-30%) is actually a profit death zone for 2 out of 3 categories. Combined with high discounts, we're giving away $276K annually through improper pricing strategies."

**Damage Calculation:**
```
Medium Discounts (5050 orders):
  Current Loss: -$45,039
```

### Recommendations

**IMMEDIATE (Week 1):**
1. **Implement Category-Specific Discount Caps**
   - Furniture: 5% maximum
   - Office Supplies: 10% maximum
   - Technology: 20% maximum
   - Require VP approval for exceptions
   - Expected Impact: +$180K annual

2. **Zero Discounts on Loss Leaders**
   - Tables, Bookcases: 0% discount allowed
   - Already unprofitable at full price
   - Expected Impact: Stop $20K annual bleeding

**SHORT-TERM (Month 1-2):**
3. **Sales Team Training**
   - Teach tiered discount strategy (5%, 10%, 15% options)
   - Shift 30% of medium discounts to low
   - Expected Impact: +$80K annual

4. **Revise Sales Compensation**
   - Tie commission to profit, not revenue
   - Penalize margin-killing discounts
   - Expected Impact: Behavioral change (unmeasured but critical)

**Total Discount Optimization Opportunity: $276K annual**

### Visualizations Created

1. **Discount Performance Bar Chart**
   - Color-coded by band (Green → Red spectrum)
   - Shows margin deterioration clearly

2. **Category × Discount Heatmap**
   - Conditional formatting (3-color scale)
   - Red cells immediately show problem areas
   - **Most impactful visual in entire project**

3. **Before/After Comparison Chart**
   - Side-by-side bars
   - No Discount vs Medium Discount margins
   - Dramatic visual impact

---

## Analysis 3: Customer Segmentation

### Objective
Identify high-value customers vs. profit drains to enable targeted retention strategies.

### Approach

**Step 1: Customer-Level Aggregation**

Created pivot table:
```
Rows: Customer Name
Values:
  - Count of Orders
  - Sum of Sales (renamed "Total Revenue")
  - Sum of Profit (renamed "Total Profit")
  - Average of Profit_Margin_Pct
```

Sorted by Total Profit (Largest to Smallest)

**Step 2: Calculated Metrics**

Added columns:
```excel
Revenue per Order = Total Revenue / Order Count
Profit per Order = Total Profit / Order Count
```

**Step 3: Segmentation Logic**

Copied pivot as values, added formula in new column:
```excel
Segment = 
IF(Total_Profit<0, "Unprofitable",
IF(Total_Profit>=PERCENTILE($Profit_Range, 0.9), "VIP (Top 10%)",
IF(Total_Profit>=PERCENTILE($Profit_Range, 0.7), "High Value (70-90%)",
IF(Total_Profit>=PERCENTILE($Profit_Range, 0.3), "Medium Value (30-70%)",
"Low Value (Bottom 30%)"))))
```

**Step 4: Behavioral Classification**

Added second formula:
```excel
Behavioral_Flag = 
IF(Order_Count=1, "One-Time Buyer",
IF(Order_Count>=5, "Loyal Customer ",
"Occasional Buyer"))
```

**Step 5: Segment Summary Analysis**

Created manual summary table using COUNTIF/SUMIF:
```excel
VIP Count: =COUNTIF(Segment_Column, "VIP (Top 10%)")
VIP Profit: =SUMIF(Segment_Column, "VIP (Top 10%)", Profit_Column)
VIP % of Total: =VIP_Profit / Total_Profit
```

Repeated for all 5 segments

**Step 6: Customer Lifetime Value (CLV) Projection**

Formula structure:
```excel
Historical CLV = Avg Orders per Customer × Avg Profit per Order
Projected 3-Year CLV = Historical CLV × 3
Segment Total Value = Customer Count × Projected CLV
```

Calculated for VIP, High Value, and Medium segments

### Key Findings

**Segment Distribution:**

| Segment | Customers | Total Revenue | Total Profit | % Profit | Avg Profit/Customer |
|---------|-----------|---------|--------------|----------|---------------------|
| VIP | 80 | 177K | $145K | 18% | $637 |
| High Value | 158 | 475K | $95K | 33% | $605 |
| Medium | 317 | 991K | $136K | 48% | $430 |
| Low | 59 | 178K | $1,217 | 0% | $21 |
| Unprofitable | 179 | 478K | $2K | 1% | $14 |

**Pareto Principle Confirmed:**
> "Top 10% of customers (80 customers):
   • Generate $177244 in revenue 
   • Deliver $50994 in profit (18% of total profit)
   • Average $637 profit per customer
   • These are your GOLD customers"

**Behavioral Insights:**

| Behavior | Count | Avg Profit | Avg Orders |
|----------|-------|------------|------------|
| One-Time Buyer | 5 | $128 | 1.0 |
| Occasional Repeat | 51 | $3,758 | 2-4 |
| Loyal Repeat (5+) | 737 | $282,509 | 8+ |

**Calculation:**
```
Loyal customers are 6.3× more valuable than one-time buyers
($950 / $150 = 6.33)
```

**Risk Assessment:**
```
 Losing 8 VIPs wipes out more profit than: Acquiring dozens of low- or medium-value customers, We will lose about 5096 dollars annually
```

### Insight Development

**Key Insights:**

1. **Extreme Concentration Risk**
> "Losing 8 VIPs wipes out more profit than: Acquiring dozens of low- or medium-value customers."

2. **o NOT over-invest in VIP acquisition**
> "VIPs are not created cheaply"
> "They emerge through relationship depth and order consistency"

3. **Retention Multiplier Effect**
> "Protect VIPs (highest profit density)"
> "Develop Medium → High (largest total value pool)"
> "Ignore Low / Unprofitable unless cost structure changes"

### Recommendations

**IMMEDIATE (Week 1-2):**
1. **Launch VIP Protection Program**
   - Dedicated account manager for each of 80 VIPs
   - Quarterly business reviews
   - Exclusive product access
   - 5% permanent pricing (vs ad-hoc heavy discounts)
   - Investment: $15K/year
   - ROI: Protect $436K in 3-year value = 29:1 ROI

**SHORT-TERM (Month 1-3):**
2. **High Value Upgrade Program**
   - Monthly personalized outreach to 158 high-value customers
   - Loyalty points (earn 2%, redeem at 5%)
   - Goal: Convert 20% to VIP status within 12 months
   - Expected Impact: +$25K annual profit

3. **Fix or Exit Unprofitable Segment**
   - Implement minimum order value ($150-200)
   - Remove/minimize discounts
   - Surcharge for small orders
   - If no improvement: Gracefully exit relationship
   - Expected Impact: +$6K recovery (stop bleeding)

**LONG-TERM (Quarter 2+):**
4. **One-Time Buyer Conversion Campaign**
   - Automated email sequence (30, 60, 90 days post-purchase)
   - 10% "welcome back" offer
   - Goal: 20% conversion rate
   - Expected Impact: +$14K annual

5. **Churn Prevention System**
   - Flag customers with no order in 180 days
   - Automated win-back campaign
   - Personal outreach for VIPs
   - Expected Impact: Reduce churn 15%, protect $50K+ value

**Total Customer Program Opportunity: $100K+ over 3 years**

### Visualizations Created

1. **Segment Distribution Donut Chart**
   - Shows % of profit by segment
   - Visual reinforcement of 50% concentration

2. **Customer Count vs Profit Stacked Bar**
   - 100% stacked bars side-by-side
   - Shows imbalance: 10% customers = 50% profit

3. **Segment Summary Table with Data Bars**
   - Conditional formatting for quick scanning
   - Color scales on profit columns

---

## Dashboard Development

### Design Philosophy

**Principles:**
1. **Executive-First:** C-suite should understand in 30 seconds
2. **Action-Oriented:** Every insight links to a recommendation
3. **Interactive:** Slicers enable self-service exploration
4. **Visual Hierarchy:** Most important info at top

### Layout Structure
```
┌─────────────────────────────────────────────────┐
│  Row 1-4: Header & Title                        │
├─────────────────────────────────────────────────┤
│  Row 6-11: KPI Scorecard (5 cards)              │
├─────────────────────────────────────────────────┤
│  Row 13-28: Profitability Section               │
│    - Category chart                             │
│    - Key finding box                            │
├─────────────────────────────────────────────────┤
│  Row 30-45: Discount Impact Section             │
│    - Heatmap                                    │
│    - Recommendation box                         │
├─────────────────────────────────────────────────┤
│  Row 47-62: Customer Segmentation               │
│    - Segment table                              │
│    - CLV summary                                │
├─────────────────────────────────────────────────┤
│  Row 64-70: Action Plan                         │
└─────────────────────────────────────────────────┘
```

### KPI Scorecard Build

**Card Structure (each card: 3 columns × 6 rows):**
- Row 1: Header (merged, 10pt, bold, dark gray)
- Rows 2-4: Big Number (merged, 32pt, bold, blue)
- Rows 5-6: Context/Trend (merged, 10pt, regular)

**5 Cards Created:**

**Card 1: Total Revenue**
```excel
=TEXT('1_Profitability'!D4,"$#,##0K")
Context: ="↑ " & TEXT((D4-2000000)/2000000,"0%") & " vs Target"
```

**Card 2: Total Profit**
```excel
=TEXT('1_Profitability'!E4,"$#,##0K")
Context: ="
