# Toyota PESTLE Analytics Dashboard

## Dashboard Preview

An industry-level Power BI dashboard built to analyze Toyota Motor Corporation's performance through a PESTLE lens (Political, Economic, Social, Technological, Legal, Environmental), focused on the Economic, Technological, Social, and Environmental dimensions. The dashboard combines stock performance, production/sales trends, R&D efficiency, and workforce/sustainability metrics into a single executive-level view, mirroring the kind of cross-functional case study analysis used in corporate strategy and FP&A teams.

## Dashboard Pages

### Page 1 — Economic Factors
Tracks Toyota's market and financial performance: trading volume vs. market price trend, yearly average stock open price, average close vs. adjusted close price over time, plus KPI cards for total revenue, total capital expenditure, max average open, and max close price. Filterable by year.

**Key Insight:** Revenue grew from ¥31,379.5B in 2022 to ¥37,154.2B in 2023 — an 18.4% YoY increase — while operating margin held at 7.3%, showing top-line growth outpaced profitability gains.
<img width="1232" height="697" alt="Economic Factors" src="https://github.com/user-attachments/assets/d6f1616d-62c0-4581-8178-2c60fd3ded42" />

### Page 2 — Economic Factors II
Compares vehicle production against units sold (2011–2023), production share by continent on a map, and YoY production/unit sales growth.

**Key Insight:** Asia accounts for the majority of global production (6.01M of 9.40M units in 2023, ~64%), with North America a distant second at 1.84M — production growth (+6.3% YoY) outpaced sales growth (+1.0% YoY) in 2023, pointing to a building inventory gap.
<img width="1234" height="695" alt="Economic Factors 2" src="https://github.com/user-attachments/assets/54ce7f1c-1b6a-440c-87ed-8a78b8be5cbe" />

### Page 3 — Technological Factor
Examines R&D investment efficiency: R&D vs. operating returns scatter plot, R&D vs. revenue performance over time, and a decomposition tree breaking down operating income drivers by year, region, and expense category.

**Key Insight:** R&D Efficiency (Operating Income / R&D Expense) sat at 2.19x in 2023, while the Innovation Productivity Index (Operating Income per ¥100 of R&D + CapEx) was 95.7 — both measures of how much operating return each R&D yen generates.
<img width="1234" height="693" alt="Technological Factor" src="https://github.com/user-attachments/assets/dbc94bcd-8f9e-49bc-9232-39c2dfce7590" />

### Page 4 — Social and Environmental Factor
Covers gender distribution of new hires, annual social investment expenditure, and revenue vs. waste production trends.

**Key Insight:** Female representation in new hires rose to 43.6% in 2023 (723 of 1,657 hires), the highest in the dataset, even as total waste production ticked up to 29.0 thousand/ton from a multi-year low of 23.0 thousand/ton in 2022.
<img width="1231" height="694" alt="Social and Environmental Factor" src="https://github.com/user-attachments/assets/a96843d5-c0da-481f-b518-728474c59388" />


## Technical Architecture

### Data Sources
- **Annual Report** — stock price (open/close/high/low/adjusted close, volume), revenue, operating income, R&D expenses, capital expenditures, ROE (2001–2023)
- **Sales & Production** — total units sold and total units produced by continent (2011–2023)
- **Employee, Social & Environmental** — new hires by gender, total waste production, annual social contribution spending (2011–2023)

Sourced from a mix of company annual report figures and stock market data, consolidated into Excel workbooks and loaded via Power Query.

### Data Model
Three-table model, each loaded independently via Power Query (`Excel.Workbook` source) and connected on `Year`:
- `Annual Report`
- `Sales & Production`
- `Employee, Social & Enviromental`

<img width="1254" height="524" alt="Relation Architecture" src="https://github.com/user-attachments/assets/214ecf93-33cb-475a-aff0-229f0a56e2d0" />

Power Query transformations include:
- Promoted headers and type changes for all numeric columns
- Continent cleanup in `Sales & Production`: `"Latin America"` remapped to `"South America"`, `"Oceanic"` remapped to `"Australia"`
- `"N/A"` text values in `Total Production` replaced and column re-typed to whole number
- Blank-row removal in the Employee/Social/Environmental table

### DAX Measures
11 measures spanning financial growth, efficiency, and sustainability metrics:

**Growth Metrics**
```dax
YOY Revenue Growth (%) =
VAR CurrentYear = MAX('Annual Report'[Year])
VAR PrevRevenue = CALCULATE(
    SUM('Annual Report'[Revenue (in Billions Yen)]),
    FILTER('Annual Report', 'Annual Report'[Year] = CurrentYear - 1)
)
RETURN IF(
    ISBLANK(PrevRevenue), BLANK(),
    DIVIDE(SUM('Annual Report'[Revenue (in Billions Yen)]) - PrevRevenue, PrevRevenue)
)
```
Similar YoY pattern applied to `YOY Unit Sales Growth (%)` and `YOY Production Growth (%)`.

**Efficiency Metrics**
```dax
Operating Margin (%) = DIVIDE(SUM('Annual Report'[Operating income (in Billions Yen)]), SUM('Annual Report'[Revenue (in Billions Yen)])) * 100
R&D Efficiency = DIVIDE(SUM('Annual Report'[Operating income (in Billions Yen)]), SUM('Annual Report'[R&D Expenses(in Billions Yen)]))
Innovation Productivity Index = DIVIDE(SUM('Annual Report'[Operating income (in Billions Yen)]) * 100, SUM('Annual Report'[R&D Expenses(in Billions Yen)]) + SUM('Annual Report'[Capital Expenditures (billions Yen)]))
CapEx to Revenue Ratio = DIVIDE(SUM('Annual Report'[Capital Expenditures (billions Yen)]), SUM('Annual Report'[Revenue (in Billions Yen)]))
Sales-to-Production Efficiency = DIVIDE(SUM('Sales & Production'[Total Unit Sold]), SUM('Sales & Production'[Total Production]))
```

**Social & Environmental Metrics**
```dax
Total New Hires = SUM('Employee, Social & Enviromental'[Male (New Hired) TMC]) + SUM('Employee, Social & Enviromental'[Female (New Hired) TMC])
Cumulative Waste = CALCULATE(
    SUM('Employee, Social & Enviromental'[Total_Waste_Production thousand/ton (TMC)]),
    FILTER(ALLSELECTED('Employee, Social & Enviromental'), 'Employee, Social & Enviromental'[Year] <= MAX('Employee, Social & Enviromental'[Year]))
)
Lowest Waste Year = CALCULATE(
    MIN('Employee, Social & Enviromental'[Year]),
    FILTER('Employee, Social & Enviromental', 'Employee, Social & Enviromental'[Total_Waste_Production thousand/ton (TMC)] = MIN('Employee, Social & Enviromental'[Total_Waste_Production thousand/ton (TMC)]))
)
```

### Visuals Used
Line charts, line/stacked-column combo charts, clustered column charts, pie charts, scatter chart, decomposition tree, map (continent-level), KPI cards, slicers, and textboxes.

## Key Findings Summary

| Finding | Metric | Business Implication |
|---|---|---|
| Revenue Acceleration | +18.4% YoY revenue growth (2022→2023) | Strong top-line recovery, though margin growth lagged |
| Production-Sales Gap | Production +6.3% vs. Sales +1.0% YoY (2023) | Inventory build-up risk if demand doesn't catch up |
| Asia Concentration | Asia = ~64% of global production (2023) | Heavy geographic concentration risk in manufacturing footprint |
| R&D Payoff | R&D Efficiency 2.19x; Innovation Productivity Index 95.7 (2023) | Each R&D yen still converts to meaningful operating return |
| Hiring Diversity | Female new hires at 43.6% (2023), highest on record | Steady progress on workforce gender balance |
| Waste Uptick | Waste rose to 29.0k tons (2023) from a low of 23.0k tons (2022) | Environmental performance regressed after years of improvement |

## Tools & Technologies

| Tool | Usage |
|---|---|
| Power BI Desktop | Dashboard development, DAX measures, visuals |
| Power Query (M) | Data transformation, continent remapping, type cleansing |
| DAX | 11 measures across growth, efficiency, and sustainability |
| Excel | Source data consolidation (Book1.xlsx, Book1 updated.xlsx) |

## Project Structure
```
Toyota-PESTLE-Analytics/
│
├── Final_project_Group6.pbix   # Main Power BI dashboard file
├── README.md                   # Project documentation
└── screenshots/
    ├── page1_economic_factors.png
    ├── page2_economic_factors_ii.png
    ├── page3_technological_factor.png
    └── page4_social_environmental_factor.png
```

## How to Open
1. Download and install Power BI Desktop (free)
2. Clone or download this repository
3. Open `Final_project_Group6.pbix` in Power BI Desktop
4. All data is embedded — no external connections required

## About
Built as a portfolio project applying PESTLE analysis to Toyota Motor Corporation, combining financial, operational, and sustainability data into a single Power BI case study.

**Author:** Roshan
