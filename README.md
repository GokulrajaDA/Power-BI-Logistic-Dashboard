# Vehicle Trip Analysis Dashboard

End-to-end analysis of Vehicle Trip and Vehicle Master data.
Project covers: **Data Cleaning in Excel** → **DAX Modeling in Power BI** → **Interactive Dashboard**

### Objective
Track delivery performance, fuel efficiency, maintenance cost, and route utilization for fleet of 7 vehicles across 50 trips.

---

## 📂 Dataset
**Source Tables:**
1. **Log** - Trip Data
   `Trip_ID, Vehicle_ID, Driver_ID, Origin, Destination, Distance, Fuel_Cons, Delivery_Status, Delivery_Date`
2. **Log 1** - Vehicle Master
   `Vehicle_ID, Vehicle_Type, Capacity, Maintenance`

**Relationship:** `Log 1[Vehicle_ID] 1:* Log[Vehicle_ID]`

---

## 🧹 1. Data Cleaning - Excel

All cleaning was done in `Cleaned_Trip_Data.xlsx`

| Step | Problem Found | Fix Applied |
| --- | --- | --- |
| **Missing Values** | Blank cells in `Distance`, `Fuel_Cons` | Replaced with `0` using `IF(ISBLANK())` |
| **NULL Values** | Empty `Delivery_Status` | Replaced NULL with `Unknown` |
| **Duplicates** | Checked `Trip_ID` | No duplicates found |
| **Date Type** | Mixed formats: `1/1/2023`, `27-01-2023` + `#######` error | 1. Increased column width 2. Converted to `Date` type `dd-mm-yyyy` 3. Standardized using `Text to Columns` |

**Excel Cleaning Formulas:**
```excel
=IF(ISBLANK([@Fuel_Cons]), 0, [@Fuel_Cons])
=IF([@Delivery_Status]="", "Unknown", [@Delivery_Status])

📊 2. Power BI - Data Modeling & DAX
Data Modeling
Loaded cleaned Excel into Power BI
Changed Delivery_Date data type to Date in Power Query
Created relationship between Log 1 and Log on Vehicle_ID

-- Count & Delivery Performance
Total Trips = COUNT('Log'[Trip_ID])

On-Time Deliveries = CALCULATE(COUNT('Log'[Trip_ID]), 'Log'[Delivery_Status] = "On-Time")

Late Deliveries = CALCULATE(COUNT('Log'[Trip_ID]), 'Log'[Delivery_Status] = "Late")

On-Time % = DIVIDE([On-Time Deliveries], [Total Trips])

Late % = DIVIDE([Late Deliveries], [Total Trips])

-- Distance & Fuel Metrics
Total Distance = SUM('Log'[Distance])
Avg Distance = AVERAGE('Log'[Distance])
Total Fuel Consumed = SUM('Log'[Fuel_Cons])
Avg Fuel Consumed = AVERAGE('Log'[Fuel_Cons])
Fuel Efficiency = DIVIDE([Total Distance], [Total Fuel Consumed]) // KM per Unit

-- Vehicle & Cost Metrics
Total Capacity = SUM('Log 1'[Capacity])
Total Maintenance Cost = SUM('Log 1'[Maintenance])
Avg Maintenance per Vehicle = AVERAGEX('Log 1', 'Log 1'[Maintenance])

-- Time Intelligence
Trips This Month = CALCULATE([Total Trips], DATESINPERIOD('Log'[Delivery_Date], TODAY(), -1, MONTH))
Late % Last 30 Days = CALCULATE([Late %], DATESINPERIOD('Log'[Delivery_Date], TODAY(), -30, DAY))

📈 3. Dashboard & Charts
Interactive dashboard built in Vehicle_Trip_Dashboard.pbixVisualPurposeFields Used1KPI CardsShow overall KPIs at top[Total Trips], [On-Time %], [Late %], [Fuel Efficiency]2Donut ChartDelivery Status distributionLegend: 'Log'[Delivery_Status], Values: [Total Trips]3Clustered Bar ChartLate Deliveries by VehicleAxis: 'Log 1'[Vehicle_ID], Values: [Late Deliveries]4Bar ChartTop 10 Busiest RoutesAxis: Route, Values: [Total Trips]5Scatter PlotDistance vs Fuel to find inefficiencyX: [Distance], Y: [Fuel_Cons], Legend: 'Log 1'[Vehicle_Type]6Column ChartMaintenance Cost by Vehicle TypeAxis: 'Log 1'[Vehicle_Type], Values: [Total Maintenance Cost]7Line ChartMonthly Trip TrendAxis: 'Log'[Delivery_Date] by Month, Values: [Total Trips]Slicers Used: Vehicle_Type, Origin, Destination, Delivery_Status
Theme: Dark theme with cards for KPI highlightingPart of this response isn't supported on this device yet. View the full response on your phone.
🔍 4. Key Insights
From the analysis of 50 trips:52% of deliveries were Late and 48% were On-TimeLongest Trip: T045 - Kolkata to Chennai: 1803 KM with highest fuel consumptionWorst Fuel Efficiency: T023 - 1729 KM with 197.55 fuel unitsHighest Maintenance Cost: Truck vehicle type contributed to 40% of total maintenanceMost Active Route: Hyderabad ↔ Delhi with 6 tripsAvg Distance per Trip: 842 KMOverall Fuel Efficiency: 8.7 KM per unit
