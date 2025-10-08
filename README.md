# Excel Assignment

1. Data Cleaning

- Created a table named 'Cleaned' to prepare the data for analysis.
- Checked for duplicate rows using “Remove Duplicates” under the Data tab — no duplicates found.
- Ensured correct data formats:
   • Date columns in Date format.
   • Numeric fields in Number format.
- Fixed missing data:
   • Replaced missing text fields (City, Salesperson, Channel) with “Unknown”.
   • Replaced missing numbers with 0.
- Used Conditional Formatting to highlight in Red colour:
   • UnitPrice < 0 (highlighted).
   • DiscountPct > 30% (highlighted).
- Fixed dates so that RequiredDate is not earlier than OrderDate using:
   =MIN(B2,E2) → Corrected OrderDate
   =MAX(B2,E2) → Corrected RequiredDate
- Added a new column LeadTimeDays = RequiredDate - OrderDate.

2. Calculated Columns

Created calculated columns to compute financial metrics using Excel formulas provided in the assignment.

3. Standardized Dimensions

- Month: =TEXT(B2,"MMM-YYYY")
- Quarter: ="Q"&ROUNDUP(MONTH(B2)/3,0)&"-"&YEAR(B2)
- Price Band (Low/Medium/High):
   =IF(U3<=QUARTILE($U:$U,1),"Low",IF(U3<=QUARTILE($U:$U,3),"Medium","High"))
- Region hierarchy was already provided in the dataset.

4. Cohort Analysis

Created a helper column 'CohortMonth' to identify the first month each country made a sale:
   =TEXT(I3,"MMM")
Used a PivotTable to group sales by Country and Month and calculate Monthly Revenue starting from each country’s first sales month.
5. I created 3 helper columns on my cleaned data named Rev per category, Cummulative Revenue and Cummulative %

6. Pivot Table Helper Columns

Created helper columns:
- Revenue per Order
- Orders per Month
- Average Orders per Month

Added OrderDate twice in the PivotTable (once as MIN and once as MAX). Applied Conditional Formatting to highlight:
- Top 3 performers (Green)
- Bottom 3 performers (Red)

8. Delivery Insights

Japan, Germany, and USA show the best on-time delivery rates (75–82%).
Nigeria, Kenya, and South Africa show the lowest (45–52%).
Accessories and Phones ship fastest; Networking and Printers are slowest.

9. Discount Insights

Three salespeople consistently exceed 20% discount thresholds: C. Otieno, J. Njeri, and A. Patel.
Extreme discounts (>30%) were also identified, with the highest being 56.5% by M. Rossi.

10. Helper Column Challenges - I got stuck

Attempted to create additional helper columns but encountered technical challenges in linking them correctly in the PivotTable.

12. Dashboard Insights

- Africa Sells Most, Profits Least: Highest sales volume, but lowest profit margins.
- Brazil & India Love Discounts: High discount rates hurt profitability.
- Europe Has the Biggest Spenders: Germany and France have the highest average order value.
- Online Shopping Hurts Retail: Online sales growth reduces retail store sales.
- Laptops Lead Sales: Top 3 products are laptops.
- Africa Has Slowest Delivery: Less than half of orders arrive within 7 days.

