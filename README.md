# **Production Forecasting & Cost Optimization Tool**  
*Advanced Excel VBA Automation · ERP-style Simulation · Data Visualization*  

---

## 📑 Table of Contents  
- [Project Description](#-project-description)  
- [Repository Contents](#-repository-contents)  
- [Project Output](#-project-output)  
- [Key Business Insights](#-key-business-insights)  
- [Recommendations](#-recommendations)  
- [Known Issues (Bug Review)](#-known-issues--bugs-identified)  
- [Future Enhancements](#-future-enhancements)  
- [Tools & Skills Used](#-tools--skills-used)  

---

## Project Description  
This is an **advanced Excel VBA automation project** designed to simulate a **3-step manufacturing workflow** (Cutting → Assembly → Finishing) and support **production planning, forecasting, and cost optimization**.  

### Key Features  
- Automates creation of **Calendar, Forecast, Cost Summary, and Dashboard**  
- Simulates ERP-style planning with **batching, onboarding limits, and milestone bonuses**  
- Provides **decision-support** by combining operational scheduling with cost analysis  
- Demonstrates how **Excel VBA can replace manual processes** and reduce bottlenecks  

---

## 📂 Repository Contents  
- **[vba_project_test.xlsx](vba_project_test.xlsx)** → Input dataset (includes HR-style instructions & parameters)  
- **[vba_project_test_macro.xlsm](vba_project_test_macro.xlsm)** → Final macro-enabled Excel tool 
- **[macrologic.bas](macrologic.bas)** → Main macro logic  
- **[README.md](README.md)** → Project documentation  
- `LICENSE` → MIT License 

---

## Project Output

When the macro runs, it automatically generates **four outputs** in the workbook:

### 📅 Calendar
- Lists each client, their batch, and the exact dates for the three production steps (Cutting, Assembly, Finishing).  
- Helps managers see **who is scheduled, and when**.  

### 📊 Forecast
- Shows monthly totals: **how many steps** are needed and **how many clients** are completed cumulatively.  
- Provides a **capacity planning view**.  

### 💰 Cost Summary
- Calculates **monthly production costs, handling costs, and total costs**.  
- Flags when the **bonus** (after 10 clients completed) is triggered.  

### 📈 Dashboard
- A chart of **Monthly Production Steps**.  
- Gives a quick **visual overview of workload across months**.  


---

## Key Business Insights  
- **Workload Balancing** → Initial month has heavy load, stabilizes afterward  
- **Cost Dynamics** → Costs increase with production steps, bonus reduces net cost after 10 clients complete  
- **Predictability** → 2-week intervals provide clear handoffs for planning  
- **Decision-Support** → Tool allows testing different batch sizes, onboarding limits, and cost structures  

---

## Recommendations  
- Use this tool for **scenario testing** (e.g., onboarding 6 vs. 10 clients/month)  
- Add a **cumulative clients vs. target chart** to highlight progress  
- Integrate outputs into **monthly management reporting**  

---

## Known Issues & Bugs (Identified)  
- **Date Calculation** → Uses `"ww"` (weeks). Safer to use `"d"` for exact day offsets  
- **Hardcoded Clients** → Fixed at 18; should be dynamic  
- **MaxClientsPerMonth** → Input exists, but not implemented in scheduling logic  
- **Bonus Logic** → Triggers once; doesn’t handle recurring bonuses  
- **Array Performance** → `ReDim Preserve` is slow for large datasets  

---

## Future Enhancements  
- Add **error handling** for missing or invalid inputs  
- Enforce **Max Clients per Month** constraint  
- Make **Total Clients dynamic** (remove hardcoding)  
- Add **salary increments** for workforce cost modeling  
- Provide a **user-friendly button or form** for one-click forecasting  
- Connect to **live data sources** (SQL, SharePoint) for automation  

---

## Tools & Skills Used  
- **Excel VBA** → Advanced macros, dynamic arrays, automation  
- **Data Modeling** → Batch scheduling, cost aggregation  
- **Visualization** → Dashboard charts for planning  
- **Business Analytics** → Production planning & cost control  

---

## License  
This project is licensed under the **MIT License**. See the LICENSE file for details.  
# Production_Forecasting_and_Cost_Optimization_Tool
