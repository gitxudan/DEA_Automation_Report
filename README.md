# DEA_Automation_Report
Python automation report for DEA analysis

This project implements **Data Envelopment Analysis (DEA)** in Python to evaluate the relative efficiency of business branches.  
It is a modernized version of the original Excel Solver approach, fully automated with **Pandas, PuLP, and OpenPyXL**.

---

## 📌 Business Background
Queens Financial Alliance (QFA) operates 20+ branches across North America (including 4 in GTA).  
Management needed a **data-driven tool** to:
- Compare branch efficiency using multiple input and output measures.
- Establish performance benchmarks for continuous improvement.
- Support resource reallocation decisions.
- Provide clear targets for underperforming branches.

---

## 🔎 Why DEA?
**Data Envelopment Analysis (DEA)** is a non-parametric linear programming method that:
- Integrates **multiple inputs and outputs** in a single efficiency model.
- Identifies “best practice” branches on the **efficiency frontier**.
- Benchmarks underperforming branches against efficient peers.
- Requires no predefined functional form, making it flexible for real-world data.

---

## 📊 Inputs & Outputs (Example)

| Inputs              | Outputs                |
|---------------------|------------------------|
| Number of Employees | Revenue                |
| Operational Expense | Number of Customers    |
| Marketing Budget    | Customer Satisfaction  |


Each branch’s efficiency score reflects its ability to **transform resources (inputs) into results (outputs).**

---

## ⚙️ Python Implementation
Key features of this automation:
- **Data ingestion**: Reads input data from `.txt` or `.csv` files.
- **Optimization**: Uses `PuLP` linear programming to calculate efficiency scores.
- **Normalization & Constraints**:
  - Input normalization (sum of inputs = 1).
  - Input ≥ Output constraints to ensure realism.
- **Reporting**: Generates Excel reports (`.xlsx`) with:
  - Original input/output data
  - Efficiency scores
  - Peer comparison and improvement targets
  - Visual charts for branch efficiency
 
---

## 🚀 How to Run
1. Clone the repository:
   ```bash
   git clone https://github.com/gitxudan/DEA_Automation_Report.git
   cd DEA_Automation_Report

📈 Example Output

Efficiency scores per branch (0–1 scale)

Identification of efficient vs. inefficient branches

Excel report with data, LP model results, and charts


📌 Notes

Original input data format follows DEA.txt convention:

First line: input variable names

Second line: output variable names

Each branch: name, input values, output values

Project is for educational and portfolio purposes.

🔗 Author

GitHub: gitxudan

Contact: xudanbetsy@gmail.com
