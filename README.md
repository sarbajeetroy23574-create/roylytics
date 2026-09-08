# Payroll Audit Case Study - Found ₹18L Monthly Leak

## Business Problem
HR Master has 500 employees, but Salary sheet has mismatch - finding fraud & errors.

## My Findings
- **212 Employees:** Salary missing → `ID not in Salary Sheet`
- **30 Cases:** Double salary drawn → Duplicate EMP0388 type
- **18 Ghost IDs:** Salary paid to fake IDs (EMP0515 etc) not in Master
- **Business Impact:** ~₹18 Lakhs/month potential leak

## Excel Skills Demonstrated
- `XLOOKUP` for missing & ghost check
- `COUNTIF` for duplicate detection
- Conditional Formatting & Filtering

## Key Formulas
`=XLOOKUP(A2,Salary_Haphazard!$A$2:$A$501,Salary_Haphazard!$B$2:$B$501,"ID not in Salary Sheet")`
`=IF(COUNTIF($A$2:A2,A2)>1,"Duplicate","Unique")`

**Tools:** Advanced Excel | **Author:** Sarbajit Roy - Kolkata | Aspiring Data Analyst
