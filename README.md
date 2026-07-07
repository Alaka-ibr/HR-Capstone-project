# HR Analytics Capstone project
## Why I'm writing this

Before I even opened the workbook properly, I wanted to write down what I actually did and why, step by step, so that if someone else opened this file they wouldn't have to guess my thought process. This isn't a polished report it's basically a log of how the cleaning went, including the parts that didn't work on the first try.

## First look at the data

I started by just scrolling through `Raw_Employee_Data` before touching anything. 157 rows, 8 columns: Emp ID, First_Name, Last_Name, Dept Code, Hire Date, Salary, Perf_Score, Employment Status. There were also two lookup sheets one mapping department codes to full names, and one mapping performance score ranges to a band and a bonus percentage.

A few things jumped out almost immediately:

- Some Emp IDs showed up more than once, and when I checked, the rows were exact copies same everything. Looked like duplicate entries, not different people who happened to share an ID.
- Dept Code was written in different ways for what was clearly the same department  `IT02`, `it02`, ` IT02 ` with random spaces. Annoying, but fixable.
- Employment Status was worse  `active`, `ACTIVE`, `Actv`, `Left`, `left`, `Resigned`, all mixed together.
- Salary looked fine at first glance until I noticed some cells had `$` and commas in them (like `$114,425`), which meant Excel was reading those as text, not numbers.
- Hire Date was a mess of formats  some looked like normal dates, some were like `02-Mar-20`, others `09/24/2019`. All in the same column.
- And then there were just plain empty cells scattered in First_Name, Salary, and Perf_Score.

So basically: duplicates, inconsistent text, wrong data types, and missing values. Pretty standard "messy real-world data" problems once I actually looked closely.

## How I decided to fix each thing (and in what order)

I didn't clean things randomly I thought about what depended on what, and tried to fix things in an order that wouldn't break a later step.

**1. Duplicates first.**
Before deleting anything I double checked the duplicate rows really were identical across every column, not just the same ID by coincidence. Once I confirmed that, I used Remove Duplicates on Emp ID. Went from 157 rows down to 150. Made sense to do this before anything else — no point cleaning data I was about to delete anyway.

**2. Cleaned up Dept Code.**
Used TRIM + UPPER to knock out the extra spaces and inconsistent casing. I did this early on purpose, because I knew I'd later need to match Dept Code exactly against the lookup table, and any leftover mess here would quietly break that.

**3. Employment Status was trickier than I expected.**
My first instinct was to just TRIM/UPPER this too, like I did with Dept Code. But that doesn't fix "Actv" — that's not a spacing or capitalization problem, it's just a different spelling. So I had to write an actual IF-based formula that specifically catches "Actv" and "Active" (in any case) and turns them both into "Active." Same for Left. Resigned I left as its own thing I went back and forth on whether to lump it in with "Left" for attrition purposes, but the assignment specifically asks for the percentage of employees whose status is 'Left', so I kept Resigned separate rather than merging it in.

**4. Salary numbers.**
Stripped out the `$` and commas with SUBSTITUTE, then wrapped it in VALUE to force it back into an actual number. I made sure to do this before filling in the missing salaries, because the formula I used to fill blanks needed to average real numbers if some salaries were still text at that point, they'd just get ignored by AVERAGEIF and throw off the numbers.

**5. Hire Date.**
This one took two attempts. Text to Columns fixed most of the text-based dates. But a handful still came through weird turned out they were actual datetime values with a hidden time attached , so they weren't displaying cleanly. I used INT() to strip off the time part and just keep the date, then reformatted the column properly.

**6. Missing values.**
- First_Name: 5 blanks, no way to recover the actual names from anything else in the sheet, so I just filled them with "Unknown" rather than leaving blanks that would've broken the Full Name formula later.
- Salary: 21 blanks. Instead of guessing a random number or using one overall average for the whole company, I used each employee's department average salary clearly differs across departments, so this felt like the more honest estimate.
- Perf_Score: 9 blanks, same idea, department average.

One thing I almost got wrong here: I initially thought about filling in the missing salaries/scores before cleaning up Dept Code, but caught myself if Dept Code still had "IT02" and "it02" as separate values at that point, the department averages would've been calculated on split, incomplete groups instead of the whole department. So I made sure Dept Code was fully cleaned first.

**7. Double-checked everything matched.**
Ran a quick COUNTIF check to confirm every cleaned Dept Code actually existed in the lookup table, and confirmed Employment Status only had three values left. Came back clean.

## The columns I built on top of the cleaned data

Full Name, Department Name, Years of Service, Performance Band, and Eligible Bonus Amount  all live formulas, nothing pasted as static values. Department Name and Performance Band both use VLOOKUP, but slightly differently: Department Name needs an exact match (FALSE), while Performance Band and Bonus Amount use an approximate match (TRUE) because the performance table is a range — 0–49, 50–69, and so on  sorted from smallest to largest. That sorting is actually required for the approximate match to work correctly; if the table weren't sorted, VLOOKUP would return wrong results.

## Things that tripped me up

Honestly, a couple of things didn't work the first time:

- The date formatting issue above took me a minute to realize some "already fixed" dates weren't actually fixed, just hiding a time value.
- Employment Status cleaning  my first pass with plain TRIM/UPPER left "Actv" untouched, and I only caught it because I ran a quick unique-values check afterward and saw it sitting there separately from "Active."
- Later on, once I got to the dashboard, I noticed my slicers weren't affecting my KPI cards at all turns out slicers only talk to PivotTables/PivotCharts, and my KPI cards were just plain formulas pointing straight at the raw data, so they had no connection to the slicers whatsoever. I fixed it by building a small separate PivotTable just to power the KPI numbers, connecting that PivotTable to the same slicers, and then pointing my KPI cards at that PivotTable instead of the raw data directly.

## Analysis: Observation vs. Insight, by Question

I wanted to be clear about the difference between what the data literally shows (observation) versus what I think it means (insight)  a number by itself isn't an insight until you interpret it.

**Q1 — Headcount & Average Salary by Department**
- *Observation:* IT is the largest department (36 people), while Finance is the smallest (17). Average salary ranges from about $79,800 (Operations) to $87,900 (Finance) — roughly an $8,000 spread across departments.
- *Insight:* Headcount and pay don't move together — the biggest department (IT) doesn't pay the most, and the smallest (Finance) does. So department size alone isn't a reliable stand-in for how well that department compensates its people; each one needs to be looked at on its own terms.

**Q2 — Attrition Rate by Department**
- *Observation:* Operations has a "Left" rate of 20.7%, clearly higher than every other department (11–17% range). Company-wide, 16.0% of employees have status "Left."
- *Insight:* Operations isn't just slightly worse — it's losing people at close to double the rate of the best-performing department (HR, 11.1%). And since Operations' average salary ($79,800) is similar to IT's ($80,000), which has a much lower attrition rate (16.7%), pay alone doesn't explain the gap. That points toward something specific to how Operations is run — workload, scheduling, management — as a more likely driver, though the dataset itself can't confirm what exactly.

**Q3 — Average Performance Score by Employment Status**
- *Observation:* Resigned employees average a 76.4 performance score — higher than both Active (67.9) and Left (66.1) employees.
- *Insight:* This is the one that surprised me most. It suggests the company isn't losing its weakest performers  voluntary resignations skew toward people who were actually doing well. That's a more costly kind of turnover than losing underperformers, because you're not just backfilling a role, you're replacing someone who was good at it.

**Q4 — Bonus Payout by Performance Band**
- *Observation:* The "Achieving" band accounts for roughly $190K of the $502K total bonus pool  the single largest share, ahead of exceding** (~$145K) and "Outstanding" (~$92K).**
- *Insight:* This isn't because top performers get shortchanged it's simply that more employees land in the middle "Achieving" band than in the very top bands, so the group as a whole draws a bigger total slice even at a lower bonus percentage each. Worth knowing if leadership assumes the bonus budget is mostly going to star performers — most of it, in dollar terms, is actually going to solidly-average employees.

## What I noticed once everything was clean

Once the numbers were actually right, a few things stood out from my own Pivot Tables:
- **Operations has the worst attrition by a clear margin** — 20.7% of Operations employees have status **Left **versus 16.0% company-wide. Every other department sits between 11–17%. That's not a small gap,Operations is losing people at roughly double the rate of HR (11.1%), so if I were handing this to an actual HR team, that's the department I'd flag first.
- **Pay doesn't explain it.** Operations has the **lowest** average salary of the six departments approximately ($79,800) but so does IT approximately($80,000), which has a much lower attrition rate (16.7%) despite similar pay. So low salary alone doesn't cleanly predict who's leaving; something specific to Operations (workload, management, shift patterns the data doesn't say) is probably driving it, not compensation alone.
- **People who resign look different from people who get marked "Left."** Employees with status "Resigned" have the *highest* average performance score of any group (76.4), actually higher than Active employees (67.9) and well above Left employees (66.1). That's a pattern worth flagging on its own: it suggests the strongest performers are the ones voluntarily walking out the door, which is a more expensive kind of attrition than losing lower-performing employees.
- **Bonus payout is heavily concentrated in the upper-middle bands, not the very top.** The "Achieving" band accounts for the single largest share of total bonus payout (approximately $190K of the approximately $502K total), more than "Exceeding" and "Outstanding" combined get individually. That's simply because there are more people clustered in the middle of the performance curve than at the very top worth knowing if leadership assumes most of the bonus budget goes to the highest performers.
- **Missing data wasn't trivial.** 21 of 150 salary records and 9 performance scores were blank before cleaning that's roughly 1 in 7 salary rows. I used department-level averages to fill these rather than one company-wide number, since salary clearly varies by department; it's still an estimate, not the real figure, so anyone using this data for actual pay decisions should know which rows were imputed.
  
**Finally**
  
  This is the Pictorial visualization of the dashbord which give more details about the whole analysis carried out

  
