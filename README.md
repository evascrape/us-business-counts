# US business counts by industry, state and job title

How many businesses are there in a given US industry, how big are they, and in which states and counties are they? How many people hold a given buyer role, and where do they work? This repository answers those questions for **20 industries and 15 job titles that B2B teams commonly sell to**, using two official public sources, and ships the script that rebuilds the files.

It is meant for sizing a market before you build a prospect list: for example, to know that the US has about 135,665 dental offices with paid employees, that 95% of them have fewer than 20 employees, or that Texas has the second-largest number.

## Sources

| Source | Publisher | Vintage | What it counts |
|---|---|---|---|
| [County Business Patterns](https://www.census.gov/programs-surveys/cbp.html) | US Census Bureau | 2023 (published 2025) | Establishments (single physical locations) with paid employees, by NAICS industry, employee size class, state and county |
| [Occupational Employment and Wage Statistics](https://www.bls.gov/oes/tables.htm) | US Bureau of Labor Statistics | May 2024 | Employment and wages by occupation (SOC), industry and state |

Both are US federal government works in the public domain. The derived files here are released under [CC0 1.0](LICENSE).

## Files

| File | Rows | Contents |
|---|---|---|
| `data/industries.csv` | 20 | National establishments, employees, annual payroll and establishments by employee size class |
| `data/industry_by_state.csv` | 1,020 | Establishments, rank, employees and payroll for each industry in each state and DC |
| `data/industry_top_counties.csv` | 300 | Top 5 counties for the 3 states with the most establishments in each industry |
| `data/occupations.csv` | 15 | US employment and median annual wage per occupation |
| `data/occupation_top_states.csv` | 90 | Top 6 states by employment per occupation, with the state median wage |
| `data/occupation_top_industries.csv` | 90 | Top 6 industries (NAICS 3-digit or OEWS designation) employing each occupation |
| `scripts/build.py` | | Rebuilds everything from the original Census and BLS files (needs `openpyxl`) |

## Industries

| Census industry | NAICS (2017) | US establishments, 2023 |
|---|---|---|
| Restaurants and other eating places | 7225 | 618,476 |
| Manufacturing (all subsectors) | 31 | 284,452 |
| Construction of buildings | 236 | 254,630 |
| Offices of physicians | 6211 | 218,066 |
| Truck transportation | 484 | 170,286 |
| Offices of lawyers | 541110 | 165,491 |
| Offices of real estate agents and brokers | 531210 | 163,894 |
| Computer systems design and related services | 5415 | 145,813 |
| Accounting, tax preparation, bookkeeping and payroll services | 5412 | 137,069 |
| Offices of dentists | 621210 | 135,665 |
| Insurance agencies and brokerages | 524210 | 133,728 |
| Plumbing, heating and air-conditioning contractors | 238220 | 111,207 |
| Administrative and general management consulting services | 541611 | 104,494 |
| Real estate property managers | 53131 | 78,886 |
| Employment services | 5613 | 56,173 |
| Hotels (except casino hotels) and motels | 721110 | 55,895 |
| Fitness and recreational sports centers | 713940 | 41,556 |
| Advertising, public relations and related services | 5418 | 38,864 |
| Architectural services | 541310 | 22,805 |
| Software publishers | 511210 | 18,425 |

## Job titles

| BLS occupation | SOC | US employment, May 2024 | Median annual wage |
|---|---|---|---|
| General and Operations Managers | 11-1021 | 3,584,420 | $102,950 |
| Software Developers | 15-1252 | 1,654,440 | $133,080 |
| Accountants and Auditors | 13-2011 | 1,448,290 | $81,680 |
| Human Resources Specialists | 13-1071 | 917,460 | $72,910 |
| Financial Managers | 11-3031 | 818,620 | $161,700 |
| Computer and Information Systems Managers | 11-3021 | 645,970 | $171,200 |
| Sales Managers | 11-2022 | 603,710 | $138,060 |
| Medical and Health Services Managers | 11-9111 | 565,840 | $117,960 |
| Marketing Managers | 11-2021 | 384,980 | $161,030 |
| Construction Managers | 11-9021 | 348,330 | $106,980 |
| Property, Real Estate, and Community Association Managers | 11-9141 | 296,640 | $66,700 |
| Personal Financial Advisors | 13-2052 | 270,480 | $102,140 |
| Human Resources Managers | 11-3121 | 215,520 | $140,030 |
| Chief Executives | 11-1011 | 211,850 | $206,420 |
| Purchasing Managers | 11-3061 | 81,240 | $139,510 |

## Things to know before you use the numbers

- **Establishments are not companies.** County Business Patterns counts physical locations. A firm with five offices counts five times. For firm counts, see [Statistics of US Businesses](https://www.census.gov/programs-surveys/susb.html).
- **Businesses without employees are not included.** Sole proprietors and freelancers are in the Census [Nonemployer Statistics](https://www.census.gov/programs-surveys/nonemployer-statistics.html), not here.
- **Withheld cells are empty, not zero.** The Census Bureau does not publish some cells to protect confidentiality, and employee or payroll figures with a high noise flag are left empty.
- **BLS counts employees only.** Self-employed owners are not in the occupation figures, and some categories are wider than the job title (for example, recruiters are part of "Human Resources Specialists").
- **Staffing agencies look large** because temporary workers on the agency payroll count as its employees.

## Rebuilding

```
# Census: cbp23us.txt, cbp23st.txt, cbp23co.txt and georef22.txt in one folder
# BLS: oesm24nat, oesm24in4 and oesm24st zips extracted into the layout used in build.py
python scripts/build.py <cbp_dir> <oews_dir> out.json
```

The BLS site asks automated downloads to identify themselves; set a User-Agent with a contact address.

## Who made this

[Evascrape](https://evascrape.com/), a pay-per-record B2B lead list tool, uses these numbers to size lists before building them. The same Census and BLS figures are shown per industry at [evascrape.com/leads](https://evascrape.com/leads), with pages for the three largest states of each industry, and per job title at [evascrape.com/titles](https://evascrape.com/titles). The data here is independent of Evascrape's product and anyone can use it. See also [Awesome B2B Lead Sources](https://github.com/evascrape/awesome-b2b-lead-sources), a list of public business data sources.
