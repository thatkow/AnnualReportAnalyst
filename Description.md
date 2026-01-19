# Annual Report Analyst — Data Analyst Workflow Summary

## What the app does
Annual Report Analyst is a desktop (Tkinter) application that helps analysts ingest a folder of annual report PDFs, identify relevant pages with configurable regex patterns, review/curate the matches, and then run an AI-assisted extraction pipeline that saves normalized tables and a combined CSV output for downstream analysis.

## Engineered pipeline (analyst perspective)
1. **Define extraction patterns and load PDFs**
   - Analysts configure regex patterns per financial category (columns such as financial statements, income, shares, etc.) and optional year patterns in the Review tab. The app compiles these patterns and scans every PDF page to find matches and identify the report year. The results are stored per PDF entry, and the Review grid is rebuilt to show matching pages and thumbnails for each category. 
   - This creates a structured inventory of candidate pages across all PDFs, which is the first pass of document discovery and classification.

2. **Review and curate matched pages**
   - In the Review tab, analysts inspect the PDF thumbnails, select pages for each category, and commit these selections. 
   - Committing writes the selections to `assigned.json` in the company folder, providing a reproducible record of which pages were selected for each PDF and category.

3. **Prepare AIScrape jobs and run AI extraction**
   - The Scrape workflow uses the selected pages to prepare jobs for each PDF/category combination. Depending on configuration, the app can export just the selected pages into temporary PDFs, upload them to OpenAI, and request structured output based on prompt templates. 
   - Responses are saved under an `openapiscrape` directory (per company or per PDF folder), and associated PDF slices are stored for traceability.

4. **Review scraped tables and preview source pages**
   - The Scrape tab presents the AI-extracted tables in a Treeview, lets analysts inspect the rows, and previews the source PDF pages for quick validation. This step closes the loop by tying each extracted row back to the original document evidence.

5. **Generate combined dataset for analysis**
   - The Combined tab reads the saved scrape outputs, builds a unified dataset, and appends dynamic columns (date columns derived from PDFs). 
   - The combined output is rendered in the UI and automatically saved as `Combined.csv`, enabling downstream analysis and reporting workflows.

6. **Post-pipeline analysis (analyst package)**
   - After `Combined.csv` is generated, the `analyst` package provides a data analysis layer that loads one or more companies into structured `Company` objects and supports downstream visuals and comparisons. 
   - Single-company analysis uses stacked financial visuals to normalize values (multipliers, per-share conversion, release-date shifts) and render interactive HTML reports. 
   - Multi-company analysis merges normalized datasets and produces comparative stacked visuals across tickers, ensuring required meta rows are present so like-for-like comparisons are consistent. 
   - Additional analytical helpers compute adjusted values, release-date-aware statistics, and diagnostic plots such as boxplots/violins for financial distributions.

## Key artifacts produced
- `assigned.json`: curated page selections per PDF/category (review checkpoint).
- `openapiscrape/`: AI extraction outputs and supporting PDF slices.
- `Combined.csv`: consolidated dataset ready for analysis.
- `visuals/*.html`: optional analyst visual outputs generated from the combined data.
