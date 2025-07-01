# TagExtractor
This Python tool recursively scans Excel files within a specified directory, extracts predefined audit tags (e.g., #QUERY, #ISSUE, #RECOMMENDATION, etc.), and compiles the results into a centralized Excel tracker — complete with Excel-formatted cell references, metadata, and deduplication logic.


🚀 Key Features
🔍 Deep Folder Scanning: Recursively searches all .xlsx and .xlsm files within the target audit folder.

🏷️ Smart Tag Detection: Identifies tagged audit elements like #QUERY, #ISSUE, #FINDING, and more.

📌 Precise Location Mapping: Captures file name, worksheet name, and exact cell address for each tag.

🧠 Hash-based Deduplication: Prevents duplicate tag entries using SHA-1 fingerprints.

🧾 Excel Output with Native Links: Outputs to a neatly structured Excel file with clickable Excel-formula cell references.

⚙️ Resilient and Silent Fail-Safes: Gracefully skips inaccessible files and handles Excel corruption issues with appropriate warnings.

🧠 Tag Types Tracked
The script identifies the following standardized tags (case-insensitive):

Tag	Category
#QUERY	Query
#ACTION	Action
#RECOMMENDATION	Recommendation
#ISSUE	Issue
#RISK	Risk
#TEST	Test Step
#FINDING	Finding

📂 How It Works
Prompt: User is prompted to input the full path of the main audit folder.

Scan: All valid Excel files are recursively scanned.

Extract: Tag-containing cells are parsed and stored with metadata.

Export: Results are exported into a time-stamped Excel file (Audit_Tag_Tracker_YYYYMMDD_HHMMSS.xlsx), including:

Deduplicated tag entries

Excel links (='[filename]SheetName'!$A$1) to directly jump to the source cell

Auto-formatted headers for ease of review

🧩 Integration Ideas
Embed into internal audit workflows for real-time tag aggregation

Use in conjunction with document review protocols

Automate follow-up tracking on #ACTION or #RECOMMENDATION tags

🛡️ Error Handling
Skips corrupt or locked Excel files (e.g., .tmp, .lock, BadZipFile)

Prints actionable error messages without halting execution

Ensures workbook closure after scanning to avoid resource locks

📈 Performance
Optimized for thousands of Excel files with minimal memory footprint

Real-time scan progress bar with memory and time diagnostics
