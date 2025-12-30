# course-outline-scraper
Scrapes important course information from UW course outlines and exports it to Excel

Information scraped:
- Course code/course name
- Assessments and weight
- Course Description
- Learning Outcomes

# Setup 
1. Download all course outlines as html file

<img width="600" height="740" alt="Screenshot 2025-12-29 180650" src="https://github.com/user-attachments/assets/c00264f2-5ba7-4cbd-ab3d-e8438d9e80be" />

2. Install dependencies:
```bash
pip install beautifulsoup4 pandas openpyxl lxml
```
3. Replace sample file paths at the beginning of scraper.py 

4. Run the script:
```bash
python scraper.py
```
5. Check CWD for file named All_Courses_Outline.xlsx
