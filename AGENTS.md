# Project Overview

This project generates monthly GA4 reports.

Flow:
GA4 → Excel → PPT

Rules:
- Do not modify excel_gen.py
- Do not modify ppt_gen.py
- Only implement orchestration in main.py

GA4 Data:
- 3 languages: ko / en / cn
- summary, channels, top pages

Goal:
Generate Excel + PPT report from CLI