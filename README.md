# SanSan Data Matcher 🔍

A Streamlit web application for matching and verifying SanSan operation records with daily capacity reports.

## Features
- **File Upload & Parsing**: Supports Excel (`.xlsx`) daily capacity reports and SanSan operation records (`.csv`).
- **Interactive Matching**: Automated verification with detailed matching metrics.
- **Export Reports**: Generate and download comprehensive comparison reports.
- **User Manual**: Integrated documentation and guide for end users.

## Setup & Running Locally

### 1. Install Dependencies
Make sure you have Python 3.9+ installed. Then install required libraries:
```bash
pip install -r requirements.txt
```
*(Or double-click `install_requirements.bat` on Windows)*

### 2. Run the Web App
```bash
streamlit run app.py
```
*(Or double-click `Launch_WebApp.bat` on Windows)*

## Deployment to Streamlit Cloud
1. Push this repository to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io/) and log in with GitHub.
3. Select this repository, set branch to `main`, and main file path to `app.py`.
4. Click **Deploy**.
