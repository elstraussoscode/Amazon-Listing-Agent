#!/bin/bash
# Startup script for Amazon Listing Agent Streamlit App

cd "$(dirname "$0")"

# Activate virtual environment
source venv/bin/activate

# Run the Streamlit application
streamlit run app.py

# Deactivate when done
deactivate

