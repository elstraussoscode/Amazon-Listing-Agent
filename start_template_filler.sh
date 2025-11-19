#!/bin/bash
# Startup script for Amazon Listing Agent - Template Filler

cd "$(dirname "$0")"

# Activate virtual environment
source venv/bin/activate

# Run the enhanced Streamlit application
streamlit run app_enhanced.py

# Deactivate when done
deactivate
