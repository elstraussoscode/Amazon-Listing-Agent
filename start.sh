#!/bin/bash
# Startup script for Amazon Listing Agent

cd "$(dirname "$0")"

# Activate virtual environment
source venv/bin/activate

# Run the application
python amazon_listing_agent.py

# Deactivate when done
deactivate
