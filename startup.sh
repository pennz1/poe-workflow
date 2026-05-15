#!/bin/bash
# Azure App Service startup script for Streamlit
streamlit run app.py \
    --server.port "${PORT:-8000}" \
    --server.address 0.0.0.0 \
    --server.headless true \
    --browser.gatherUsageStats false
