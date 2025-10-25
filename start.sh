#!/bin/bash

# Install LibreOffice if not already installed
if ! command -v libreoffice &> /dev/null; then
    echo "Installing LibreOffice..."
    apt-get update
    apt-get install -y libreoffice
fi

# Start the Flask application
python app.py