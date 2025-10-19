#!/bin/bash
echo "Starting Invoice PDF → Excel app..."
uvicorn web.main:app --host 0.0.0.0 --port 8000
