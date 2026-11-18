# Use the official Python 3.9 slim image
# This guarantees that the container runs Python 3.9 regardless of host OS
FROM python:3.9-slim

# Set environment variables:
# - PYTHONDONTWRITEBYTECODE: Prevents Python from writing .pyc files
# - PYTHONUNBUFFERED: Keeps stdout/stderr unbuffered so logs appear immediately
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

# Set working directory inside the container
WORKDIR /app

# Copy application script
COPY app.py .

# Expose port 8080 for web traffic
EXPOSE 8080

# Run the application
CMD ["python", "app.py"]
