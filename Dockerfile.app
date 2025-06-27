# Dockerfile.app
# Use an official Python runtime as a parent image
FROM python:3.10-slim-buster

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container at /app
COPY requirements_app.txt .

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements_app.txt

# Copy the rest of the application code into the container
COPY app.py .

# Expose the port Streamlit runs on
EXPOSE 8501

# Command to run the Streamlit application
# Use environment variables for sensitive data and service URLs
# CONVERSION_SERVICE_URL will be set during deployment (e.g., in Kubernetes Deployment, Cloud Run)
# OPENAI_API_KEY will also be set during deployment
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
