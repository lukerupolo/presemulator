        # Dockerfile.app
        # Use an official Python runtime as a parent image
        FROM python:3.10-slim-buster

        # Install git for cloning repositories
        RUN apt-get update && \
            apt-get install -y --no-install-recommends git && \
            rm -rf /var/lib/apt/lists/* && \
            apt-get clean

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
        CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
        
