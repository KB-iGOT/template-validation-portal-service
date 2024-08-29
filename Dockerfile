# Use a lightweight Python image as a base
FROM python:3.10-slim

# Set environment variables
ENV FLASK_APP=app.py
ENV FLASK_RUN_HOST=0.0.0.0
ENV FLASK_RUN_PORT=5000

# Create and set the working directory
WORKDIR /app

# Install virtualenv
RUN pip install --no-cache-dir virtualenv

# Create a virtual environment named 'tvp'
RUN python3 -m venv tvp

# Copy the requirements file into the container
COPY apiServices/src/main/requirements.txt .

# Activate the virtual environment and install the requirements
RUN /bin/bash -c "source tvp/bin/activate && pip install --no-cache-dir -r requirements.txt"

# Copy the application code into the container
COPY apiServices/src/main/ ./

# Copy the entire backend folder into the container
COPY backend /app/backend

# Expose the port the application will run on
EXPOSE 5000

# Command to run the Flask application
CMD ["/bin/bash", "-c", "source tvp/bin/activate && flask run --host=0.0.0.0 --port=5000"]
