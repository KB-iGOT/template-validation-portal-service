# Use a lightweight Python image as a base
FROM python:3.10-slim

# Set environment variables for Flask
ENV FLASK_APP=app.py
ENV FLASK_RUN_HOST=0.0.0.0
ENV FLASK_RUN_PORT=5000

# Set the working directory inside the container
WORKDIR /app

# Install virtualenv
RUN pip install --no-cache-dir virtualenv

# Create a virtual environment
RUN python3 -m venv tvp

# Copy the application code into the container
COPY apiServices/src/main/ .
COPY backend/ ./backend  

# Print the directory structure to verify correct placement
RUN ls -R

# Activate the virtual environment and install the requirements
RUN /bin/bash -c "source tvp/bin/activate && pip install --no-cache-dir -r requirements.txt"

# Expose the application port
EXPOSE 5000

# Command to run the Flask application
CMD ["/bin/bash", "-c", "source tvp/bin/activate && flask run --host=0.0.0.0 --port=5000"]












