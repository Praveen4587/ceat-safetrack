FROM python:3.11-slim

WORKDIR /app

# Copy all application files
COPY . /app

# Environment variables
ENV PORT=8091
ENV HOST=0.0.0.0
ENV PYTHONUNBUFFERED=1

EXPOSE 8091

CMD ["python", "app.py", "serve"]
