# Bodyshop Translator MVP

A FastAPI-powered web application that processes and translates body shop repair documents (PDFs). Upload an estimate or repair order PDF and get back a clean, structured output — making it easier to read, compare, and manage auto body repair data.

## Features

- Upload body shop PDF documents via a web interface
- Extract and parse tabular repair data using `pdfplumber`
- Process and restructure data with `pandas`
- Generate clean output PDFs with `reportlab`
- Lightweight REST API built with FastAPI

## Tech Stack

| Layer | Technology |
|---|---|
| API Framework | [FastAPI](https://fastapi.tiangolo.com/) |
| PDF Parsing | [pdfplumber](https://github.com/jsvine/pdfplumber) |
| Data Processing | [pandas](https://pandas.pydata.org/) |
| PDF Generation | [reportlab](https://www.reportlab.com/) |
| Server | [uvicorn](https://www.uvicorn.org/) |
| Containerization | Docker |
| Deployment | [Render](https://render.com/) |

## Project Structure

```
bodyshop-translator-mvp/
├── app/                  # FastAPI application code
├── Dockerfile            # Container definition
├── render.yaml           # Render deployment config
└── requirements.txt      # Python dependencies
```

## Getting Started

### Prerequisites

- Python 3.9+
- pip

### Local Development

1. Clone the repository:
   ```bash
   git clone https://github.com/rkhan15/bodyshop-translator-mvp.git
   cd bodyshop-translator-mvp
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the development server:
   ```bash
   uvicorn app.main:app --reload
   ```

4. Open your browser at `http://localhost:8000`

### Running with Docker

```bash
docker build -t bodyshop-translator .
docker run -p 8000:8000 bodyshop-translator
```

## Deployment

This project is configured for one-click deployment to [Render](https://render.com/) via `render.yaml`. Push to `main` and Render will automatically build the Docker image and deploy the service.

## API

Interactive API docs are available at `/docs` (Swagger UI) or `/redoc` once the server is running.

## License

This project is currently unlicensed. All rights reserved by the repository owner.
