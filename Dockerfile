FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copiamos codigo solamente. La carpeta data NO se empaqueta en GitHub/Railway.
# Los catalogos grandes deben vivir en Railway Volume usando DATA_DIR=/data.
COPY . /app

ENV DATA_DIR=/data
ENV QUANTIA_ENABLE_AI=0
ENV ENABLE_AI_MATERIAL_MATCHING=0
EXPOSE 8000

CMD ["python", "-m", "uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
