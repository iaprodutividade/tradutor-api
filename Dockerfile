FROM python:3.12-slim
WORKDIR /app
RUN groupadd -r tradutor && useradd -r -g tradutor tradutor
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN chown -R tradutor:tradutor /app
USER tradutor
EXPOSE 8123
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8123"]
