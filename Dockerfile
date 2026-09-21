FROM python:3.12-slim
WORKDIR /app
RUN groupadd -r tradutor && useradd -r -g tradutor tradutor
# libreoffice-writer: converte .docx -> .pdf sem interface grafica, usado
# pra gerar a imagem de previa do DOCX (mesmo layout visual do Word).
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN chown -R tradutor:tradutor /app
USER tradutor
EXPOSE 8123
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8123"]
