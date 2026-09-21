FROM python:3.12-slim
WORKDIR /app
RUN groupadd -r tradutor && useradd -r -g tradutor tradutor
# libreoffice-writer: converte .docx -> .pdf sem interface grafica, usado
# pra gerar a imagem de previa do DOCX (mesmo layout visual do Word).
# fonts-crosextra-carlito/caladea: substitutos metricamente compativeis das
# fontes proprietarias da Microsoft (Aptos/Cambria Math) que o LibreOffice
# nao tem — sem isso a previa sai com fonte serifada generica, bem diferente
# do documento real.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer fonts-crosextra-carlito fonts-crosextra-caladea \
    && rm -rf /var/lib/apt/lists/*
COPY fontconfig/29-aptos-substitute.conf /etc/fonts/conf.d/29-aptos-substitute.conf
RUN fc-cache -f
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN chown -R tradutor:tradutor /app
USER tradutor
EXPOSE 8123
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8123"]
