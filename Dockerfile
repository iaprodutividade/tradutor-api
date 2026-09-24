FROM python:3.12-slim
WORKDIR /app
RUN groupadd -r tradutor && useradd -r -g tradutor tradutor
# libreoffice-writer: converte .docx -> .pdf sem interface grafica, usado
# pra gerar a imagem de previa do DOCX (mesmo layout visual do Word).
# fonts-crosextra-carlito/caladea: substitutos metricamente compativeis das
# fontes proprietarias da Microsoft (Aptos/Cambria Math) que o LibreOffice
# nao tem — sem isso a previa sai com fonte serifada generica, bem diferente
# do documento real.
# build-essential/zlib1g-dev/libjpeg-dev/...: o iopaint trava Pillow==9.5.0
# (lancado antes do Python 3.12 existir) — nao tem pacote pronto pra essa
# combinacao em nenhuma arquitetura, entao o pip compila do zero e precisa
# desses cabecalhos, senao quebra com "RequiredDependencyException: zlib".
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer fonts-crosextra-carlito fonts-crosextra-caladea \
    build-essential zlib1g-dev libjpeg-dev libpng-dev libfreetype6-dev \
    liblcms2-dev libopenjp2-7-dev libtiff-dev libwebp-dev \
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
