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
# libgl1/libglib2.0-0: o opencv-python (dependencia transitiva do
# paddleocr/paddlex) precisa do libGL mesmo rodando sem interface grafica
# nenhuma — sem isso quebra com "ImportError: libGL.so.1: cannot open
# shared object file" ao importar o paddleocr.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer fonts-crosextra-carlito fonts-crosextra-caladea \
    build-essential zlib1g-dev libjpeg-dev libpng-dev libfreetype6-dev \
    liblcms2-dev libopenjp2-7-dev libtiff-dev libwebp-dev \
    libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*
COPY fontconfig/29-aptos-substitute.conf /etc/fonts/conf.d/29-aptos-substitute.conf
RUN fc-cache -f
# Instala o torch CPU-only ANTES do requirements.txt -- iopaint nao fixa
# versao de torch, entao o pip resolve pra ultima disponivel, que por padrao
# vem com CUDA (pacotes nvidia-*/triton, uteis so com GPU). Essa VPS e
# ARM64 sem GPU nenhuma (Oracle A1.Flex) -- a variante CUDA nunca roda o
# codigo dela, so ocupa ~5GB de imagem a toa. Fixado na mesma versao que o
# iopaint ja resolvia (2.14.0) pra nao conflitar com o requirement dele.
RUN pip install --no-cache-dir torch==2.14.0 --index-url https://download.pytorch.org/whl/cpu
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN chown -R tradutor:tradutor /app
USER tradutor
# O usuario tradutor nao tem diretorio home de verdade (useradd -r, sem
# -m) — sem isso o PaddleX quebra tentando criar cache em /home/tradutor
# (PermissionError). Mesma familia de problema que ja tinha aparecido com
# o LibreOffice (UserInstallation explicito no pipeline.py).
ENV HOME=/tmp
EXPOSE 8123
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8123"]
