# tradutor-api
Tradutor - backend (Python/FastAPI + PyMuPDF). Extraao, traducao e reinsercao de texto em PDF/DOCX.

## Testes

```
pip install -r requirements-dev.txt
pytest
```

Nunca chama a API real da OpenAI (`_call_translate` vem sempre stubado) — roda de graça, sem
internet e sem depender de chave válida. `pytest -m "not slow"` pula os testes que acionam
OCR/inpaint real (ainda sem custo de IA, só mais lentos).
