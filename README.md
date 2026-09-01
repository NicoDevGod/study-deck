# Study Deck — Dos Proyectos de IA

Presentación de estudio (`dos-proyectos-de-ia.pptx`) que documenta en profundidad
dos proyectos de IA construidos para practicar:

- [RAG Chatbot](https://github.com/NicoDevGod/rag-chatbot-cv) — chatbot con
  retrieval-augmented generation que responde preguntas sobre mi experiencia.
- [Waste Classifier](https://github.com/NicoDevGod/waste-classifier) — clasificador
  de residuos reciclables con transfer learning (MobileNetV2 + ONNX).

21 diapositivas: arquitectura de cada proyecto, conceptos clave explicados desde
cero, los bugs reales que aparecieron en el camino al deploy y cómo se resolvieron,
y las lecciones que se repiten entre ambos.

## Regenerar la presentación

El deck se genera con [`build.js`](build.js) usando `pptxgenjs`:

```bash
npm install
node build.js
```

Produce `dos-proyectos-de-ia.pptx` en el mismo directorio.
