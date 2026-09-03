# Study Deck — Tres Proyectos de IA

Presentación de estudio (`tres-proyectos-de-ia.pptx`) que documenta en profundidad
los proyectos de IA construidos para practicar:

- [RAG Chatbot](https://github.com/NicoDevGod/rag-chatbot-cv) — chatbot con
  retrieval-augmented generation que responde preguntas sobre mi experiencia.
- [Waste Classifier](https://github.com/NicoDevGod/waste-classifier) — clasificador
  de residuos reciclables con transfer learning (MobileNetV2 + ONNX).
- [Tool-Use Agent](https://github.com/NicoDevGod/tool-agent) — agente con function
  calling (calculadora, clima, Wikipedia, hora) vía la API nativa de Groq.

29 diapositivas: arquitectura de cada proyecto, conceptos clave explicados desde
cero, los bugs reales que aparecieron en el camino al deploy y cómo se resolvieron,
y las lecciones que se repiten entre los tres. Este deck crece con cada proyecto
nuevo del portafolio de IA.

## Regenerar la presentación

El deck se genera con [`build.js`](build.js) usando `pptxgenjs`:

```bash
npm install
node build.js
```

Produce `tres-proyectos-de-ia.pptx` en el mismo directorio.
