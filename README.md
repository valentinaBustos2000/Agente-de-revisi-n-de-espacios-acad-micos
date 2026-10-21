## Agente de consulta de disponibilidad de salones

Este proyecto implementa un agente conversacional para consultar la disponibilidad de salones de la Universidad de Ibagué. El usuario puede realizar solicitudes en lenguaje natural indicando fechas, horarios, capacidad, características requeridas o el código de un salón específico.

El agente permite:

- Buscar salones por capacidad y características como sillas, mesas, pantalla, video beam, laboratorio, sala de cómputo, modalidad híbrida y aire acondicionado.
- Consultar la disponibilidad de un salón específico.
- Procesar hasta cinco fechas en una misma solicitud.
- Interpretar rangos de fechas y recurrencias, por ejemplo, “todos los lunes de noviembre”.
- Validar horarios y solicitar únicamente la información faltante.
- Presentar alternativas cuando un salón específico no está disponible.
- Responder con información obtenida directamente del sistema institucional, sin realizar reservas.

La solución integra **n8n** para la orquestación del flujo conversacional, **Ollama con Llama 3.2** como modelo de lenguaje y un
