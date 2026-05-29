# App web de propuestas (Streamlit) — prototipo

App web sencilla para que tu hermano genere la propuesta desde el navegador: llena un formulario, pulsa "Generar" y descarga el Excel de cálculo y el documento Word. Reutiliza el mismo motor (`motor_propuestas.py`), así que los números coinciden con la plantilla maestra.

## Archivos necesarios (deben ir juntos)

- `app.py` — la interfaz web.
- `motor_propuestas.py` — el motor de cálculo (no editar salvo para reglas).
- `requirements.txt` — dependencias.
- `config_proyecto.json` — opcional, configuración de ejemplo para cargar.

## Probar en tu computador (local)

```
pip install -r requirements.txt
streamlit run app.py
```

Se abre en el navegador (http://localhost:8501).

## Cómo se usa

1. Pestaña "1. Datos": código, cliente, objeto, n° títulos, utilidad, IVA.
2. "2. Actividades": activa cada actividad, define días de TPS (campo/entregables) y, si quieres, ajusta personal y logística.
3. "3. Entregables": productos y tiempo (meses).
4. "4. Desembolsos": porcentajes (deben sumar 100%).
5. "5. Catálogo": precios base (se ajusta pocas veces).
6. "▶ Generar": muestra los totales y botones para descargar Excel y Word.

En la barra lateral puedes guardar/cargar la configuración del proyecto como `.json` y restablecer el ejemplo.

## Publicarla gratis (Streamlit Community Cloud) — recomendado

1. Crea un repositorio en GitHub con estos archivos: `app.py`, `motor_propuestas.py`, `requirements.txt` (y opcional `config_proyecto.json`).
2. Entra a https://share.streamlit.io , conecta tu cuenta de GitHub y elige el repositorio.
3. Archivo principal: `app.py`. Pulsa "Deploy".
4. Obtendrás un link (p. ej. `https://tu-app.streamlit.app`) que le pasas a tu hermano. No instala nada.

Notas del plan gratis: la app "duerme" tras inactividad y tarda ~20s en despertar; no guarda histórico.

## Alternativas de hosting

- Render.com (free/economico): comando de inicio `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`.
- Si más adelante quieren catálogo de precios versionado e histórico de propuestas, se sube a base de datos (Supabase) sin rehacer el motor.

## Notas

- El prototipo no guarda datos entre sesiones; usa el `.json` para conservar configuraciones.
- Mantienes las tres vías: plantilla Excel, macro VBA y esta app web. Todas usan la misma lógica de costeo.
