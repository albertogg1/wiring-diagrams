# Generador de Diagramas de Interconexión

Script Python para generar diagramas de interconexión en formato Draw.io desde archivos Excel.

![Ejemplo de diagrama generado](pinout.svg)


## 📦 Instalación

```bash
pip install -r requirements.txt
```

## 🚀 Uso

```bash
python interconnection_drawio.py tu_archivo.xlsx
```

## 📊 Formato del Excel

El archivo Excel debe tener estas columnas:

| Modulo1 | Conector1 | Pin1 | Señal1 | Señal2 | Pin2 | Conector2 | Modulo2 |
|---------|-----------|------|--------|--------|------|-----------|---------|
| MCU     | J1        | 1    | VCC    | 5V     | 1    | PWR_IN    | Power   |
| MCU     | J1        | 2    | GND    | GND    | 2    | PWR_IN    | Power   |

## ✨ Características

- Líneas ortogonales sin solapamientos
- Agrupamiento inteligente de módulos
- Pines ordenados automáticamente
- Sin duplicados (múltiples cables del mismo pin)
- Filtrado automático de pines sin señal
- Formato Draw.io nativo totalmente editable

## 📝 Salida

Genera un archivo `.drawio` que puedes abrir y editar en:
- https://app.diagrams.net
- Draw.io Desktop

---

¡Listo para generar diagramas profesionales! 🎉
