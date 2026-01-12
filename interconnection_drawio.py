#!/usr/bin/env python3
"""
Genera un diagrama Draw.io de interconexión entre módulos y conectores.

El Excel debe tener las siguientes columnas:
- Modulo1: Nombre del módulo de origen
- Conector1: Nombre del conector de origen
- Pin1: Número del pin de origen
- Señal1: Nombre de la señal en el origen
- Señal2: Nombre de la señal en el destino
- Pin2: Número del pin de destino
- Conector2: Nombre del conector de destino
- Modulo2: Nombre del módulo de destino
"""

import pandas as pd
import sys
from pathlib import Path
from collections import defaultdict
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import base64
import zlib


def generate_drawio_diagram(excel_file, output_drawio=None, sheet_name=0):
    """
    Genera un diagrama Draw.io de interconexión entre módulos y conectores.
    """
    # Leer el Excel
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Normalizar nombres de columnas
    df.columns = df.columns.str.strip()
    
    # Verificar que existan las columnas necesarias
    required_cols = ['Modulo1', 'Conector1', 'Pin1', 'Señal1', 
                     'Señal2', 'Pin2', 'Conector2', 'Modulo2']
    
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Faltan columnas requeridas: {', '.join(missing_cols)}")
    
    # Filtrar filas vacías o con valores NaN en campos críticos
    df = df.dropna(subset=['Pin1', 'Pin2'])
    
    # Convertir Pin1 y Pin2 a string y limpiar
    df['Pin1'] = df['Pin1'].astype(str).str.strip()
    df['Pin2'] = df['Pin2'].astype(str).str.strip()
    
    # Eliminar filas donde Pin1 o Pin2 sean cadenas vacías o 'nan'
    df = df[df['Pin1'] != '']
    df = df[df['Pin2'] != '']
    df = df[df['Pin1'].str.lower() != 'nan']
    df = df[df['Pin2'].str.lower() != 'nan']
    
    # Eliminar filas sin módulo o conector
    df = df.dropna(subset=['Modulo1', 'Conector1', 'Modulo2', 'Conector2'])
    
    # Resetear índice después de filtrar
    df = df.reset_index(drop=True)
    
    if len(df) == 0:
        raise ValueError("No hay filas válidas después de filtrar. Verifica que tu Excel tenga valores en Pin1 y Pin2.")
    
    # Organizar datos por módulo y conector
    left_connectors = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    right_connectors = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    connections = []
    
    # Función para verificar si una señal es válida (no vacía ni NC)
    def is_valid_signal(signal):
        if not signal:
            return False
        signal_lower = signal.lower().strip()
        # Excluir señales vacías y variantes de "No Connect"
        if signal_lower in ['', 'nc', 'n/c', 'n.c.', 'no connect', 'noconnect', 'n_c', 'na', 'n/a']:
            return False
        return True
    
    for idx, row in df.iterrows():
        mod1 = str(row['Modulo1']).strip()
        conn1 = str(row['Conector1']).strip()
        pin1 = row['Pin1']
        signal1 = str(row['Señal1']).strip() if pd.notna(row['Señal1']) else ''
        
        mod2 = str(row['Modulo2']).strip()
        conn2 = str(row['Conector2']).strip()
        pin2 = row['Pin2']
        signal2 = str(row['Señal2']).strip() if pd.notna(row['Señal2']) else ''
        
        # Validar señales
        signal1_valid = is_valid_signal(signal1)
        signal2_valid = is_valid_signal(signal2)
        
        # Ignorar conexiones donde ambas señales sean inválidas
        if not signal1_valid and not signal2_valid:
            continue
        
        # Solo agregar al conector izquierdo si tiene señal válida
        if signal1_valid:
            left_connectors[mod1][conn1][pin1].append({
                'signal': signal1,
                'row_idx': idx
            })
        
        # Solo agregar al conector derecho si tiene señal válida
        if signal2_valid:
            right_connectors[mod2][conn2][pin2].append({
                'signal': signal2,
                'row_idx': idx
            })
        
        # Guardar conexión solo si ambas señales son válidas
        if signal1_valid and signal2_valid:
            connections.append({
                'mod1': mod1,
                'conn1': conn1,
                'pin1': pin1,
                'signal1': signal1,
                'mod2': mod2,
                'conn2': conn2,
                'pin2': pin2,
                'signal2': signal2,
                'row_idx': idx
            })
    
    # Configuración del diagrama
    connector_width = 240
    connector_padding = 20
    pin_height = 28
    module_spacing = 50
    horizontal_gap = 450
    margin_x = 50
    margin_y = 100
    
    # Colores por tipo de señal (Draw.io usa formato HTML)
    signal_colors = {
        'power': '#e74c3c',
        'gnd': '#34495e',
        'ground': '#34495e',
        'vcc': '#e74c3c',
        'vdd': '#e74c3c',
        '5v': '#e74c3c',
        '3v3': '#e74c3c',
        'i2c': '#3498db',
        'spi': '#9b59b6',
        'uart': '#2ecc71',
        'gpio': '#f39c12',
        'adc': '#e67e22',
        'pwm': '#1abc9c',
        'can': '#16a085',
        'data': '#8e44ad',
        'clk': '#d35400',
        'clock': '#d35400',
        'reset': '#c0392b',
        'int': '#27ae60',
    }
    
    def get_signal_color(signal_name, row_idx):
        """Determina el color basado en el nombre de la señal"""
        signal_lower = signal_name.lower()
        for keyword, color in signal_colors.items():
            if keyword in signal_lower:
                return color
        # Colores por defecto
        default_colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', 
                         '#1abc9c', '#e67e22', '#95a5a6', '#16a085', '#d35400']
        return default_colors[row_idx % len(default_colors)]
    
    # Calcular el orden óptimo de los módulos para minimizar distancias
    # Crear un grafo de conexiones entre módulos izquierda-derecha
    connection_count = defaultdict(lambda: defaultdict(int))
    
    for conn in connections:
        left_mod = conn['mod1']
        right_mod = conn['mod2']
        connection_count[left_mod][right_mod] += 1
    
    # Función para calcular el "centro de masa" de las conexiones de un módulo
    def calculate_center_of_mass(module_key, is_left, positions_dict, other_positions_dict):
        if module_key not in positions_dict:
            return 0
        
        total_weight = 0
        weighted_sum = 0
        
        # Para cada conector en este módulo
        for conn_key in positions_dict.keys():
            if not conn_key.startswith(module_key + ':'):
                continue
            
            # Buscar conexiones relacionadas
            for conn in connections:
                if is_left:
                    if f"{conn['mod1']}:{conn['conn1']}" == conn_key:
                        target_key = f"{conn['mod2']}:{conn['conn2']}"
                        if target_key in other_positions_dict:
                            weight = 1
                            pos_y = other_positions_dict[target_key]['y']
                            weighted_sum += pos_y * weight
                            total_weight += weight
                else:
                    if f"{conn['mod2']}:{conn['conn2']}" == conn_key:
                        target_key = f"{conn['mod1']}:{conn['conn1']}"
                        if target_key in other_positions_dict:
                            weight = 1
                            pos_y = other_positions_dict[target_key]['y']
                            weighted_sum += pos_y * weight
                            total_weight += weight
        
        if total_weight == 0:
            return 0
        return weighted_sum / total_weight
    
    # Calcular posiciones iniciales sin optimización para referencia
    def calculate_initial_positions(connectors_dict, is_left=True):
        positions = {}
        current_y = margin_y
        
        for mod_name in sorted(connectors_dict.keys()):
            for conn_name in sorted(connectors_dict[mod_name].keys()):
                pins_dict = connectors_dict[mod_name][conn_name]
                
                def sort_key(pin):
                    try:
                        return (0, float(pin))
                    except:
                        return (1, pin)
                
                sorted_pins = sorted(pins_dict.keys(), key=sort_key)
                
                pins_list = []
                for pin in sorted_pins:
                    connections_list = pins_dict[pin]
                    signal = connections_list[0]['signal']
                    pins_list.append({
                        'pin': pin,
                        'signal': signal,
                        'connections': connections_list
                    })
                
                connector_height = len(pins_list) * pin_height + 60
                
                x_pos = margin_x if is_left else margin_x + connector_width + horizontal_gap
                
                key = f"{mod_name}:{conn_name}"
                positions[key] = {
                    'x': x_pos,
                    'y': current_y,
                    'width': connector_width,
                    'height': connector_height,
                    'pins': pins_list,
                    'module': mod_name,
                    'connector': conn_name
                }
                
                current_y += connector_height + connector_padding
            
            current_y += module_spacing
        
        return positions
    
    # Primero calcular posiciones temporales para ambos lados
    temp_left = calculate_initial_positions(left_connectors, is_left=True)
    temp_right = calculate_initial_positions(right_connectors, is_left=False)
    
    # Calcular el centro de masa para cada módulo del lado izquierdo
    left_modules = {}
    for mod_name in left_connectors.keys():
        center = calculate_center_of_mass(mod_name, True, temp_left, temp_right)
        left_modules[mod_name] = center
    
    # Ordenar módulos izquierdos por su centro de masa
    sorted_left_modules = sorted(left_modules.keys(), key=lambda m: left_modules[m])
    
    # Calcular posiciones finales de los conectores del lado izquierdo Y derecho coordinadamente
    # Los conectores izquierdos se posicionan según sus módulos derechos conectados
    
    left_positions = {}
    right_positions = {}
    current_y = margin_y
    
    for mod_name in sorted_left_modules:
        for conn_name in sorted(left_connectors[mod_name].keys()):
            left_key = f"{mod_name}:{conn_name}"
            pins_dict = left_connectors[mod_name][conn_name]
            
            # Calcular altura del conector izquierdo
            def sort_key(pin):
                try:
                    return (0, float(pin))
                except:
                    return (1, pin)
            
            sorted_pins = sorted(pins_dict.keys(), key=sort_key)
            
            pins_list = []
            for pin in sorted_pins:
                connections_list = pins_dict[pin]
                signal = connections_list[0]['signal']
                pins_list.append({
                    'pin': pin,
                    'signal': signal,
                    'connections': connections_list
                })
            
            connector_height = len(pins_list) * pin_height + 60
            
            # Guardar posición del conector izquierdo
            left_positions[left_key] = {
                'x': margin_x,
                'y': current_y,
                'width': connector_width,
                'height': connector_height,
                'pins': pins_list,
                'module': mod_name,
                'connector': conn_name
            }
            
            # Ahora procesar los módulos derechos que conectan con este conector izquierdo
            right_start_y = current_y
            
            # Obtener los módulos derechos conectados
            connected_right_modules = set()
            for conn in connections:
                if f"{conn['mod1']}:{conn['conn1']}" == left_key:
                    connected_right_modules.add(conn['mod2'])
            
            # Para cada módulo derecho conectado
            for right_mod_name in sorted(connected_right_modules):
                for right_conn_name in sorted(right_connectors.get(right_mod_name, {}).keys()):
                    right_key = f"{right_mod_name}:{right_conn_name}"
                    
                    # Verificar si este conector ya fue procesado
                    if right_key in right_positions:
                        continue
                    
                    # Verificar si tiene conexiones con este left_key
                    has_connection = False
                    for conn in connections:
                        if (f"{conn['mod1']}:{conn['conn1']}" == left_key and 
                            conn['mod2'] == right_mod_name and 
                            conn['conn2'] == right_conn_name):
                            has_connection = True
                            break
                    
                    if not has_connection:
                        continue
                    
                    right_pins_dict = right_connectors[right_mod_name][right_conn_name]
                    
                    sorted_right_pins = sorted(right_pins_dict.keys(), key=sort_key)
                    
                    right_pins_list = []
                    for pin in sorted_right_pins:
                        connections_list = right_pins_dict[pin]
                        signal = connections_list[0]['signal']
                        right_pins_list.append({
                            'pin': pin,
                            'signal': signal,
                            'connections': connections_list
                        })
                    
                    right_connector_height = len(right_pins_list) * pin_height + 60
                    right_x = margin_x + connector_width + horizontal_gap
                    
                    right_positions[right_key] = {
                        'x': right_x,
                        'y': right_start_y,
                        'width': connector_width,
                        'height': right_connector_height,
                        'pins': right_pins_list,
                        'module': right_mod_name,
                        'connector': right_conn_name
                    }
                    
                    right_start_y += right_connector_height + connector_padding
            
            # Calcular la siguiente posición Y para el próximo conector izquierdo
            # Debe estar debajo del último módulo derecho conectado a este conector izquierdo
            max_right_y = current_y + connector_height  # Al menos la altura del conector izquierdo
            
            for right_mod_name in connected_right_modules:
                for right_conn_name in sorted(right_connectors.get(right_mod_name, {}).keys()):
                    right_key = f"{right_mod_name}:{right_conn_name}"
                    if right_key in right_positions:
                        right_bottom = right_positions[right_key]['y'] + right_positions[right_key]['height']
                        max_right_y = max(max_right_y, right_bottom)
            
            # El próximo conector izquierdo empieza después del módulo derecho más bajo
            current_y = max_right_y + module_spacing
    
    # Procesar módulos derechos que no fueron procesados (sin conexión con ningún izquierdo)
    for mod_name in sorted(right_connectors.keys()):
        for conn_name in sorted(right_connectors[mod_name].keys()):
            right_key = f"{mod_name}:{conn_name}"
            
            if right_key in right_positions:
                continue
            
            pins_dict = right_connectors[mod_name][conn_name]
            
            def sort_key(pin):
                try:
                    return (0, float(pin))
                except:
                    return (1, pin)
            
            sorted_pins = sorted(pins_dict.keys(), key=sort_key)
            
            pins_list = []
            for pin in sorted_pins:
                connections_list = pins_dict[pin]
                signal = connections_list[0]['signal']
                pins_list.append({
                    'pin': pin,
                    'signal': signal,
                    'connections': connections_list
                })
            
            connector_height = len(pins_list) * pin_height + 60
            right_x = margin_x + connector_width + horizontal_gap
            
            right_positions[right_key] = {
                'x': right_x,
                'y': current_y,
                'width': connector_width,
                'height': connector_height,
                'pins': pins_list,
                'module': mod_name,
                'connector': conn_name
            }
            
            current_y += connector_height + connector_padding
    
    # Crear estructura XML de Draw.io
    mxfile = ET.Element('mxfile', host="app.diagrams.net", modified=f"2024-01-01T00:00:00.000Z", 
                        agent="Python Script", version="22.1.0", type="device")
    diagram = ET.SubElement(mxfile, 'diagram', id="interconnection", name="Interconexión")
    mxGraphModel = ET.SubElement(diagram, 'mxGraphModel', dx="1000", dy="1000", grid="1", 
                                  gridSize="10", guides="1", tooltips="1", connect="1", 
                                  arrows="1", fold="1", page="1", pageScale="1", 
                                  pageWidth="1600", pageHeight="2000", math="0", shadow="0")
    root = ET.SubElement(mxGraphModel, 'root')
    
    # Células base requeridas por Draw.io
    ET.SubElement(root, 'mxCell', id="0")
    ET.SubElement(root, 'mxCell', id="1", parent="0")
    
    cell_id = 2
    
    # Diccionario para guardar IDs de pines (para conectar cables)
    pin_ids = {}
    
    # Función para crear un conector en Draw.io
    def create_connector_box(pos_data, is_right=False):
        nonlocal cell_id
        
        x = pos_data['x']
        y = pos_data['y']
        w = pos_data['width']
        h = pos_data['height']
        module = pos_data['module']
        connector = pos_data['connector']
        pins = pos_data['pins']
        
        # Contenedor principal (rectángulo del conector)
        box_id = str(cell_id)
        cell_id += 1
        
        # Formato correcto para Draw.io: usar html=1 y texto simple con formato
        label_text = f"<b>{module}</b><br>{connector}"
        
        box_cell = ET.SubElement(root, 'mxCell', 
                                 id=box_id,
                                 value=label_text,
                                 style="rounded=1;whiteSpace=wrap;html=1;fillColor=#F5F5F5;strokeColor=#000000;strokeWidth=1;align=center;verticalAlign=top;fontSize=12;fontColor=#000000;arcSize=5;",
                                 vertex="1",
                                 parent="1")
        ET.SubElement(box_cell, 'mxGeometry', x=str(x), y=str(y), width=str(w), height=str(h), 
                     attrib={'as': 'geometry'})
        
        # Crear pines dentro del conector
        pin_start_y = 50
        
        for i, pin_data in enumerate(pins):
            pin_y = pin_start_y + i * pin_height
            pin = pin_data['pin']
            signal = pin_data['signal']
            connections_list = pin_data['connections']  # Lista de todas las conexiones
            
            # Pin (círculo pequeño) - uno solo para todas las conexiones
            pin_id = str(cell_id)
            cell_id += 1
            
            if is_right:
                # Conector derecho: pin a la izquierda (hacia el centro)
                pin_x = -10
            else:
                # Conector izquierdo: pin a la derecha (hacia el centro)
                pin_x = w
            
            # Guardar el ID del pin para TODAS las conexiones asociadas
            for conn in connections_list:
                row_idx = conn['row_idx']
                if is_right:
                    pin_ids[f"right:{module}:{connector}:{row_idx}"] = pin_id
                else:
                    pin_ids[f"left:{module}:{connector}:{row_idx}"] = pin_id
            
            pin_cell = ET.SubElement(root, 'mxCell',
                                    id=pin_id,
                                    value="",
                                    style="ellipse;whiteSpace=wrap;html=1;aspect=fixed;fillColor=#FFFFFF;strokeColor=#000000;strokeWidth=1;",
                                    vertex="1",
                                    parent=box_id)
            ET.SubElement(pin_cell, 'mxGeometry', x=str(pin_x), y=str(pin_y), 
                         width="8", height="8", attrib={'as': 'geometry'})
            
            # Texto de señal
            text_id = str(cell_id)
            cell_id += 1
            
            if is_right:
                # Conector derecho: mostrar [pin] primero, luego señal
                text_value = f"[{pin}] <b>{signal}</b>"
                text_x = 10
                text_align = "left"
            else:
                # Conector izquierdo: mostrar señal primero, luego [pin]
                text_value = f"<b>{signal}</b> [{pin}]"
                text_x = 10
                text_align = "right"
            
            text_cell = ET.SubElement(root, 'mxCell',
                                     id=text_id,
                                     value=text_value,
                                     style=f"text;html=1;strokeColor=none;fillColor=none;align={text_align};verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=10;fontColor=#000000;fontFamily=Courier New;",
                                     vertex="1",
                                     parent=box_id)
            ET.SubElement(text_cell, 'mxGeometry', x=str(text_x), y=str(pin_y - 3), 
                         width=str(w - 20), height="16", attrib={'as': 'geometry'})
    
    # Crear todos los conectores del lado izquierdo
    for key in sorted(left_positions.keys()):
        create_connector_box(left_positions[key], is_right=False)
    
    # Crear todos los conectores del lado derecho
    for key in sorted(right_positions.keys()):
        create_connector_box(right_positions[key], is_right=True)
    
    # Calcular offsets acumulativos para cada conector
    # Esto asegura que los cables de diferentes conectores no se solapen
    spacing = 12  # Espaciado entre carriles (líneas)
    base_offset = 30  # Distancia mínima desde el conector
    
    left_connector_offsets = {}
    current_left_offset = 0
    
    for mod_name in sorted_left_modules:
        for conn_name in sorted(left_connectors[mod_name].keys()):
            key = f"{mod_name}:{conn_name}"
            if key not in left_positions:
                continue
            
            left_connector_offsets[key] = current_left_offset
            # Incrementar el offset según el número de pines en este conector
            num_pins = len(left_positions[key]['pins'])
            current_left_offset += num_pins * spacing
            # Añadir espacio adicional entre conectores
            current_left_offset += spacing
    
    right_connector_offsets = {}
    current_right_offset = 0
    
    # Calcular offsets para módulos derechos en el orden en que fueron posicionados
    for right_key in right_positions.keys():
        right_connector_offsets[right_key] = current_right_offset
        num_pins = len(right_positions[right_key]['pins'])
        current_right_offset += num_pins * spacing
        current_right_offset += spacing
    
    # Crear las líneas de conexión con líneas ortogonales y offsets acumulativos
    for i, conn in enumerate(connections):
        left_key = f"{conn['mod1']}:{conn['conn1']}"
        right_key = f"{conn['mod2']}:{conn['conn2']}"
        
        # Verificar que ambos lados existan
        if left_key not in left_positions or right_key not in right_positions:
            continue
        
        row_idx = conn['row_idx']
        
        # Obtener IDs de los pines
        source_id = pin_ids.get(f"left:{conn['mod1']}:{conn['conn1']}:{row_idx}")
        target_id = pin_ids.get(f"right:{conn['mod2']}:{conn['conn2']}:{row_idx}")
        
        if not source_id or not target_id:
            continue
        
        # Color de la línea - siempre negro
        color = '#000000'
        
        # Obtener posiciones de los conectores
        left_pos = left_positions[left_key]
        right_pos = right_positions[right_key]
        
        # Encontrar el índice del pin en cada conector para calcular el offset
        left_pin_idx = None
        for idx, p in enumerate(left_pos['pins']):
            if any(c['row_idx'] == row_idx for c in p['connections']):
                left_pin_idx = idx
                break
        
        right_pin_idx = None
        for idx, p in enumerate(right_pos['pins']):
            if any(c['row_idx'] == row_idx for c in p['connections']):
                right_pin_idx = idx
                break
        
        if left_pin_idx is None or right_pin_idx is None:
            continue
        
        # Calcular posiciones Y de los pines
        pin_start_y = 50
        left_pin_y = left_pos['y'] + pin_start_y + left_pin_idx * pin_height
        right_pin_y = right_pos['y'] + pin_start_y + right_pin_idx * pin_height
        
        # Calcular posiciones X
        left_pin_x = left_pos['x'] + left_pos['width'] + 5
        right_pin_x = right_pos['x'] - 5
        
        # Calcular offset horizontal con offsets acumulativos del conector
        left_connector_base = left_connector_offsets.get(left_key, 0)
        right_connector_base = right_connector_offsets.get(right_key, 0)
        
        left_offset_x = left_pin_x + base_offset + left_connector_base + (left_pin_idx * spacing)
        right_offset_x = right_pin_x - base_offset - right_connector_base - (right_pin_idx * spacing)
        
        # Crear cable (edge) con waypoints para líneas ortogonales, sin flechas
        edge_id = str(cell_id)
        cell_id += 1
        
        edge_cell = ET.SubElement(root, 'mxCell',
                                 id=edge_id,
                                 value="",
                                 style=f"edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor={color};strokeWidth=1;startArrow=none;endArrow=none;exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=0;entryY=0.5;entryDx=0;entryDy=0;",
                                 edge="1",
                                 parent="1",
                                 source=source_id,
                                 target=target_id)
        
        geometry = ET.SubElement(edge_cell, 'mxGeometry', relative="1", attrib={'as': 'geometry'})
        
        # Crear array de puntos para el camino ortogonal
        array = ET.SubElement(geometry, 'Array', attrib={'as': 'points'})
        
        # Punto 1: Salir horizontalmente del pin izquierdo (centro) hasta el offset calculado
        ET.SubElement(array, 'mxPoint', x=str(left_offset_x), y=str(left_pin_y + 4))
        
        # Punto 2: Bajar/subir verticalmente hasta la altura del centro del pin derecho
        ET.SubElement(array, 'mxPoint', x=str(left_offset_x), y=str(right_pin_y + 4))
        
        # Punto 3: Ir horizontalmente hacia el offset del lado derecho
        ET.SubElement(array, 'mxPoint', x=str(right_offset_x), y=str(right_pin_y + 4))
    
    # Convertir a string XML
    xml_str = ET.tostring(mxfile, encoding='unicode')
    
    # Draw.io necesita que ciertos caracteres HTML estén escapados en el XML
    # pero interpreta <b>, <br>, etc. dentro de los valores de las celdas
    # Vamos a reemplazar nuestros marcadores temporales por el formato correcto
    
    # Formatear con minidom
    dom = minidom.parseString(xml_str)
    pretty_xml = dom.toprettyxml(indent="  ")
    
    # Remover líneas vacías extras y la declaración XML duplicada
    lines = [line for line in pretty_xml.split('\n') if line.strip()]
    # Mantener solo la primera declaración XML
    if lines[0].startswith('<?xml') and lines[1].startswith('<?xml'):
        lines.pop(1)
    pretty_xml = '\n'.join(lines)
    
    # Guardar archivo
    if output_drawio is None:
        output_drawio = Path(excel_file).stem + '_interconnection.drawio'
    
    with open(output_drawio, 'w', encoding='utf-8') as f:
        f.write(pretty_xml)
    
    print(f"✓ Diagrama Draw.io generado: {output_drawio}")
    print(f"  - {len(connections)} conexiones procesadas")
    print(f"  - {len(left_positions)} conectores lado izquierdo")
    print(f"  - {len(right_positions)} conectores lado derecho")
    print(f"\n📝 Para editar: Abre el archivo en Draw.io (https://app.diagrams.net)")
    return output_drawio


def main():
    if len(sys.argv) < 2:
        print("Uso: python interconnection_drawio.py <archivo_excel> [archivo_salida.drawio] [nombre_hoja]")
        print("\nEjemplo:")
        print("  python interconnection_drawio.py conexiones.xlsx")
        print("  python interconnection_drawio.py conexiones.xlsx diagrama.drawio")
        print("  python interconnection_drawio.py conexiones.xlsx diagrama.drawio Sheet2")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    output_drawio = sys.argv[2] if len(sys.argv) > 2 else None
    sheet_name = sys.argv[3] if len(sys.argv) > 3 else 0
    
    try:
        generate_drawio_diagram(excel_file, output_drawio, sheet_name)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
