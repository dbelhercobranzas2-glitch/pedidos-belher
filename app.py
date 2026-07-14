from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

# =====================================================================
# RUTA DEL PROYECTO
# =====================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# =====================================================================
# FUNCIONES PARA LEER EXCEL
# =====================================================================

def cargar_productos():
    archivo = os.path.join(BASE_DIR, "lista_precios.xls")
    if not os.path.exists(archivo):
        print(f"⚠️ Archivo no encontrado: {archivo}")
        return []
    try:
        df = pd.read_excel(archivo, header=None)
        productos = []
        marca_actual = ""
        
        for _, fila in df.iterrows():
            try:
                col_a = str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
                
                if col_a and col_a != "nan" and not col_a.replace(".", "").replace("-", "").isdigit():
                    marca_actual = col_a
                    continue
                
                codigo = col_a
                if not codigo or codigo == "nan":
                    codigo = str(fila.iloc[1]).strip() if pd.notna(fila.iloc[1]) else ""
                
                descripcion_partes = []
                for i in [3, 4, 5, 6]:
                    try:
                        valor = str(fila.iloc[i]).strip()
                        if valor and valor != "nan" and valor != "":
                            descripcion_partes.append(valor)
                    except:
                        continue
                descripcion = " ".join(descripcion_partes) if descripcion_partes else ""
                
                precio = 0.0
                for i in [11, 12]:
                    try:
                        if pd.notna(fila.iloc[i]):
                            precio = float(fila.iloc[i])
                            break
                    except:
                        continue
                
                if codigo and codigo != "nan" and codigo != "" and descripcion and descripcion != "nan":
                    productos.append({
                        "articulo_id": codigo,
                        "descripcion": descripcion,
                        "precio_sin_iva": round(precio, 2),
                        "marca": marca_actual
                    })
            except:
                continue
        
        print(f"✅ Productos cargados: {len(productos)}")
        return productos
    except Exception as e:
        print(f"❌ Error cargando productos: {e}")
        return []

def cargar_clientes():
    archivo = os.path.join(BASE_DIR, "lista_clientes.xls")
    if not os.path.exists(archivo):
        print(f"⚠️ Archivo no encontrado: {archivo}")
        return []
    try:
        df = pd.read_excel(archivo, header=None)
        clientes = []
        for _, fila in df.iterrows():
            try:
                codigo = str(fila.iloc[1]).strip() if pd.notna(fila.iloc[1]) else ""
                nombre_partes = []
                for i in [4, 5, 6, 7, 8]:
                    try:
                        valor = str(fila.iloc[i]).strip()
                        if valor and valor != "nan" and valor != "":
                            nombre_partes.append(valor)
                    except:
                        continue
                nombre = " ".join(nombre_partes) if nombre_partes else ""
                
                if codigo and codigo != "nan" and codigo != "" and nombre and nombre != "nan" and nombre != "":
                    clientes.append({
                        "cliente_id": codigo,
                        "nombre": nombre
                    })
            except:
                continue
        print(f"✅ Clientes cargados: {len(clientes)}")
        return clientes
    except Exception as e:
        print(f"❌ Error cargando clientes: {e}")
        return []

def cargar_vendedores():
    archivo = os.path.join(BASE_DIR, "lista_vendedores.xls")
    if not os.path.exists(archivo):
        print(f"⚠️ Archivo no encontrado: {archivo}")
        return []
    try:
        df = pd.read_excel(archivo, header=None)
        vendedores = []
        for _, fila in df.iterrows():
            try:
                codigo = str(fila.iloc[1]).strip() if pd.notna(fila.iloc[1]) else ""
                nombre_partes = []
                for i in [3, 4, 5, 6, 7]:
                    try:
                        valor = str(fila.iloc[i]).strip()
                        if valor and valor != "nan" and valor != "":
                            nombre_partes.append(valor)
                    except:
                        continue
                nombre = " ".join(nombre_partes) if nombre_partes else ""
                
                if codigo and codigo != "nan" and codigo != "" and nombre and nombre != "nan" and nombre != "":
                    vendedores.append({
                        "vendedor_id": codigo,
                        "nombre": nombre
                    })
            except:
                continue
        print(f"✅ Vendedores cargados: {len(vendedores)}")
        return vendedores
    except Exception as e:
        print(f"❌ Error cargando vendedores: {e}")
        return []

# =====================================================================
# DATOS EN MEMORIA
# =====================================================================
PRODUCTOS = cargar_productos()
CLIENTES = cargar_clientes()
VENDEDORES = cargar_vendedores()
MARCAS = sorted(list(set([p.get("marca", "") for p in PRODUCTOS if p.get("marca")])))

print(f"📊 Resumen:")
print(f"   - {len(PRODUCTOS)} productos")
print(f"   - {len(CLIENTES)} clientes")
print(f"   - {len(VENDEDORES)} vendedores")
print(f"   - {len(MARCAS)} marcas")

# =====================================================================
# ENDPOINTS
# =====================================================================

@app.route('/')
def inicio():
    return jsonify({
        "mensaje": "API de Pedidos Móviles - Belher",
        "productos": len(PRODUCTOS),
        "clientes": len(CLIENTES),
        "vendedores": len(VENDEDORES)
    })

@app.route('/app')
def servir_app():
    return send_from_directory('.', 'index.html')

@app.route('/productos')
def obtener_productos():
    return jsonify(PRODUCTOS)

@app.route('/clientes')
def obtener_clientes():
    return jsonify(CLIENTES)

@app.route('/vendedores')
def obtener_vendedores():
    return jsonify(VENDEDORES)

@app.route('/marcas')
def obtener_marcas():
    return jsonify(MARCAS)

@app.route('/productos/buscar/<texto>')
def buscar_productos(texto):
    texto = texto.lower().strip()
    resultados = []
    for p in PRODUCTOS:
        if texto in p["descripcion"].lower() or texto in p["articulo_id"].lower():
            resultados.append(p)
    return jsonify(resultados[:50])

@app.route('/clientes/buscar/<texto>')
def buscar_clientes(texto):
    texto = texto.lower().strip()
    resultados = []
    for c in CLIENTES:
        if texto in c["cliente_id"].lower() or texto in c["nombre"].lower():
            resultados.append(c)
    return jsonify(resultados[:20])

@app.route('/pedidos/crear', methods=['POST'])
def crear_pedido():
    data = request.get_json()
    articulos = data.get('articulos', [])
    total = sum(item['cantidad'] * item['precio_unitario'] for item in articulos)
    pedido_id = f"PROF-{datetime.now().strftime('%Y%m%d')}-0001"
    return jsonify({
        "status": "success",
        "mensaje": f"Pedido creado: {pedido_id}",
        "pedido_id": pedido_id,
        "total": round(total, 2)
    })

@app.route('/pedidos/centralizados', methods=['GET'])
def pedidos_centralizados():
    """Endpoint para centralizar pedidos por vendedor"""
    # Por ahora retorna un mensaje simple
    return jsonify({
        "mensaje": "Centralización de pedidos (en desarrollo)",
        "vendedores": []
    })

@app.route('/pedidos/historial', methods=['GET'])
def historial_pedidos():
    """Endpoint para ver historial de pedidos"""
    return jsonify({
        "mensaje": "Historial de pedidos (en desarrollo)",
        "pedidos": []
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
