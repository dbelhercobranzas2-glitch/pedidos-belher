from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from typing import List, Optional
import pandas as pd
import os
from datetime import datetime

app = FastAPI(title="API de Pedidos Móviles - Profit Plus")

# Configuración de CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =====================================================================
# MODELOS DE DATOS
# =====================================================================

class ItemPedido(BaseModel):
    articulo_id: str
    cantidad: int
    precio_unitario: float
    descuento_lineal: Optional[float] = 0.0

class PedidoSchema(BaseModel):
    cliente_id: str
    vendedor_id: str
    observaciones: Optional[str] = None
    articulos: List[ItemPedido]
    descuento_pronto_pago: Optional[float] = 0.0
    subtotal: Optional[float] = 0.0
    total: Optional[float] = 0.0

# =====================================================================
# VARIABLES GLOBALES
# =====================================================================

PRODUCTOS_CACHE = []
CLIENTES_CACHE = []
VENDEDORES_CACHE = []
MARCAS_CACHE = []
pedidos_procesados = []

# =====================================================================
# FUNCIONES DE CARGA DE DATOS
# =====================================================================

def cargar_productos_excel():
    """Carga los productos desde lista_precios.xls"""
    global PRODUCTOS_CACHE, MARCAS_CACHE
    archivo = "lista_precios.xls"
    
    if not os.path.exists(archivo):
        print(f"⚠️ Error: '{archivo}' no encontrado")
        return []
    
    try:
        df = pd.read_excel(archivo, header=None)
        productos = []
        marcas = []
        marca_actual = ""
        
        for _, fila in df.iterrows():
            try:
                primer_valor = str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
                
                # Detectar marca
                if primer_valor and primer_valor != "nan":
                    if not primer_valor.replace(".", "").replace("-", "").isdigit():
                        marca_actual = primer_valor
                        if marca_actual not in marcas:
                            marcas.append(marca_actual)
                        continue
                
                if marca_actual:
                    # Nombre del producto (columnas D-G: índices 3-6)
                    nombre_partes = []
                    for i in [3, 4, 5, 6]:
                        try:
                            valor = str(fila.iloc[i]).strip()
                            if valor and valor != "nan" and valor != "":
                                nombre_partes.append(valor)
                        except:
                            continue
                    nombre_producto = " ".join(nombre_partes) if nombre_partes else ""
                    
                    # Código (columna A o B)
                    codigo = str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
                    if not codigo or codigo == "nan":
                        codigo = str(fila.iloc[1]).strip() if pd.notna(fila.iloc[1]) else ""
                    
                    # PRECIO: Columnas L y M (índices 11 y 12)
                    precio = 0.0
                    for i in [11, 12]:
                        try:
                            if pd.notna(fila.iloc[i]):
                                precio = float(fila.iloc[i])
                                break
                        except:
                            continue
                    
                    if codigo and codigo != "nan" and nombre_producto and nombre_producto != "nan":
                        codigo = codigo.replace(".0", "").strip()
                        productos.append({
                            "articulo_id": codigo,
                            "descripcion": nombre_producto,
                            "precio_sin_iva": round(precio, 2),
                            "marca": marca_actual
                        })
            except:
                continue
        
        PRODUCTOS_CACHE = productos
        MARCAS_CACHE = marcas
        print(f"✅ Cargados {len(productos)} productos en {len(marcas)} marcas")
        return productos
    except Exception as e:
        print(f"❌ Error productos: {e}")
        return []

def cargar_clientes_excel():
    """Carga los clientes desde lista_clientes.xls"""
    global CLIENTES_CACHE
    archivo = "lista_clientes.xls"
    
    if not os.path.exists(archivo):
        print(f"⚠️ Error: '{archivo}' no encontrado")
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
                    codigo = codigo.replace(".0", "").strip()
                    clientes.append({
                        "cliente_id": codigo,
                        "nombre": nombre
                    })
            except:
                continue
        
        CLIENTES_CACHE = clientes
        print(f"✅ Cargados {len(clientes)} clientes")
        return clientes
    except Exception as e:
        print(f"❌ Error clientes: {e}")
        return []

def cargar_vendedores_excel():
    """Carga los vendedores desde lista_vendedores.xls"""
    global VENDEDORES_CACHE
    archivo = "lista_vendedores.xls"
    
    if not os.path.exists(archivo):
        print(f"⚠️ Error: '{archivo}' no encontrado")
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
                    codigo = codigo.replace(".0", "").strip()
                    vendedores.append({
                        "vendedor_id": codigo,
                        "nombre": nombre
                    })
            except:
                continue
        
        VENDEDORES_CACHE = vendedores
        print(f"✅ Cargados {len(vendedores)} vendedores")
        return vendedores
    except Exception as e:
        print(f"❌ Error vendedores: {e}")
        return []

# =====================================================================
# CARGAR TODOS LOS DATOS AL INICIAR
# =====================================================================

print("📂 Cargando datos...")
cargar_productos_excel()
cargar_clientes_excel()
cargar_vendedores_excel()
print(f"📊 Resumen:")
print(f"   - {len(PRODUCTOS_CACHE)} productos")
print(f"   - {len(CLIENTES_CACHE)} clientes")
print(f"   - {len(VENDEDORES_CACHE)} vendedores")
print(f"   - {len(MARCAS_CACHE)} marcas")
print("✅ Datos cargados correctamente")

# =====================================================================
# ENDPOINTS DE LA API
# =====================================================================

@app.get("/")
def inicio():
    return {
        "mensaje": "API de Pedidos Móviles - Profit Plus",
        "version": "2.0.0",
        "estado": "activa",
        "productos_cargados": len(PRODUCTOS_CACHE),
        "clientes_cargados": len(CLIENTES_CACHE),
        "vendedores_cargados": len(VENDEDORES_CACHE),
        "marcas_cargadas": len(MARCAS_CACHE)
    }

@app.get("/app")
async def servir_app():
    """Sirve la interfaz web (index.html) para acceso desde el móvil"""
    try:
        with open("index.html", "r", encoding="utf-8") as f:
            html = f.read()
        return HTMLResponse(content=html)
    except Exception as e:
        return {"error": f"No se encontró index.html: {e}"}

@app.get("/productos")
def obtener_productos():
    return PRODUCTOS_CACHE

@app.get("/productos/buscar/{texto}")
def buscar_productos(texto: str):
    texto = texto.lower().strip()
    resultados = []
    
    if len(texto) < 2:
        return resultados
    
    palabras = texto.split()
    
    for p in PRODUCTOS_CACHE:
        descripcion = p["descripcion"].lower()
        codigo = p["articulo_id"].lower()
        
        coincidencia = False
        
        if texto in descripcion or texto in codigo:
            coincidencia = True
        else:
            for palabra in palabras:
                if len(palabra) >= 2:
                    if palabra in descripcion or palabra in codigo:
                        coincidencia = True
                        break
        
        if not coincidencia:
            descripcion_sin_tildes = descripcion.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
            texto_sin_tildes = texto.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
            if texto_sin_tildes in descripcion_sin_tildes:
                coincidencia = True
        
        if coincidencia:
            resultados.append(p)
    
    return resultados[:50]

@app.get("/productos/marca/{marca}")
def obtener_productos_por_marca(marca: str):
    productos = [p for p in PRODUCTOS_CACHE if p.get("marca", "").upper() == marca.upper()]
    return productos

@app.get("/marcas")
def obtener_marcas():
    return MARCAS_CACHE

@app.get("/clientes")
def obtener_clientes():
    return CLIENTES_CACHE

@app.get("/clientes/buscar/{texto}")
def buscar_clientes(texto: str):
    texto = texto.lower().strip()
    resultados = []
    
    if len(texto) < 2:
        return resultados
    
    for c in CLIENTES_CACHE:
        if texto in c["cliente_id"].lower() or texto in c["nombre"].lower():
            resultados.append(c)
    
    return resultados[:20]

@app.get("/vendedores")
def obtener_vendedores():
    return VENDEDORES_CACHE

@app.post("/pedidos/crear")
def crear_pedido(pedido: PedidoSchema):
    try:
        for item in pedido.articulos:
            producto = next((p for p in PRODUCTOS_CACHE if p["articulo_id"] == item.articulo_id), None)
            if not producto:
                raise HTTPException(
                    status_code=400, 
                    detail=f"Artículo {item.articulo_id} no encontrado"
                )
        
        subtotal = 0.0
        for item in pedido.articulos:
            precio = item.precio_unitario
            descuento = getattr(item, 'descuento_lineal', 0) or 0
            precio_con_descuento = precio * (1 - descuento / 100)
            subtotal += precio_con_descuento * item.cantidad
        
        subtotal = round(subtotal, 2)
        
        descuento_pronto = getattr(pedido, 'descuento_pronto_pago', 0) or 0
        total = subtotal * (1 - descuento_pronto / 100)
        total = round(total, 2)
        
        pedido_id = f"PROF-{datetime.now().strftime('%Y%m%d')}-{len(pedidos_procesados)+1:04d}"
        
        pedido_data = {
            "id": pedido_id,
            "cliente": pedido.cliente_id,
            "vendedor": pedido.vendedor_id,
            "observaciones": pedido.observaciones,
            "articulos": [
                {
                    "articulo_id": item.articulo_id,
                    "cantidad": item.cantidad,
                    "precio_unitario": round(item.precio_unitario, 2),
                    "descuento_lineal": getattr(item, 'descuento_lineal', 0) or 0
                }
                for item in pedido.articulos
            ],
            "subtotal": subtotal,
            "descuento_pronto_pago": descuento_pronto,
            "total": total,
            "fecha": datetime.now().isoformat(),
            "estado": "PROCESADO"
        }
        
        pedidos_procesados.append(pedido_data)
        
        return {
            "status": "success",
            "mensaje": f"Pedido creado: {pedido_id}",
            "pedido_id": pedido_id,
            "subtotal": subtotal,
            "descuento_pronto_pago": descuento_pronto,
            "total": total,
            "articulos": len(pedido.articulos)
        }
        
    except HTTPException as e:
        raise e
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/pedidos/historial")
def historial_pedidos():
    return {
        "total_pedidos": len(pedidos_procesados),
        "pedidos": pedidos_procesados[-10:]
    }

@app.get("/pedidos/{pedido_id}")
def obtener_pedido(pedido_id: str):
    for pedido in pedidos_procesados:
        if pedido["id"] == pedido_id:
            return pedido
    raise HTTPException(status_code=404, detail="Pedido no encontrado")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8050, reload=True)