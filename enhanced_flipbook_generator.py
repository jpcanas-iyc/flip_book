import pandas as pd
from sqlalchemy import create_engine, text
import logging
import random
import os
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# -------------------------
# 1. Conectar a la BD
# -------------------------
def conectar_db():
    connection_string = (
        r'mssql+pyodbc:///?odbc_connect='
        r'Driver={SQL Server};'
        r'Server=PSR-S670-N\BIGDATA;'
        r'Database=DWH_INCOLMOTOS;'
        r'Trusted_Connection=yes'
    )
    engine = create_engine(connection_string)
    return engine


COLORES_DEPOSITO = {
    1: "#FF0000",
    2: "#0A2D82",
    3: "#616365",
    4: "#BCBDBC",
    5: "#FF6A00"
}


# -------------------------
# 2. Obtener los datos
# -------------------------


def color_por_deposito(deposito_key):
    return COLORES_DEPOSITO.get(deposito_key, "#999999")



def obtener_tareas(engine):
    query = """
    SELECT 
        T.Tarea_Project_Key,
        T.Project_Project_Key,
        T.Jefatura_Project_Key,
        T.Codigo_Esquema_Tarea,
        T.Codigo_Tarea,
        T.Nombre_Tarea_Project,
        T.Descripcion_Tarea_Project,
        T.Estado_Tarea_Project_key,
        T.Objs_Estrat_Area_Project_Key,
        T.Objs_Div_TI_Project_Key,
        T.Gcia_Project_Key,
        T.Categoria_YMC_key,
        T.Codigo_MTP_key,
        T.Porcentaje_Ejecucion,
        T.Deposito_Project_Key,
        T.Fecha_Inicio,
        T.Fecha_Fin,
        T.Fecha_Estimada_Entrega,
        T.Notas_IA_Project,
        D.Nombre_Deposito_Project,
        G.Nom_Gcia_Project,
        E.Nom_Estado_Tarea_Project
    FROM [DWH_INCOLMOTOS].[ti].[Dim_Tareas_Project] T
    LEFT JOIN [DWH_INCOLMOTOS].[ti].[Dim_Depositos_Project] D
        ON T.Deposito_Project_Key = D.Deposito_Project_Key
    LEFT JOIN [DWH_INCOLMOTOS].[ti].[Dim_Gcias_Involucradas_Project] G
        ON T.Gcia_Project_Key = G.Gcia_Project_Key
    LEFT JOIN [DWH_INCOLMOTOS].[ti].[Dim_Estado_Tareas_Project] E
        ON T.Estado_Tarea_Project_key = E.Estado_Tarea_Project_key
    WHERE T.Jefatura_Project_Key = 2
    """
    df = pd.read_sql(query, engine)
    return df

# -------------------------
# 3. Obtener imágenes aleatorias
# -------------------------
def obtener_imagenes_aleatorias(carpeta_img="img", cantidad=None):
    """
    Obtiene lista de imágenes de la carpeta especificada
    """
    extensiones = ['.jpg', '.jpeg', '.png', '.gif', '.webp']
    imagenes = []
    
    if os.path.exists(carpeta_img):
        for archivo in os.listdir(carpeta_img):
            if any(archivo.lower().endswith(ext) for ext in extensiones):
                imagenes.append(f"../{carpeta_img}/{archivo}")
    
    if not imagenes:
        logger.warning(f"No se encontraron imágenes en {carpeta_img}")
        return []
    
    # Si se especifica cantidad, seleccionar aleatoriamente
    if cantidad and len(imagenes) > cantidad:
        return random.sample(imagenes, cantidad)
    
    return imagenes

# -------------------------
# 4. Calcular estadísticas
# -------------------------
def generar_estadisticas(df):
    total_tareas = len(df)
    
    # Agrupaciones
    por_deposito = df.groupby("Nombre_Deposito_Project").size().to_dict()
    por_gerencia = df.groupby("Nom_Gcia_Project").size().to_dict()
    por_estado = df.groupby("Nom_Estado_Tarea_Project").size().to_dict()
    
    estadisticas = {
        "total_tareas": total_tareas,
        "por_deposito": por_deposito,
        "por_gerencia": por_gerencia,
        "por_estado": por_estado,
    }
    return estadisticas

# -------------------------
# 5. Generadores de layouts
# -------------------------


def pagina_foto():
    return """
    <div class="page foto-fija-page">
        <img src="../img/img_0283.jpeg" alt="Fondo">
    </div>
    """


def pagina_tabla_contenido(df):
    items = ""

    for _, row in df.iterrows():
        color = COLORES_DEPOSITO.get(row["Deposito_Project_Key"], "#999999")

        items += f"""
        <div class="toc-item">
            <span class="toc-dot" style="background:{color}"></span>
            <div class="toc-text">
                <div class="toc-title">{row['Nombre_Tarea_Project']}</div>
                <div class="toc-meta">
                    {row['Nombre_Deposito_Project']} · {row['Nom_Gcia_Project']}
                </div>
            </div>
        </div>
        """

    return f"""
    <div class="page toc-page">
        <div class="toc-container">
            <h2>TABLA DE CONTENIDO</h2>
            <div class="toc-subtitle">
                Proyectos incluidos en el informe
            </div>
            <div class="toc-list">
                {items}
            </div>
        </div>
    </div>
    """





def layout_left_text(row, img_url):
    """Layout con texto a la izquierda e imagen a la derecha"""
    return f"""
    <div class="double layout-left-text">
        <div class="text-column">
            <div class="text-content">
                <div class="accent-line"></div>
                <div class="task-code">{row['Codigo_Tarea']}</div>
                <h2>{row['Nombre_Tarea_Project']}</h2>
                <div class="task-info">
                    <div class="task-label">Depósito</div>
                    <div class="task-value">{row['Nombre_Deposito_Project']}</div>
                </div>
                <div class="task-info">
                    <div class="task-label">Gerencia</div>
                    <div class="task-value">{row['Nom_Gcia_Project']}</div>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: {row['Porcentaje_Ejecucion']*100}%"></div>
                </div>
                <div class="percentage-badge">{row['Porcentaje_Ejecucion']*100:.0f}%</div>
                <div class="task-info">
                    <div class="date-display">Inicio: {row['Fecha_Inicio']}</div>
                    <div class="date-display">Fin: {row['Fecha_Fin']}</div>
                </div>
            </div>
        </div>
        <div class="image-column">
            <img src="{img_url}" alt="Background">
        </div>
    </div>
    """

def layout_right_text(row, img_url):
    """Layout con texto a la derecha e imagen a la izquierda"""
    return f"""
    <div class="double layout-right-text">
        <div class="text-column">
            <div class="text-content">
                <h2>{row['Nombre_Tarea_Project']}</h2>
                <div class="divider"></div>
                <div class="task-label">Código: {row['Codigo_Tarea']}</div>
                <div class="task-info">
                    <strong>Estado:</strong> {row['Nom_Estado_Tarea_Project']}
                </div>
                <div class="task-info">
                    <strong>Notas:</strong><br>{str(row['Notas_IA_Project'] or '')[:100]}
                </div>
                <div class="task-info">
                    <strong>Progreso:</strong> {row['Porcentaje_Ejecucion']*100:.0f}%
                </div>
                <div class="task-info">
                    <div class="date-display">Entrega: {row['Fecha_Estimada_Entrega']}</div>
                </div>
            </div>
        </div>
        <div class="image-column">
            <img src="{img_url}" alt="Background">
        </div>
    </div>
    """

def layout_diagonal(row, img_url):
    """Layout con división diagonal"""
    return f"""
    <div class="double layout-diagonal">
        <div class="diagonal-bg"></div>
        <div class="content-wrapper">
            <div class="text-section">
                <div class="task-code" style="color: var(--accent-red);">{row['Codigo_Tarea']}</div>
                <h2>{row['Nombre_Tarea_Project']}</h2>
                <div class="task-info">
                    <strong>Depósito:</strong> {row['Nombre_Deposito_Project']}
                </div>
                <div class="task-info">
                    <strong>Gerencia:</strong> {row['Nom_Gcia_Project']}
                </div>
                <div class="percentage-badge">{row['Porcentaje_Ejecucion']*100:.0f}% Completado</div>
            </div>
            <div class="image-section">
                <img src="{img_url}" alt="Background">
            </div>
        </div>
    </div>
    """

def layout_center_margins(row, img_url):
    """Layout con márgenes laterales y contenido central"""
    return f"""
    <div class="double layout-center-margins">
        <div class="left-margin">
            <div class="margin-text">{row['Codigo_Tarea']}</div>
        </div>
        <div class="center-content">
            <div class="image-container">
                <img src="{img_url}" alt="Background">
                <div class="text-overlay">
                    <h2>{row['Nombre_Tarea_Project']}</h2>
                    <div class="task-info">
                        <strong>{row['Nom_Estado_Tarea_Project']}</strong> - {row['Porcentaje_Ejecucion']*100:.0f}%
                    </div>
                    <div class="task-info">
                        {str(row['Notas_IA_Project'] or '')[:100]}...
                    </div>
                </div>
            </div>
        </div>
        <div class="right-margin">
            <div class="margin-text">PROYECTO</div>
        </div>
    </div>
    """

def layout_full_overlay(row, img_url):
    """Layout con imagen de fondo completa y overlay de texto"""
    return f"""
    <div class="double layout-full-overlay">
        <div class="background-image">
            <img src="{img_url}" alt="Background">
        </div>
        <div class="gradient-overlay"></div>
        <div class="content-box">
            <h2>
                {row['Nombre_Tarea_Project'].split()[0]}
                <span class="red-accent">{' '.join(row['Nombre_Tarea_Project'].split()[1:])}</span>
            </h2>
            <div class="task-info" style="font-size: 16px; margin-bottom: 15px;">
                <strong>Código:</strong> {row['Codigo_Tarea']}
            </div>
            <div class="task-info">
                <strong>Depósito:</strong> {row['Nombre_Deposito_Project']}
            </div>
            <div class="task-info">
                <strong>Estado:</strong> {row['Nom_Estado_Tarea_Project']}
            </div>
            <div class="percentage-badge" style="margin-top: 20px; font-size: 18px;">
                {row['Porcentaje_Ejecucion']*100:.0f}%
            </div>
        </div>
    </div>
    """

def layout_grid(row, img_url):
    """Layout con grid asimétrico"""
    return f"""
    <div class="double layout-grid">
        <div class="text-area">
            <h2>{row['Nombre_Tarea_Project']}</h2>
            <div class="task-info">
                <div class="task-label">Código</div>
                <div class="task-value">{row['Codigo_Tarea']}</div>
            </div>
            <div class="task-info">
                <div class="task-label">Notas</div>
                <div class="task-value">{str(row['Notas_IA_Project'] or '')[:100]}..</div>
            </div>
            <div class="highlight-box">
                <strong>Progreso:</strong> {row['Porcentaje_Ejecucion']*100:.0f}%<br>
                <strong>Estado:</strong> {row['Nom_Estado_Tarea_Project']}
            </div>
        </div>
        <div class="image-area-1">
            <img src="{img_url}" alt="Background">
        </div>
        <div class="image-area-2">
            <div>
                <div class="task-label">Depósito</div>
                <div style="font-size: 18px; font-weight: 700; margin-bottom: 10px;">
                    {row['Nombre_Deposito_Project']}
                </div>
                <div class="task-label">Gerencia</div>
                <div style="font-size: 16px;">
                    {row['Nom_Gcia_Project']}
                </div>
            </div>
        </div>
    </div>
    """

# -------------------------
# 6. Generar flipbook HTML
# -------------------------
def generar_flipbook(df, output="output/flipbook.html"):
    estadisticas = generar_estadisticas(df)
    paginas = []
    
    # Obtener todas las imágenes disponibles
    imagenes = obtener_imagenes_aleatorias()
    if not imagenes:
        logger.warning("No hay imágenes disponibles. Se usará un placeholder.")
        imagenes = ["../img/placeholder.jpg"] * len(df)
    
    # Definir layouts disponibles organizados por color de fondo
    layouts_azules = [
        layout_left_text,      # Azul
        layout_full_overlay,   # Azul
    ]
    
    layouts_oscuros = [
        layout_right_text,     # Gris oscuro
        layout_center_margins, # Negro
        layout_grid,           # Negro
    ]
    
    layouts_claros = [
        layout_diagonal,       # Blanco/claro
    ]
    
    # --- Portada ---
    img_portada = "../img/Portada.jpeg"
    portada = f"""
    <div class="double portada-custom">
        <div class="bg-image">
            <img src="{img_portada}" alt="Portada">
        </div>
        <div class="overlay-content">
            <h1>INFORME DE PROYECTOS</h1>
            <div class="accent-bar"></div>
            <div class="subtitle">Infraestructura De Datos</div>
            <div style="margin-top: 40px; font-size: 18px; font-family: 'Montserrat', sans-serif;">
                Total de Proyectos: <strong>{estadisticas['total_tareas']}</strong>
            </div>
        </div>
    </div>
    """
    paginas.append(portada)

    # Página 2 (izquierda)
    paginas.append(pagina_foto())

    # Página 3 (derecha)
    paginas.append(pagina_tabla_contenido(df))



    
    
    # --- Página de estadísticas CON GRÁFICAS DE BARRAS ---
    
    # Gráfica de barras para Depósitos
    max_deposito = max(estadisticas['por_deposito'].values()) if estadisticas['por_deposito'] else 1
    items_deposito_barras = ''.join([
        f'''<div class="bar-item">
            <div class="bar-label">
                <span class="bar-name">{k}</span>
                <span class="bar-value">{v}</span>
            </div>
            <div class="bar-track">
                <div class="bar-fill color-1" style="width: {(v/max_deposito)*100}%"></div>
            </div>
        </div>'''
        for k, v in estadisticas['por_deposito'].items()
    ])
    
    # Gráfica de barras para Gerencias con colores alternos
    max_gerencia = max(estadisticas['por_gerencia'].values()) if estadisticas['por_gerencia'] else 1
    items_gerencia_barras = ''.join([
        f'''<div class="bar-item">
            <div class="bar-label">
                <span class="bar-name">{k}</span>
                <span class="bar-value">{v}</span>
            </div>
            <div class="bar-track">
                <div class="bar-fill color-{(i % 4) + 1}" style="width: {(v/max_gerencia)*100}%"></div>
            </div>
        </div>'''
        for i, (k, v) in enumerate(estadisticas['por_gerencia'].items())
    ])
    
    items_estado = ''.join([
        f'<div class="state-badge"><span class="state-badge-value">{v}</span><span class="state-badge-label">{k}</span></div>'
        for k, v in estadisticas['por_estado'].items()
    ])
    
    estadistica_html = f"""
    <div class="double stats-premium">
        <div class="stats-container">
            
            <!-- Hero Section -->
            <div class="stats-hero-section">
                <div class="stats-kicker">Resumen Ejecutivo</div>
                <h2>PANORAMA GENERAL</h2>
                <div class="hero-display">
                    <div class="hero-number">{estadisticas['total_tareas']}</div>
                    <div class="hero-label">Tareas Registradas</div>
                </div>
            </div>
            
            <!-- Content Grid -->
            <div class="stats-content-grid">
                
                <!-- Card 1: Por Depósito CON BARRAS -->
                <div class="stat-glass-card">
                    <div class="stat-card-title">Por Depósito</div>
                    <div class="bar-chart">
                        {items_deposito_barras}
                    </div>
                </div>
                
                <!-- Card 2: Por Gerencia CON BARRAS -->
                <div class="stat-glass-card">
                    <div class="stat-card-title">Distribución por Gerencia</div>
                    <div class="bar-chart">
                        {items_gerencia_barras}
                    </div>
                </div>
                
            </div>
            
            <!-- Card 3: Por Estado (Full Width Below) -->
            <div class="stat-glass-card highlight" style="margin-top: 14px;">
                <div class="stat-card-title">Distribución por Estado</div>
                <div class="states-grid">
                    {items_estado}
                </div>
            </div>
            
        </div>
    </div>
    """
    paginas.append(estadistica_html)
    
    # --- Páginas de las tareas con layouts alternando colores ---
    color_anterior = None  # Para tracking del color previo
    
    for idx, (_, row) in enumerate(df.iterrows()):
        # Seleccionar imagen aleatoria
        img_url = random.choice(imagenes)
        
        # Alternar entre azul/oscuro y claro/diagonal
        # Si el anterior fue azul, siguiente debe ser oscuro o claro
        # Si el anterior fue oscuro, siguiente puede ser azul o claro
        
        if color_anterior == 'azul':
            # Siguiente debe ser oscuro o claro
            grupo = random.choice([layouts_oscuros, layouts_claros])
            color_anterior = 'oscuro' if grupo == layouts_oscuros else 'claro'
        elif color_anterior == 'oscuro':
            # Siguiente debe ser azul o claro
            grupo = random.choice([layouts_azules, layouts_claros])
            color_anterior = 'azul' if grupo == layouts_azules else 'claro'
        else:  # None o 'claro'
            # Puede ser cualquiera excepto repetir claro
            if color_anterior == 'claro':
                grupo = random.choice([layouts_azules, layouts_oscuros])
                color_anterior = 'azul' if grupo == layouts_azules else 'oscuro'
            else:
                # Primera iteración
                grupo = random.choice([layouts_azules, layouts_oscuros, layouts_claros])
                if grupo == layouts_azules:
                    color_anterior = 'azul'
                elif grupo == layouts_oscuros:
                    color_anterior = 'oscuro'
                else:
                    color_anterior = 'claro'
        
        layout_func = random.choice(grupo)
        
        # Generar página con el layout seleccionado
        pagina = layout_func(row, img_url)
        paginas.append(pagina)
    
    # --- HTML final ---
    html = f"""
    <!doctype html>
    <html lang="es">
    <head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=1050, user-scalable=no" />
    <link rel="stylesheet" href="../css/enhanced-flipbook.css">
    <script type="text/javascript" src="../extras/jquery.min.1.7.js"></script>
    <script type="text/javascript" src="../extras/modernizr.2.5.3.min.js"></script>
    </head>
    <body>
    <div class="flipbook-viewport">
        <div class="container">
        <div class="flipbook">
            {''.join(paginas)}
        </div>

        <!-- FLECHAS DE NAVEGACIÓN -->
        <div class="nav-arrow nav-left" id="prevPage">&#10094;</div>
        <div class="nav-arrow nav-right" id="nextPage">&#10095;</div>

        </div>
    </div>

    <script type="text/javascript" src="../lib/turn.min.js"></script>

    <!-- Inicialización del flipbook -->
    <script type="text/javascript">
        var $flipbook = $('.flipbook');

        $flipbook.turn({{ 
            elevation: 50, 
            gradients: true, 
            autoCenter: true,
            duration: 1000,
            acceleration: true,
            display: 'double',
            when: {{
                turning: function(event, page, view) {{
                    console.log('Página actual:', page);
                }}
            }}
        }});
    </script>

    <!-- LÓGICA DE FLECHAS -->
    <script type="text/javascript">
        document.getElementById('nextPage').addEventListener('click', function () {{
            $flipbook.turn('next');
        }});

        document.getElementById('prevPage').addEventListener('click', function () {{
            $flipbook.turn('previous');
        }});

        // Navegación con teclado
        document.addEventListener('keydown', function (e) {{
            if (e.key === 'ArrowRight') {{
                $flipbook.turn('next');
            }}
            if (e.key === 'ArrowLeft') {{
                $flipbook.turn('previous');
            }}
        }});

        // Ocultar flechas en extremos
        $flipbook.bind('turned', function (event, page) {{
            var total = $flipbook.turn('pages');

            document.getElementById('prevPage').style.display =
                page <= 1 ? 'none' : 'block';

            document.getElementById('nextPage').style.display =
                page >= total ? 'none' : 'block';
        }});
    </script>

    </body>
    </html>
    """

    # Crear directorio de salida si no existe
    os.makedirs(os.path.dirname(output), exist_ok=True)
    
    with open(output, "w", encoding="utf-8") as f:
        f.write(html)
    logger.info(f"✅ Flipbook generado en {output}")
    logger.info(f"📊 Total de páginas: {len(paginas)}")
    logger.info(f"🎨 Layouts con alternancia de colores activada")
    logger.info(f"🖼️  Estadísticas con imagen de fondo incluida")

# -------------------------
# 7. Ejecutar
# -------------------------
if __name__ == "__main__":
    try:
        engine = conectar_db()
        df = obtener_tareas(engine)
        generar_flipbook(df)
    except Exception as e:
        logger.error(f"❌ Error: {e}")
        raise