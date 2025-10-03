import pandas as pd
from sqlalchemy import create_engine, text
import logging

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

# -------------------------
# 2. Obtener los datos
# -------------------------
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
# 3. Calcular estadísticas
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
# 4. Generar flipbook HTML
# -------------------------
def generar_flipbook(df, output="output/flipbook.html"):
    estadisticas = generar_estadisticas(df)
    paginas = []

    # --- Portada ---
    portada = """
    <div class="double portada">
        <div class="page-content" style="padding:0;">
            <img src="../img/portada.jpg" alt="Portada" 
                 style="width:100%; height:100%; object-fit:cover; border:none;">
        </div>
    </div>
    """
    paginas.append(portada)

    # --- Página de estadísticas ---
    estadistica_html = f"""
    <div class="double">
        <div class="page-content">
            <h2>📊 Estadísticas Generales</h2>
            <p><b>Total de Tareas:</b> {estadisticas['total_tareas']}</p>

            <h3>Por Depósito:</h3>
            <ul>
                {''.join([f"<li>{k}: {v}</li>" for k, v in estadisticas['por_deposito'].items()])}
            </ul>

            <h3>Por Gerencia:</h3>
            <ul>
                {''.join([f"<li>{k}: {v}</li>" for k, v in estadisticas['por_gerencia'].items()])}
            </ul>

            <h3>Por Estado de Tarea:</h3>
            <ul>
                {''.join([f"<li>{k}: {v}</li>" for k, v in estadisticas['por_estado'].items()])}
            </ul>
        </div>
    </div>
    """
    paginas.append(estadistica_html)

    # --- Páginas de las tareas ---
    for _, row in df.iterrows():
        pagina = f"""
        <div class="double">
            <div class="page-content">
                <div class="page-header">
                    <p>Ejecución: {row['Porcentaje_Ejecucion']*100:.0f}%</p>
                    <p>Fecha Fin: {row['Fecha_Fin']}</p>
                </div>
                <h2>{row['Codigo_Tarea']} - {row['Nombre_Tarea_Project']}</h2>
                <p><b>Notas:</b> {row['Notas_IA_Project']}</p>
                <p><b>Inicio:</b> {row['Fecha_Inicio']} &nbsp;&nbsp; → &nbsp;&nbsp; 
                   <b>Entrega Estimada:</b> {row['Fecha_Estimada_Entrega']}</p>
            </div>
        </div>
        """
        paginas.append(pagina)

    # --- HTML final ---
    html = f"""
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8"/>
      <meta name="viewport" content="width=1050, user-scalable=no" />
      <link rel="stylesheet" href="../css/double-page.css">
      <script type="text/javascript" src="../extras/jquery.min.1.7.js"></script>
      <script type="text/javascript" src="../extras/modernizr.2.5.3.min.js"></script>
    </head>
    <body>
      <div class="flipbook-viewport">
        <div class="container">
          <div class="flipbook">
            {''.join(paginas)}
          </div>
        </div>
      </div>

      <script type="text/javascript" src="../lib/turn.min.js"></script>
      <script type="text/javascript">
        $('.flipbook').turn({{ 
            elevation:50, 
            gradients:true, 
            autoCenter:true 
        }});
      </script>
    </body>
    </html>
    """

    with open(output, "w", encoding="utf-8") as f:
        f.write(html)
    logger.info(f"✅ Flipbook generado en {output}")

# -------------------------
# 5. Ejecutar
# -------------------------
if __name__ == "__main__":
    engine = conectar_db()
    df = obtener_tareas(engine)
    generar_flipbook(df)
