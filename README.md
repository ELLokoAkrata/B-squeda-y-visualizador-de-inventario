# 📦 Sistema Profesional de Gestión de Inventario

Sistema completo de gestión de inventario desarrollado con Streamlit, diseñado para especialistas en control de inventarios. Incluye validaciones automáticas, análisis ABC, alertas de stock, dashboards interactivos y reportes profesionales.

## 🚀 Características Principales

### ✅ Validaciones Automáticas
- **Validación de unicidad:** Previene items duplicados en el inventario
- **Validación de valores:** No permite stocks o precios negativos
- **Cálculo automático de totales:** TOTAL = CANT. × PRECIO UNIT
- **Campos obligatorios:** ITEM, DESCRIPCIÓN, CANT. y PRECIO UNIT son requeridos
- **Timestamps automáticos:** Registro de fecha de creación y modificación

### 📊 Dashboard y KPIs en Tiempo Real
- Valor total del inventario
- Número total de items y cantidad total
- Alertas de items con stock bajo
- Distribución por categorías y ubicaciones
- Precio promedio del inventario
- Visualizaciones interactivas con gráficos

### 🔍 Búsqueda y Filtros Avanzados
- **Búsqueda global:** Busca texto en todas las columnas
- **Filtros por categoría:** Filtra por categoría de producto
- **Filtros por ubicación:** Localiza items por ubicación física
- **Filtros por status:** Activo, Inactivo, etc.
- **Filtros por rangos:** Precio mínimo/máximo, cantidad mínima
- **Exportación de resultados:** Descarga solo los items filtrados

### 📈 Reportes y Analíticas Profesionales

#### Clasificación ABC (Análisis de Pareto)
Sistema automático de clasificación de inventario:
- **Clase A:** Items que representan el 80% del valor (alta prioridad)
- **Clase B:** Items que representan el 15% del valor (prioridad media)
- **Clase C:** Items que representan el 5% del valor (baja prioridad)

#### Visualizaciones Interactivas
- Gráfico de valor total por categoría
- Distribución ABC en gráfico de pastel
- Top 10 items por valor total
- Distribución de items por ubicación

### ⚠️ Sistema de Alertas
- **Alertas de stock bajo:** Notificación cuando CANT. ≤ STOCK MÍNIMO
- **Indicadores visuales:** Alertas en dashboard y sidebar
- **Stock crítico:** Lista detallada de items que requieren reorden

### 🕐 Auditoría y Trazabilidad
- **FECHA_CREACIÓN:** Timestamp de cuando se agregó el item
- **FECHA_MODIFICACIÓN:** Timestamp de última modificación
- **Historial en session state:** Tracking de cambios durante la sesión

## 📋 Modelo de Datos Extendido

### Columnas Principales
| Columna | Tipo | Descripción | Requerido |
|---------|------|-------------|-----------|
| ITEM | String | Identificador único | ✅ |
| DESCRIPCIÓN | String | Descripción del producto | ✅ |
| MARCA | String | Fabricante/marca | |
| MODELO | String | Modelo del producto | |
| P/N | String | Número de parte | |
| S/N | String | Número de serie | |
| OBSERVACIONES | Text | Notas adicionales | |
| STATUS | String | Estado (Activo/Inactivo) | |
| UBICACIÓN | String | Ubicación física | |
| MEDIDA | String | Unidad de medida | |
| CANT. | Numérico | Cantidad disponible | ✅ |
| PRECIO UNIT | Numérico | Precio unitario | ✅ |
| TOTAL | Numérico | Calculado automáticamente | Auto |

### Nuevas Columnas Profesionales
| Columna | Tipo | Descripción | Default |
|---------|------|-------------|---------|
| CATEGORÍA | String | Categoría del producto | "Sin categoría" |
| STOCK MÍNIMO | Numérico | Punto de reorden | 0 |
| FECHA_CREACIÓN | DateTime | Timestamp de creación | Auto |
| FECHA_MODIFICACIÓN | DateTime | Timestamp de modificación | Auto |

## 🎨 Interfaz de Usuario

### Organización por Tabs

#### 1. 📊 Dashboard
- KPIs principales del inventario
- Alertas de stock bajo destacadas
- Gráficos de valor por categoría
- Distribución ABC

#### 2. 📦 Inventario Completo
- Visualización completa del inventario
- Ordenamiento dinámico por cualquier columna
- Opción de mostrar clasificación ABC
- Descarga del inventario actualizado

#### 3. 🔍 Búsqueda Avanzada
- Búsqueda global por texto
- Filtros combinados (categoría, ubicación, status, rangos)
- Estadísticas de resultados
- Exportación de resultados filtrados

#### 4. 📈 Reportes & Analíticas
- Análisis ABC detallado con gráficos
- Top 10 items por valor
- Distribución por ubicación
- Métricas por clase ABC

#### 5. ➕ Agregar Item
- Formulario con validaciones en tiempo real
- Ayuda contextual en cada campo
- Cálculo automático de totales
- Alertas inmediatas de stock bajo

## 🛠️ Tecnologías Utilizadas

- **Python 3.11**
- **Streamlit:** Framework web interactivo
- **Pandas:** Análisis y manipulación de datos
- **OpenPyXL:** Lectura/escritura de archivos Excel
- **Plotly:** Visualizaciones interactivas

## 📦 Instalación

```bash
pip install -r requirements.txt
```

## ▶️ Ejecución

```bash
streamlit run inventario_online_v3.py
```

O con configuración específica:

```bash
streamlit run inventario_online_v3.py --server.enableCORS false --server.enableXsrfProtection false
```

## 💡 Uso

1. **Cargar inventario:** Usa el panel lateral para subir un archivo CSV o XLSX
2. **Explorar dashboard:** Visualiza KPIs y alertas en el tab de Dashboard
3. **Buscar y filtrar:** Usa filtros avanzados para encontrar items específicos
4. **Agregar items:** Completa el formulario con validaciones automáticas
5. **Analizar:** Revisa reportes ABC y visualizaciones
6. **Exportar:** Descarga el inventario actualizado o resultados filtrados

## 🔄 Compatibilidad con Archivos Existentes

El sistema es **100% compatible** con archivos de inventario existentes:
- Las nuevas columnas se agregan automáticamente si no existen
- Los archivos antiguos funcionan sin modificaciones
- Se preserva el formato original al exportar
- HEADER_OFFSET configurable para archivos con metadatos

## ⚙️ Configuración

### Constantes Configurables (inventario_online_v3.py)

```python
HEADER_OFFSET = 3  # Filas a saltar antes de la cabecera

COLUMNAS_REQUERIDAS = ['ITEM', 'DESCRIPCIÓN', 'CANT.', 'PRECIO UNIT']

COLUMNAS_EXTENDIDAS = ['CATEGORÍA', 'STOCK MÍNIMO', 'FECHA_CREACIÓN', 'FECHA_MODIFICACIÓN']
```

### Tema Visual (.streamlit/config.toml)

```toml
[theme]
backgroundColor="#000000"
secondaryBackgroundColor="#054fde"
textColor="#f5f5ff"
```

## 📝 Validaciones Implementadas

### Al Agregar Items
1. ✅ ITEM debe ser único (no duplicados)
2. ✅ ITEM y DESCRIPCIÓN son obligatorios
3. ✅ CANT. no puede ser negativa
4. ✅ PRECIO UNIT no puede ser negativo
5. ✅ STOCK MÍNIMO no puede ser negativo
6. ✅ TOTAL se calcula automáticamente
7. ✅ Timestamps se generan automáticamente
8. ✅ Alerta si stock ≤ stock mínimo

## 📊 Clasificación ABC - Fundamentos

La clasificación ABC es una técnica de gestión de inventarios basada en el Principio de Pareto:

- **Regla 80/20:** El 80% del valor proviene del 20% de los items
- **Priorización:** Permite enfocar recursos en items de alto valor
- **Optimización:** Mejora la eficiencia en control de inventarios

### Aplicación
- **Clase A:** Control estricto, revisión frecuente, stock de seguridad bajo
- **Clase B:** Control moderado, revisión mensual
- **Clase C:** Control simple, revisión trimestral

## 🆕 Nuevas Funcionalidades vs Versión Anterior

| Característica | Versión Anterior | Versión Optimizada |
|----------------|------------------|-------------------|
| Validación de unicidad | ❌ | ✅ |
| Cálculo automático de TOTAL | ❌ | ✅ |
| Alertas de stock bajo | ❌ | ✅ |
| Clasificación ABC | ❌ | ✅ |
| Dashboard con KPIs | ❌ | ✅ |
| Gráficos interactivos | ❌ | ✅ |
| Filtros avanzados | ❌ | ✅ |
| Categorización | ❌ | ✅ |
| Auditoría con timestamps | ❌ | ✅ |
| Exportación de filtrados | ❌ | ✅ |
| Interfaz con tabs | ❌ | ✅ |
| Validaciones de negocio | Básicas | Completas |

## 🤝 Contribuciones

Sistema diseñado y optimizado para uso profesional en gestión de inventarios.

## 📄 Licencia

Proyecto de uso interno para gestión de inventarios.
