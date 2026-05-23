# TechTeam · Gestión de desarrollo técnico de equipos

Aplicación en **Streamlit** para comparar evaluaciones técnicas de un equipo, construir una **fotografía global del grupo** y ayudar a un director técnico o director de especie a **equilibrar capacidades, detectar gaps y definir un plan de desarrollo colectivo**.

Esta app **no está orientada a elegir un senior**.  
Su finalidad es otra:

- comparar técnicos sin jerarquizarlos como “mejor” o “peor”;
- identificar **fortalezas y debilidades del equipo**;
- ver qué áreas están suficientemente cubiertas y cuáles no;
- detectar personas que pueden **cubrir gaps del equipo** o desarrollarse para hacerlo;
- proponer un **plan de desarrollo técnico del equipo**.

---

## Qué hace la app

La aplicación permite:

- cargar entre **2 y 15 assessments** individuales;
- filtrar por:
  - **bloque productivo**:
    - Monogástricos
    - Rumiantes
  - **especie / subespecie**
- comparar técnicos dentro de un mismo grupo;
- leer los resultados de assessment por los **4 troncos**:
  - Alimentación
  - Sanidad
  - Manejo
  - Herramientas
- visualizar:
  - distribución de niveles del equipo;
  - cobertura técnica por troncos;
  - fortalezas y gaps del grupo;
  - personas que ya pueden actuar como referencia por área;
  - personas con potencial de desarrollo para cubrir huecos;
- construir un **plan de desarrollo del equipo**;
- generar informes en:
  - **HTML corporativo**
  - **PDF corporativo**

---

## Enfoque de la herramienta

La app está pensada para responder preguntas como estas:

- ¿Cómo está realmente mi equipo técnico hoy?
- ¿En qué troncos estamos fuertes como equipo?
- ¿Dónde tenemos gaps de cobertura o dependencia excesiva?
- ¿Qué técnicos ya pueden actuar como referencia en un área?
- ¿A quién conviene desarrollar para reforzar el equilibrio del equipo?
- ¿Qué plan de desarrollo debería priorizar el director técnico?

No busca clasificar a una persona como ganadora, sino ayudar a construir un equipo más robusto, homogéneo y complementario.

---

## Acceso restringido

La app solicita contraseña cada vez que se inicia una nueva sesión del navegador.

**Contraseña actual:**

```text
TechTeam2026+
```

> Es una barrera básica de acceso.  
> Si más adelante quieres más seguridad, conviene moverla a `st.secrets`.

---

## Identidad corporativa

La app utiliza imagen corporativa con:

- **Logo Nutreco**
- **Logo TechTeam**
- **Franja rosa corporativa**

La misma identidad se aplica también a los informes descargables.

---

## Entradas necesarias

### 1. Assessments individuales
La app acepta archivos:

- `.xlsm`
- `.xlsx`

Cada archivo debe corresponder a una evaluación individual completa.

### 2. Número de archivos
La herramienta trabaja con un mínimo de:

- **2 archivos**

y un máximo de:

- **15 archivos**

### 3. Cohorte comparada
Los assessments pueden ser:

- de una sola especie / subespecie,
- o mezclados dentro del mismo gran bloque productivo:

#### Monogástricos
- Avicultura
- Porcino
- Conejos
- Caballos

#### Rumiantes
- Vacuno leche
- Vacuno carne
- Ovino y caprino

---

## Troncos evaluados

La app trabaja con estos 4 troncos del assessment:

- **Alimentación**
- **Sanidad**
- **Manejo**
- **Herramientas**

Estos cuatro bloques son la base de la fotografía del equipo y del plan de desarrollo.

---

## Qué muestra la app

## 1. Fotografía general del equipo
Resume la situación del grupo técnico y ayuda a interpretar de un vistazo:

- equilibrio global;
- dispersión de niveles;
- fortalezas colectivas;
- debilidades colectivas.

## 2. Distribución de niveles
Muestra cómo se reparte el equipo en términos de nivel técnico.

## 3. Cobertura por troncos
Permite ver qué áreas están mejor cubiertas y cuáles están menos reforzadas.

## 4. Heatmap comparativo
Incluye un **heatmap del equipo vs objetivo** para visualizar diferencias entre técnicos y troncos.

## 5. Radar comparativo entre técnicos
Incluye un **gráfico de radar** para comparar perfiles técnicos de varios miembros del equipo.

## 6. Referencias técnicas por área
La app identifica qué personas ya pueden actuar como apoyo o referencia relativa en un tronco.

## 7. Personas a desarrollar
También señala qué personas pueden desarrollarse para reforzar áreas del equipo que hoy están infrarepresentadas.

## 8. Plan de desarrollo del equipo
La herramienta genera una propuesta directiva de desarrollo colectivo:
- qué áreas reforzar;
- qué personas potenciar;
- qué equilibrio perseguir como equipo.

---

## Lógica de interpretación del equipo

La app clasifica el estado de cada tronco del equipo en categorías como:

- **Fortaleza consolidada**
- **Aceptable con riesgo**
- **Gap moderado**
- **Gap crítico**

La lectura se basa en dos ideas:
- nivel medio del equipo en ese tronco;
- número de personas que ya pueden actuar como referencia relativa.

---

## Informes

La app genera informes corporativos en:

- **HTML**
- **PDF**

### Qué incluyen
- cabecera corporativa;
- resumen del equipo;
- diagnóstico de fortalezas y gaps;
- cobertura por troncos;
- heatmap;
- radar comparativo;
- propuesta de desarrollo del equipo.

---

## Archivos necesarios en el repositorio

Debes dejar en la raíz del repositorio, como mínimo:

- `main.py`
- `requirements.txt`
- `README.md`

Y además, para mantener la identidad corporativa:

- `Logo Nutreco.jpg`
- `Logo TechTeam 2.jpg`
- `Solapa rosa.jpg`

---

## Estructura recomendada del repositorio

```text
/
├── main.py
├── requirements.txt
├── README.md
├── Logo Nutreco.jpg
├── Logo TechTeam 2.jpg
├── Solapa rosa.jpg
├── CTC 3.xlsm
├── CTC 4.xlsm
└── ...
```

---

## Despliegue en Streamlit Community Cloud

### Paso 1. Subir a GitHub
Sube al repositorio:

- `main.py`
- `requirements.txt`
- `README.md`
- logos corporativos
- assessments de ejemplo opcionales

### Paso 2. Crear la app
En Streamlit Community Cloud:

1. conecta tu cuenta de GitHub;
2. pulsa **Create app**;
3. selecciona:
   - repositorio
   - rama
   - archivo principal: `main.py`
4. pulsa **Deploy**.

---

## Dependencias

El archivo debe llamarse exactamente:

```text
requirements.txt
```

Contenido recomendado para esta versión:

```txt
streamlit>=1.44.0
pandas>=2.2.0
openpyxl>=3.1.0
plotly>=5.20.0
reportlab>=4.0.0
kaleido>=0.2.1
```

---

## Cómo usar la app

1. Abre la app.
2. Introduce la contraseña.
3. Selecciona el **bloque productivo**:
   - Monogástricos
   - Rumiantes
4. Selecciona la **especie / subespecie**.
5. Sube entre **2 y 15 assessments**.
6. Revisa:
   - la fotografía general del equipo;
   - la distribución de niveles;
   - la cobertura por troncos;
   - el heatmap;
   - el radar comparativo;
   - las áreas a reforzar;
   - y las personas a desarrollar.
7. Descarga el informe corporativo en HTML o PDF.

---

## Qué debe comprobar el usuario

Conviene revisar especialmente:

- que todos los archivos cargados pertenezcan a una cohorte comparable;
- que el grupo esté bien filtrado por bloque productivo y especie;
- que los Excel hayan sido recalculados y guardados correctamente antes de subirlos;
- que el plan propuesto tenga sentido desde el punto de vista organizativo, no solo técnico.

---

## Limitaciones conocidas

- La contraseña está integrada en el código.
- La app depende de una estructura de Excel compatible con la plantilla esperada.
- Si los archivos no han sido recalculados correctamente antes de guardarse, algunos resultados pueden no reflejar el estado real del libro.
- La interpretación del equipo ayuda a tomar decisiones, pero no sustituye el criterio del director técnico.
- La calidad de la fotografía del equipo depende de que la cohorte comparada sea coherente.

---

## Recomendaciones

- Usa grupos homogéneos dentro del mismo gran bloque productivo.
- No mezcles monogástricos y rumiantes en un mismo análisis.
- Revisa no solo el nivel medio, sino también la **dependencia de pocas personas** en un tronco.
- Utiliza la app como herramienta de planificación del desarrollo técnico del equipo, no solo como reporte descriptivo.

---

## Objetivo de negocio

Esta app ayuda a:

- visualizar el estado técnico real del equipo;
- equilibrar capacidades;
- reducir gaps;
- reforzar áreas críticas;
- identificar referencias técnicas internas;
- y construir un plan de desarrollo de equipo más sólido, explicable y accionable.

