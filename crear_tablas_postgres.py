import psycopg2

conn = psycopg2.connect("postgresql://tesch:LdYWnIgIMZ5E5ngP1HASUC4LzhyYcqut@dpg-d90nhr0js32c73dcc72g-a.ohio-postgres.render.com/tesch")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS usuarios (
    usuario TEXT PRIMARY KEY,
    password TEXT,
    rol TEXT,
    cambio_password INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS alumnos (
    matricula TEXT PRIMARY KEY,
    nombre TEXT,
    carrera TEXT
);

CREATE TABLE IF NOT EXISTS actividades (
    id INTEGER PRIMARY KEY,
    nombre TEXT,
    docente TEXT
);

CREATE TABLE IF NOT EXISTS periodos (
    id INTEGER PRIMARY KEY,
    periodo TEXT,
    activo INTEGER
);

CREATE TABLE IF NOT EXISTS inscripciones (
    id INTEGER PRIMARY KEY,
    matricula TEXT,
    semestre TEXT,
    genero TEXT,
    telefono TEXT,
    actividad TEXT,
    periodo TEXT
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_inscripcion
ON inscripciones(matricula, periodo);

CREATE TABLE IF NOT EXISTS control_fechas (
    id INTEGER PRIMARY KEY,
    fecha_inicio DATE,
    fecha_fin DATE
);

CREATE TABLE IF NOT EXISTS eventos (
    id INTEGER PRIMARY KEY,
    docente TEXT,
    actividad TEXT,
    nombre_evento TEXT,
    institucion TEXT,
    fecha DATE,
    participantes INTEGER,
    mujeres INTEGER,
    hombres INTEGER,
    resultados TEXT,
    periodo TEXT
);

CREATE TABLE IF NOT EXISTS resultados (
    id INTEGER PRIMARY KEY,
    matricula TEXT,
    actividad TEXT,
    periodo TEXT,
    resultado TEXT
);

CREATE TABLE IF NOT EXISTS evaluaciones (
    id INTEGER PRIMARY KEY,
    matricula TEXT,
    actividad TEXT,
    periodo TEXT,
    c1 INTEGER,
    c2 INTEGER,
    c3 INTEGER,
    c4 INTEGER,
    c5 INTEGER,
    c6 INTEGER,
    c7 INTEGER,
    promedio REAL,
    nivel TEXT,
    observaciones TEXT
);
""")

conn.commit()
conn.close()

print("✅ TABLAS CREADAS CORRECTAMENTE")
