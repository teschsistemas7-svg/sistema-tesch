import psycopg2
import csv

conn = psycopg2.connect("postgresql://tesch:LdYWnIgIMZ5E5ngP1HASUC4LzhyYcqut@dpg-d90nhr0js32c73dcc72g-a.ohio-postgres.render.com/tesch")
cursor = conn.cursor()

# ✅ USUARIOS
with open('usuarios.csv', encoding='utf-8') as f:
    reader = csv.reader(f)
    headers = next(reader)

    for fila in reader:
        try:
            # rowid, usuario, password, rol, cambio_password
            usuario = fila[1]
            password = fila[2]
            rol = fila[3]
            cambio = fila[4]

            cursor.execute("""
                INSERT INTO usuarios (usuario, password, rol, cambio_password)
                VALUES (%s, %s, %s, %s)
            """, (usuario, password, rol, cambio))

        except Exception as e:
            print("Error usuarios:", e)
            conn.rollback()

# ✅ ALUMNOS
with open('alumnos.csv', encoding='utf-8') as f:
    reader = csv.reader(f)
    next(reader)

    for fila in reader:
        try:
            matricula = fila[1]
            nombre = fila[2]
            carrera = fila[3]

            cursor.execute("""
                INSERT INTO alumnos (matricula, nombre, carrera)
                VALUES (%s, %s, %s)
            """, (matricula, nombre, carrera))

        except:
            conn.rollback()


# ✅ ACTIVIDADES
with open('actividades.csv', encoding='utf-8') as f:
    reader = csv.reader(f)
    next(reader)

    for fila in reader:
        try:
            id = fila[1]
            nombre = fila[2]
            docente = fila[3]

            cursor.execute("""
                INSERT INTO actividades (id, nombre, docente)
                VALUES (%s, %s, %s)
            """, (id, nombre, docente))

        except:
            conn.rollback()

conn.commit()
conn.close()

print("✅ MIGRACIÓN CORRECTA")