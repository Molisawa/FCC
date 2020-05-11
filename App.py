import xlrd
from tabulate import tabulate

filePath = "BDD.xlsx"
openFile = xlrd.open_workbook(filePath)
horarios = openFile.sheet_by_name("horarios")
materias = openFile.sheet_by_name("materias")
profesores = openFile.sheet_by_name("profesores")
secciones = openFile.sheet_by_name("secciones")

entrada = input("Ingrese la opcion: ")


def opcion1():
    todo = []
    for i in range(1, materias.nrows):
        todo.append([str(materias.cell_value(i, 0)), materias.cell_value(i, 1)])
    print(tabulate(todo, headers=["Id_materia", "nombre"], tablefmt="orgtbl"))


def opcion2():
    query = str(input("Ingrese el ID de la materia: "))
    materia = ""
    index = 0
    for i in materias.col_values(0, 1):
        index += 1
        if str(i) == query:
            materia = materias.cell_value(index, 1)
    seccionesProfesoresID = []
    seccionesHorariosID = []
    index = 0
    for i in secciones.col_values(2, 1):
        index += 1
        if str(i) == query:
            seccionesProfesoresID.append(secciones.cell_value(index, 1))
            seccionesHorariosID.append(secciones.cell_value(index, 3))
    index = 0
    seccionesProfesores = []
    for i in profesores.col_values(0, 1):
        index += 1
        for j in seccionesProfesoresID:
            if j == str(i):
                seccionesProfesores.append(profesores.cell_value(index, 1))
    seccionesHorarios = []
    index = 0
    for i in horarios.col_values(0, 1):
        index += 1
        for j in seccionesHorariosID:
            if j == str(i):
                seccionesHorarios.append(horarios.cell_value(index, 1))
    todo = []
    for i in range(len(seccionesHorariosID)):
        todo.append([materia, seccionesProfesores[i], seccionesHorarios[i]])
    print(tabulate(todo, headers=["materia", "profesor", "horario"], tablefmt="orgtbl"))

def opcion3():
    todo = []
    for i in range(1, profesores.nrows):
        todo.append([str(profesores.cell_value(i, 0)), profesores.cell_value(i, 1)])
    print(tabulate(todo, headers=["Id_profesor", "nombre"], tablefmt="orgtbl"))


def opcion4():
    query = str(input("Ingrese el ID del profesor: "))
    seccionesMateriasID = []
    seccionesHorariosID = []
    index = 0
    for i in secciones.col_values(1, 1):
        index += 1
        if str(i) == query:
            seccionesMateriasID.append(secciones.cell_value(index, 2))
            seccionesHorariosID.append(secciones.cell_value(index, 3))
    index = 0
    seccionesMaterias = []
    for i in materias.col_values(0, 1):
        index += 1
        for j in seccionesMateriasID:
            if j == str(i):
                seccionesMaterias.append(materias.cell_value(index, 1))
    seccionesHorarios = []
    index = 0
    for i in horarios.col_values(0, 1):
        index += 1
        for j in seccionesHorariosID:
            if j == str(i):
                seccionesHorarios.append(horarios.cell_value(index, 1))
    todo = []
    for i in range(len(seccionesHorariosID)):
        todo.append([seccionesMaterias[i], seccionesHorarios[i]])
    print(tabulate(todo, headers=["materia", "horario"], tablefmt="orgtbl"))


def noEncontrado():
    print("No existe la opcion", entrada)


valores = {
    "1": opcion1,
    "2": opcion2,
    "3": opcion3,
    "4": opcion4
}

valores.get(entrada, noEncontrado)()
