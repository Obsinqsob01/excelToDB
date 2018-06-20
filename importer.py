import pandas as pd
import pymysql
import numpy
import pymysql.cursors

#database
connextion = pymysql.connect(host='cbtis148.edu.mx',
                                user='cbtis182_lfcg',
                                password='obsinqsob01.',
                                db='cbtis182_cbtis',
                                port=3306,
                                cursorclass=pymysql.cursors.DictCursor)

#Work Excel
file = 'HorasClub.xlsm'

xl = pd.ExcelFile(file)

# print(xl.sheet_names)

df1 = xl.parse('Base')

# print(df1['No. Control'])

noControl = df1['No. Control']
nombre = df1['Nombre']
club1 = df1['Club 1']
horas1 = df1['Horas']
club2 = df1['Club 2']
horas2 = df1['Horas2']
club3 = df1['Club 3']
horas3 = df1['Horas3']
arbol = df1['Arbol']
extra = df1['Extra']
faltasTutoria = df1['Ftuto17']
faltasEF = df1['Fedu17']
observaciones = df1['Observaciones']

def changeNan(val):
    # print(type(val))
    if str(val) == 'nan':
        if type(val) == str:
            # print("es str")
            return ' ' 
        if type(val) == numpy.float64:
            # print("es numpyu")
            return 0
        
        if type(val) == int:
            # print("es int")
            return 0

        if type(val) == float:
            return 0

    return val;

for i in range(562):
    try:
        with connextion.cursor() as cursor:

            p1 = noControl[i]
            p1 = changeNan(p1)

            p2 = nombre[i]
            p2 = changeNan(p2)

            p3 = club1[i]
            p3 = changeNan(p3)

            p4 = horas1[i]
            p4 = changeNan(p4)

            p5 = club2[i]
            p5 = changeNan(p5)

            p6 = horas2[i] 
            p6 = changeNan(p6)

            p7 = club3[i]
            p7 = changeNan(p7)

            p8 = horas3[i]
            p8 = changeNan(p8)
            
            p9 = arbol[i]
            p9 = changeNan(p9)

            p10 = extra[i]
            p10 = changeNan(p10)

            p11 = faltasTutoria[i]
            p11 = changeNan(p11)

            p12 = faltasEF[i]
            p12 = changeNan(p12)

            p13 = observaciones[i]
            p13 = changeNan(p13)

            sql = """INSERT INTO alumnos(noControl, nombre, club1, horas1, club2, horas2,
                    club3, horas3, club4, horas4, extra1, extra2, faltasTutoria, faltasEF,
                    observaciones) VALUES ({0}, '{1}', '{2}', {3}, '{4}', {5}, '{6}', {7}, '', 0,
                    '{8}', {9}, {10}, {11}, '{12}')""".format(p1, p2, p3,p4, p5, p6, p7, p8, p9, p10, p11, p12, p13)
            
            print(sql)

            cursor.execute(sql)

            connextion.commit()
    finally:
        print("Exito!")

connextion.close()