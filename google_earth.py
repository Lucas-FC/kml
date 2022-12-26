import xlrd
import simplekml

# initiation 
monkml = simplekml.Kml()

book = xlrd.open_workbook("24_11_22.xls")
sh = book.sheet_by_index(0)
row = sh.nrows

print(row)
        
for i in range(0, row):
        nom_liaison = []
        siteB = []
        siteA = []
        latitudeB = []
        longitudeB = []
        latitudeA = []
        longitudeA = []
        nom_liaison.append(sh.cell_value(rowx=i, colx=0))
        siteB.append(sh.cell_value(rowx=i, colx=1))
        siteA.append(sh.cell_value(rowx=i, colx=6))        
        latitudeB.append(sh.cell_value(rowx=i, colx=2))       
        longitudeB.append(sh.cell_value(rowx=i, colx=3))        
        latitudeA.append(sh.cell_value(rowx=i, colx=7))       
        longitudeA.append(sh.cell_value(rowx=i, colx=8))
        project_name = (siteA[i])
        monkml.document.name = project_name
        monkml.newpoint(name=siteB[i], coords=[(longitudeB[i],latitudeB[i])])
        monkml.newpoint(name=siteA[i], coords=[(longitudeA[i],latitudeA[i])])
        name = str(project_name) + ".kml"
        monkml.save(name)

      

