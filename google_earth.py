import xlrd
import simplekml

book = xlrd.open_workbook("input.xls")
sh = book.sheet_by_index(0)
row = int(sh.nrows)

nom_liaison = []
siteB = []
siteA = []
latitudeB = []
longitudeB = []
latitudeA = []
longitudeA = []
azimutA = []
azimutB = []
antenneA = []
antenneB = []
hmasiteA = []
        
for i in range(0, row):
        nom_liaison.append(sh.cell_value(rowx=i, colx=0))
        siteB.append(sh.cell_value(rowx=i, colx=1))
        siteA.append(sh.cell_value(rowx=i, colx=6))        
        latitudeB.append(sh.cell_value(rowx=i, colx=2))       
        longitudeB.append(sh.cell_value(rowx=i, colx=3))        
        latitudeA.append(sh.cell_value(rowx=i, colx=7))       
        longitudeA.append(sh.cell_value(rowx=i, colx=8))
        azimutA.append(sh.cell_value(rowx=i, colx=10))
        azimutB.append(sh.cell_value(rowx=i, colx=4))
        antenneA.append(sh.cell_value(rowx=i, colx=11))
        antenneB.append(sh.cell_value(rowx=i, colx=5))
        hmasiteA.append(sh.cell_value(rowx=i, colx=9))

for x in range(1, row):

        #SiteB
        monkml = simplekml.Kml()
        project_name = (nom_liaison[x])
        monkml.document.name = project_name
        descriptionB = "Azimuth Site B : " + str(azimutB[x]) + "<br>Type d'antenne B : " + str(antenneB[x])
        point1 = monkml.newpoint(name=siteB[x], coords=[(longitudeB[x],latitudeB[x])])
        point1.description = (str(descriptionB))
        point1.style.iconstyle.icon.href = 'https://cdn-icons-png.flaticon.com/512/9020/9020807.png'
        
        #SiteA
        descriptionA = "Azimuth Site A : " + str(azimutA[x]) + "<br>Type d'antenne A : " + str(antenneA[x]) + "<br>HMA antenne : " + str(hmasiteA[x])
        point2 = monkml.newpoint(name=siteA[x], coords=[(longitudeA[x],latitudeA[x])])
        point2.description = (str(descriptionA))
        point2.style.iconstyle.icon.href = 'https://cdn-icons-png.flaticon.com/512/9020/9020807.png'
        
        #Ligne
        line = monkml.newlinestring(name=nom_liaison[x], coords=[(longitudeB[x], latitudeB[x]), (longitudeA[x], latitudeA[x])])
        line.linestyle.color = simplekml.Color.red
        name = str(x) + ".kml"
        monkml.save(name)
