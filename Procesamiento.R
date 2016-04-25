require(openxlsx)
require(data.table)
rm(list=ls())
#read the data
# dx=data.table(read.xlsx("/Users/aldopareja/Google Drive/CCB/BD/ReporteDiagnosticos_25-04-2016-112831.xlsx"))
# save(dx,file="backUps/dx15Abril2016.RData")
load("backUps/dx15Abril2016.RData")
CIIU=data.table(read.xlsx("CIIU.xlsx"))
dx=dx[CIIU.DX%in%CIIU[,CIIU]]
dx[Estado.DX=="Terminado",Estado.DX:="Diligenciado"]
dx[is.na(Usuario.Asignado)|grepl("ESPITIA",Usuario.Asignado),
   Clasificación.DX:="POTENCIAL."]
dx[!(is.na(Usuario.Asignado)|grepl("ESPITIA",Usuario.Asignado)),
   Clasificación.DX:="ALTO POTENCIAL."]
dx[,ID.empresa:=NULL]
save(dx,file="backUps/dx.RData")
#create workbook to write data
wb=createWorkbook()
addWorksheet(wb,"detalleClientes")
writeDataTable(wb,"detalleClientes",dx)
#now summary by type
dx=dx[,list(numeroClientes=.N),by='Nombre.DX,Clasificación.DX,Estado.DX']
addWorksheet(wb,"consolidadoPorTypo")
writeDataTable(wb,"consolidadoPorTypo",dx)
#now summary by CIIU
load("backUps/dx.RData")
dx=dx[CIIU.DX%in%CIIU[,CIIU]]
dx[Estado.DX=="Terminado",Estado.DX:="Diligenciado"]
dx=dx[,list(numeroClientes=.N),by='CIIU.DX,Nombre.DX,Clasificación.DX,Estado.DX']
addWorksheet(wb,"consolidadoPorCIIU")
writeDataTable(wb,"consolidadoPorCIIU",dx)
#now summary by Advisor
load("backUps/dx.RData")
dx=dx[CIIU.DX%in%CIIU[,CIIU]]
dx[Estado.DX=="Terminado",Estado.DX:="Diligenciado"]
dx=dx[,list(numeroClientes=.N),by='Usuario.Asignado,Nombre.DX,Clasificación.DX,Estado.DX']
dx=dx[order(Usuario.Asignado,Clasificación.DX,Nombre.DX,Estado.DX)]
addWorksheet(wb,"consolidadoPorAsesor")
writeDataTable(wb,"consolidadoPorAsesor",dx)

#now summary by Advisor and CIIU
load("backUps/dx.RData")

dx=dx[CIIU.DX%in%CIIU[,CIIU]]
dx[Estado.DX=="Terminado",Estado.DX:="Diligenciado"]
dx=dx[,list(numeroClientes=.N),by='Usuario.Asignado,CIIU.DX,Nombre.DX,Clasificación.DX,Estado.DX']
dx=dx[order(Usuario.Asignado,CIIU.DX,Clasificación.DX,Nombre.DX,Estado.DX)]
addWorksheet(wb,"consolidadoPorAsesorYCIIU")
writeDataTable(wb,"consolidadoPorAsesorYCIIU",dx)

#write file
saveWorkbook(wb,"reportes/reporte.xlsx",overwrite = T)
