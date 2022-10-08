Browser("Browser").Page("Google Maps").WebEdit("Buscar en Google Maps").Set "thales colombia" @@ script infofile_;_ZIP::ssf39.xml_;_
Browser("Browser").Page("Google Maps").WebButton("Buscar").Click
Browser("Browser").Page("Thales International Suc.").WebButton("Dirección: Cra. 12 #93-8,").Check CheckPoint("Dirección: Cra. 12 #93-8,") @@ script infofile_;_ZIP::ssf71.xml_;_
Browser("Browser").Page("Thales International Suc.").WebButton("Dirección: Cra. 12 #93-8,").Click @@ script infofile_;_ZIP::ssf40.xml_;_
Browser("Browser").Page("Thales International Suc.").Link("Sitio web: thalesgroup.com").Check CheckPoint("Sitio web: thalesgroup.com_5") @@ script infofile_;_ZIP::ssf69.xml_;_
Browser("Browser").Page("Thales International Suc.").Link("Sitio web: thalesgroup.com").Click @@ hightlight id_;_9701966_;_script infofile_;_ZIP::ssf42.xml_;_
Browser("Browser").Page("Thales International Suc._2").WebButton("Teléfono: 17442442").Check CheckPoint("Teléfono: 17442442") @@ script infofile_;_ZIP::ssf70.xml_;_
Browser("Browser").Page("Thales International Suc._2").WebButton("Teléfono: 17442442").Click
Browser("Browser").Page("Thales International Suc.").Image("Cómo llegar").Click @@ script infofile_;_ZIP::ssf35.xml_;_
Browser("Browser").Page("Google Maps_2").WebEdit("Punto de partida Cota").Set DataTable("Location", dtGlobalSheet)
Browser("Browser").Page("Google Maps_2").WebButton("Buscar").Click @@ script infofile_;_ZIP::ssf37.xml_;_
DataTable("Time Arriving 1st Option_min", dtGlobalSheet) = Browser("Browser").Page("de Cota, Cundinamarca_3").WebElement("37 min").GetROProperty("innerText")
DataTable("Route_km", dtGlobalSheet) = Browser("Browser").Page("de Cota, Cundinamarca_3").WebElement("21,3 km").GetROProperty("innerText")
Browser("Browser").Close


DataTable.Export ("C:\Exported_Data_Locations_Cundinamarca_Thales\times_arrivings_and_routes.xls") @@ script infofile_;_ZIP::ssf57.xml_;_
 @@ script infofile_;_ZIP::ssf60.xml_;_
