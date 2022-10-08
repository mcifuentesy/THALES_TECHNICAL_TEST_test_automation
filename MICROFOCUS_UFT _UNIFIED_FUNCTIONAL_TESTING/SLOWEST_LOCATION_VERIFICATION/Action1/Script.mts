Browser("Browser").Page("Google Maps").WebEdit("Buscar en Google Maps").Set "thales colombia"
Browser("Browser").Page("Google Maps").WebButton("Buscar").Click
Browser("Browser").Page("Thales International Suc.").WebButton("Dirección: Cra. 12 #93-8,").Click
Browser("Browser").Page("Thales International Suc.").Link("Sitio web: thalesgroup.com").Click
Browser("Browser").Page("Thales International Suc._2").WebButton("Teléfono: 17442442").Click
Browser("Browser").Page("Thales International Suc.").Image("Cómo llegar").Click
Browser("Browser").Page("Google Maps_2").WebEdit("Punto de partida Cota").Set DataTable("Location", dtGlobalSheet)
Browser("Browser").Page("Google Maps_2").WebButton("Buscar").Click
DataTable.SetCurrentRow(2)
DataTable("Time Arriving 1st Option_min", dtGlobalSheet) = Browser("Browser").Page("de Cota, Cundinamarca_3").WebElement("37 min").GetROProperty("innerText")
DataTable("Time Arriving 1st Option_min", dtGlobalSheet) = Browser("Browser").Page("de Cota, Cundinamarca_3").WebElement("37 min").GetROProperty("innerText")
DataTable("Route_km", dtGlobalSheet) = Browser("Browser").Page("de Cota, Cundinamarca_3").WebElement("21,3 km").GetROProperty("innerText")
Browser("Browser").Close

DataTable.SetCurrentRow(1)
initial_time_data = DataTable("Time Arriving 1st Option_min", dtGlobalSheet)
initial_route_distance  = DataTable("Route_km", dtGlobalSheet)
Location_to_verify = DataTable("Location", dtGlobalSheet)

DataTable.SetCurrentRow(2)

verification_time_data = DataTable("Time Arriving 1st Option_min", dtGlobalSheet)
verification_route_distance  = DataTable("Route_km", dtGlobalSheet)

DataTable("Location", dtGlobalSheet) = Location_to_verify + " Verification"

DataTable.SetCurrentRow(3)

DataTable("Location", dtGlobalSheet) = "Verification"

If initial_time_data = verification_time_data Then
	DataTable("Time Arriving 1st Option_min", dtGlobalSheet) = "Not changed"
	else
	DataTable("Time Arriving 1st Option_min", dtGlobalSheet) = "Changed"
End If

If initial_route_distance = verification_route_distance Then
	DataTable("Route_km", dtGlobalSheet) = "Not changed"
	else
	DataTable("Route_km", dtGlobalSheet) = "Changed"
End If

DataTable.Export ("C:\Exported_Data_Locations_Cundinamarca_Thales\times_arrivings_and_routes_verification_slowest.xls")
 @@ hightlight id_;_67102_;_script infofile_;_ZIP::ssf51.xml_;_
ExitTest
 @@ script infofile_;_ZIP::ssf53.xml_;_

 @@ script infofile_;_ZIP::ssf66.xml_;_
 @@ script infofile_;_ZIP::ssf57.xml_;_
 @@ script infofile_;_ZIP::ssf60.xml_;_
