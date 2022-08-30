Set oStream = CreateObject("ADODB.Stream")
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
Set oFSO = CreateObject("Scripting.FileSystemObject")
	oXMLHTTP.Open "GET", "https://piston-data.mojang.com/v1/objects/43db9b498cb67058d2e12d394e6507722e71bb45/client.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/minecraft.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/jinput.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jopt-simple-4.5.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/jopt-simple-4.5.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl_test.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_test.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl_util.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_util.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl_util_applet.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_util_applet.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lzma.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lzma.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl-debug.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl-debug.jar"
    oStream.Close
	'download natives
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-dx8.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-dx8.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-dx8_64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-dx8_64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-raw.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-raw.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-raw_64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-raw_64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/lwjgl.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/lwjgl64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/OpenAL32.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/OpenAL32.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/OpenAL64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/OpenAL64.dll"
    oStream.Close