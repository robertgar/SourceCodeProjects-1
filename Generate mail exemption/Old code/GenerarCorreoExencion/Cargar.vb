Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System
Imports System.IO
Imports System.Data.SqlClient
Imports Newtonsoft.Json.Linq
Imports System.Net


Public Class Cargar
    'funcion que retorna vacio cuando no encuentra ningun valor en la tabla
    Function retornavacio(ByVal consulta As String, ByVal mystring As String) As Integer
        Dim transaccion As Integer
        transaccion = 0

        Using mySqlConnection As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection.Open()

            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection)
            Dim myDataReader As Data.SqlClient.SqlDataReader

            myDataReader = mySqlCommand.ExecuteReader()

            Do While (myDataReader.Read())
                transaccion = 1
            Loop

            myDataReader.Close()
            mySqlConnection.Close()

        End Using

        Return transaccion

    End Function
    'funcion que retorna un valor decimal
    Function retornardecimal(ByVal consulta As String, ByVal mystring As String) As Decimal

        Dim VarDecimal As Decimal

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            VarDecimal = myDataReader1.GetDecimal(0)
            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return VarDecimal

    End Function

    'funcion que retorna una cadena
    Function retornarcadena(ByVal consulta As String, ByVal mystring As String) As String

        Dim cadena As String

        Using mySqlConnection2 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection2.Open()
            Dim mySqlCommand2 As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection2)
            Dim myDataReader2 As Data.SqlClient.SqlDataReader
            myDataReader2 = mySqlCommand2.ExecuteReader()
            'myDataReader2.Read()

            cadena = ""

            Do While (myDataReader2.Read())
                If myDataReader2.IsDBNull(0) = False Then
                    cadena = myDataReader2.GetString(0)
                Else
                    cadena = "Null"
                End If
            Loop

            myDataReader2.Close()
            mySqlConnection2.Close()
        End Using

        Return cadena
    End Function

    'funcion que retorna un valor entero
    Function retornarentero(ByVal consulta As String, ByVal mystring As String) As Integer

        Dim entero As Integer

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            entero = myDataReader1.GetInt32(0)
            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return entero
    End Function

    'funcion que retorna un valor entero corto
    Function retornarbyte(ByVal consulta As String, ByVal mystring As String) As Byte

        Dim entero As Byte

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            entero = myDataReader1.GetByte(0)
            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return entero
    End Function

    'funcion que retorna una fecha en formato de cadena año/mes/dia
    Function retornafecha(ByVal consulta As String, ByVal mystring As String) As String

        Dim fecha As String

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            fecha = myDataReader1.GetDateTime(0).ToString("yyyy/MM/dd")
            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return fecha
    End Function

    Function retornafechaSinHora(ByVal consulta As String, ByVal mystring As String) As String

        Dim fecha As String

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            If myDataReader1.HasRows = True Then
                fecha = myDataReader1.GetDateTime(0).ToString("dd/MM/yyyy")
            Else
                fecha = ""
            End If
            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return fecha
    End Function

    Sub insertarmodificareliminar(ByVal Query As String, ByVal MyConstring As String)

        Using mySqlConnection As New System.Data.SqlClient.SqlConnection(MyConstring)
            mySqlConnection.Open()
            Dim mySqlCommandUpdate As New System.Data.SqlClient.SqlCommand(Query, mySqlConnection)
            Dim Filas As Integer = mySqlCommandUpdate.ExecuteNonQuery()
            mySqlConnection.Close()
        End Using

    End Sub




    'funcion que retorna un valor decimal
    ' si la consulta da NULL devuelve 0 
    ' si la consulta retorna valor devuelve el valor
    Function retornarsuma(ByVal consulta As String, ByVal mystring As String) As Decimal

        Dim VarDecimal As Decimal

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            If myDataReader1.IsDBNull(0) Then
                VarDecimal = 0
            Else
                VarDecimal = myDataReader1.GetDecimal(0)
            End If

            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return VarDecimal

    End Function

    'funcion que retorna un valor decimal
    Function retornarboolean(ByVal consulta As String, ByVal mystring As String) As Boolean

        Dim VarBoolean As Boolean

        Using mySqlConnection1 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection1.Open()
            Dim mySqlCommand As New System.Data.SqlClient.SqlCommand(consulta, mySqlConnection1)
            Dim myDataReader1 As Data.SqlClient.SqlDataReader
            myDataReader1 = mySqlCommand.ExecuteReader()
            myDataReader1.Read()
            VarBoolean = myDataReader1.GetBoolean(0)
            myDataReader1.Close()
            mySqlConnection1.Close()
        End Using

        Return VarBoolean

    End Function


    Function Obtener_Cadena_Conexion() As String
        Dim valor As String
        Dim cnnstring As String
        Dim key() As Byte
        Dim IV() As Byte


        valor = CStr(CULng(CInt("180014") ^ 3))
        key = System.Text.Encoding.ASCII.GetBytes(valor)
        IV = System.Text.Encoding.ASCII.GetBytes(valor)

        'cnnstring = System.Configuration.ConfigurationManager.AppSettings("servidor").ToString
        'cnnstring = "orbGYHc3IglcRQ0bCHjciNMt8mYwV/kpD4m8x4z8CBaZ+OYMTxp8PRsVfvY4krQUDel9BAay8UHPPJXDPWihExm43QVTdYtm9CtNo7ufcHVVz1q8elaT2LtFl9OLGg2pT3ILks9k4VXYl+XGOUjKFA=="

        'cnnstring = "User ID=Aplicacion;Database=GuatemalaDigital;PASSWORD=Guate123581321;server=172.30.0.187"
        cnnstring = "User ID=Aplicacion;Database=GuatemalaDigital;PASSWORD=Guate123581321;server=guatemaladigital.c6aj5b1cqmrh.us-east-1.rds.amazonaws.com"

        'Dim mb() As Byte = Convert.FromBase64String(cnnstring)
        'Dim myConstring As String = decryptStringFromBytes_AES(mb, key, IV)
        Dim myConstring As String = cnnstring

        Return myConstring
    End Function


    Function Obtener_Cadena_Conexion_Desarrollo() As String
        Dim valor As String
        Dim cnnstring As String
        Dim key() As Byte
        Dim IV() As Byte


        valor = CStr(CULng(CInt("180014") ^ 3))
        key = System.Text.Encoding.ASCII.GetBytes(valor)
        IV = System.Text.Encoding.ASCII.GetBytes(valor)

        cnnstring = "User ID=Aplicacion;Database=GuatemalaDigital;PASSWORD=Guate123581321;server=GUATEMALADIGITAL.org" 'pruebas
        'cnnstring = "User ID=Aplicacion;Database=GuatemalaDigital;PASSWORD=Guate123581321;server=localhost"
        Dim myConstring As String = cnnstring

        Return myConstring
    End Function

    Function Obtener_Cadena_Conexion_Localhost() As String
        Dim valor As String
        Dim cnnstring As String
        Dim key() As Byte
        Dim IV() As Byte


        valor = CStr(CULng(CInt("180014") ^ 3))
        key = System.Text.Encoding.ASCII.GetBytes(valor)
        IV = System.Text.Encoding.ASCII.GetBytes(valor)

        'cnnstring = System.Configuration.ConfigurationManager.AppSettings("servidor").ToString
        'cnnstring = "orbGYHc3IglcRQ0bCHjciNMt8mYwV/kpD4m8x4z8CBaZ+OYMTxp8PRsVfvY4krQUDel9BAay8UHPPJXDPWihExm43QVTdYtm9CtNo7ufcHVVz1q8elaT2LtFl9OLGg2pT3ILks9k4VXYl+XGOUjKFA=="

        'MAQUINA WALTER
        cnnstring = "Data Source=WALTER-PC; Initial Catalog=Desarrollo; user id=sa; Password=12345; Timeout=0"

        'MAQUINA DIEGO
        'cnnstring = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GuatemalaDigital;Data Source=DIEGO\GUATEMALADIGITAL;Connection Timeout=0"

        'MAQUINA GLORIA
        'cnnstring = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DESARROLLO;Data Source=DESA-GLORIA\SQLEXPRESS;Connection Timeout=0"

        'Dim mb() As Byte = Convert.FromBase64String(cnnstring)
        'Dim myConstring As String = decryptStringFromBytes_AES(mb, key, IV)
        Dim myConstring As String = cnnstring

        Return myConstring
    End Function

    Function decryptStringFromBytes_AES(ByVal cipherText() As Byte, ByVal Key() As Byte, ByVal IV() As Byte) As String

        If cipherText Is Nothing OrElse cipherText.Length <= 0 Then
            Throw New ArgumentNullException("cipherText")
        End If
        If Key Is Nothing OrElse Key.Length <= 0 Then
            Throw New ArgumentNullException("Key")
        End If
        If IV Is Nothing OrElse IV.Length <= 0 Then
            Throw New ArgumentNullException("IV")
        End If

        Dim aesAlg As RijndaelManaged = Nothing
        Dim plaintext As String = Nothing

        Try
            aesAlg = New RijndaelManaged()

            aesAlg.Key = Key

            aesAlg.IV = IV
            Dim decryptor As ICryptoTransform = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV)

            Using msDecrypt As New MemoryStream(cipherText)
                Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)
                    Using srDecrypt As New StreamReader(csDecrypt)

                        plaintext = srDecrypt.ReadToEnd()

                    End Using
                End Using
            End Using

        Catch ex As Exception

            'lblErrorCodigo.Text = "Código no valido"


        Finally

            If Not (aesAlg Is Nothing) Then
                aesAlg.Clear()
            End If
        End Try
        Return plaintext


    End Function

    'Devuelve una cadena con los valores obtenidos de la tabla parametros separados por comas
    'recibe una cadena con los codigos de los parametros separados por comas
    ' si se envía ("1,2") obtendra los valores de los parámetros con codigodeparametro 1 (tasa de cambio) y 2 (Costo fijo)
    Function Obtener_Parametros(ByVal CodigoParametros As String, ByVal mystring As String) As String
        Dim Cadena As String
        Dim Consulta As String

        Consulta = "select Valor from Parametro where CodigoParametro in (" & CodigoParametros & ") order by CodigoParametro"

        Cadena = ""
        Using mySqlConnection2 As New System.Data.SqlClient.SqlConnection(mystring)
            mySqlConnection2.Open()
            Dim mySqlCommand2 As New System.Data.SqlClient.SqlCommand(Consulta, mySqlConnection2)
            Dim myDataReader2 As Data.SqlClient.SqlDataReader
            myDataReader2 = mySqlCommand2.ExecuteReader()


            Do While myDataReader2.Read()
                If myDataReader2.IsDBNull(0) = False Then
                    If Cadena = "" Then
                        Cadena = myDataReader2.GetString(0)
                    Else
                        Cadena = Cadena & "," & myDataReader2.GetString(0)
                    End If
                End If


            Loop



            myDataReader2.Close()
            mySqlConnection2.Close()
        End Using

        Obtener_Parametros = Cadena

    End Function

    'función traduce de ingles a español 
    Function Traducir_Espanol(ByVal Datos As String) As String
        Dim Url As String
        Dim Cadena As String
        Dim Resultado As String
        Dim DatoOriginal As String

        DatoOriginal = Datos
        Guardar_Datos_Archivo_Texto_Traduccion("Función Traducir Espanol Texto Original = " & Datos)


        Try
            Resultado = Datos

            If Trim(Datos) <> "" Then


                Datos = Replace(Datos, " ", "%20")
                Datos = Replace(Datos, "&", "and")

                Url = "https://www.googleapis.com/language/translate/v2?key=AIzaSyBtGchgxOpaOy8Oi380x9LJw3xANxL2vdE&source=en&target=es&q=" & Datos

                Using client As New WebClient
                    client.Encoding = System.Text.Encoding.UTF8
                    Cadena = client.DownloadString(Url)
                    'Cadena = client.Encoding.
                End Using



                Dim jo = JObject.Parse(Cadena)

                Resultado = jo("data")("translations")(0)("translatedText").ToString
                'Resultado = Server.HtmlDecode(Resultado)

            End If
            Traducir_Espanol = Resultado

        Catch ex As Exception
            Traducir_Espanol = DatoOriginal
        End Try


    End Function

    Function Reemplazar_Cadena_Url(ByVal Campo As String) As String
        Campo = "REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(" & Campo & ", 'á', 'a'), 'é','e'), 'í', 'i'), 'ó', 'o'), 'ú','u'),'ñ','n'),' ','-'),',',''),'#',''),'?',''),'+',''),'""',''),'%',''),'.',''),char(10),''),char(13),''),char(160),'') "

        Reemplazar_Cadena_Url = Campo
    End Function

    Function Reemplazar_Cadena_Url_VB(ByVal Cadena As String) As String
        Cadena = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Cadena, "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u"), "ñ", "n"), " ", "-"), ",", ""), "#", ""), "?", ""), "+", ""), """", ""), "%", ""), ".", ""), Chr(10), ""), Chr(13), ""), Chr(160), "")

        Reemplazar_Cadena_Url_VB = Cadena
    End Function

    Function Longitud_Url(ByVal Cadena As String) As String
        If Len(Cadena) > 250 Then
            Cadena = Left(Cadena, 249)

            If Cadena.Substring(Len(Cadena) - 1, 1) = "%" Then
                Cadena = Left(Cadena, Len(Cadena) - 1)
            ElseIf Cadena.Substring(Len(Cadena) - 2, 1) = "%" Then
                Cadena = Left(Cadena, Len(Cadena) - 2)
            End If

            Cadena = Cadena + "/"
        End If
        Longitud_Url = Cadena
    End Function


    Sub Guardar_Datos_Archivo_Texto_Traduccion(ByVal Cadena As String)
        Dim path As String = "C:\inetpub\wwwroot\Sistema\CARGA\TraduccionesWindowsPromocion" & Date.Now.Day.ToString & Date.Now.Month.ToString & ".txt"

        ' This text is added only once to the file. 
        If Not File.Exists(path) Then
            ' Create a file to write to.
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Date.Now.ToString)
                sw.WriteLine(Cadena)
                sw.WriteLine(" ")
            End Using

        Else
            ' This text is always added, making the file longer over time 
            ' if it is not deleted.
            Using sw As StreamWriter = File.AppendText(path)
                sw.WriteLine(Date.Now.ToString)
                sw.WriteLine(Cadena)
                sw.WriteLine(" ")
            End Using

        End If

    End Sub



    Sub Guardar_Datos_Archivo_Texto_ErrorPromocion(ByVal Cadena As String)
        Dim path As String = "C:\inetpub\wwwroot\Sistema\CARGA\ErrorPromocion" & Date.Now.Day.ToString & Date.Now.Month.ToString & ".txt"

        ' This text is added only once to the file. 
        If Not File.Exists(path) Then
            ' Create a file to write to.
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Date.Now.ToString)
                sw.WriteLine(Cadena)
                sw.WriteLine(" ")
            End Using

        Else
            ' This text is always added, making the file longer over time 
            ' if it is not deleted.
            Using sw As StreamWriter = File.AppendText(path)
                sw.WriteLine(Date.Now.ToString)
                sw.WriteLine(Cadena)
                sw.WriteLine(" ")
            End Using

        End If

    End Sub

    Function ejecuta_query_dt(ByVal query As String, ByRef dt As DataTable, ByVal mystring As String) As DataTable
        dt.Clear()

        Using cnn As New SqlConnection(mystring)
            cnn.Open()
            Using dad As New SqlDataAdapter(query, cnn)
                dad.SelectCommand.CommandTimeout = 180
                Try
                    dad.Fill(dt)
                Catch ex As SqlException

                End Try
                dad.Dispose()
            End Using
            cnn.Close()
        End Using

        ejecuta_query_dt = dt
    End Function

End Class

