
Public Class Bridge

    Dim sensor As MiniScanXE.MSXE
    Const PUNTOS_ESPECTRALES As Integer = 31
    Public Sub New()
        sensor = New MiniScanXE.MSXE
    End Sub
    Public Function prueba() As String

        Return "Esta es una prueba de comunicacion."
    End Function

    Public Function Beep() As Boolean

        Dim respuesta As Boolean = True

        Try
            sensor.Beep()
        Catch ex As Exception
            respuesta = False
        End Try

        Return respuesta
    End Function

    Public Function abrirPuerto() As Boolean

        Dim i As Integer = 1
        Dim conectado As Boolean = False
        sensor.BaudRate = 9600
        'prueba con cada uno de los puertos COM para conectar el miniscan
        While i < 5 And conectado = False
            sensor.CommPort = i
            Try
                sensor.OpenCommPort()
                conectado = True
            Catch ex As Exception
                i += 1
            End Try
        End While

        Return conectado
    End Function

    Public Function cerrarPuerto() As Boolean

        Dim desconectado As Boolean = True

        Try
            sensor.CloseCommPort()
        Catch ex As Exception
            desconectado = False
        End Try

        Return desconectado
    End Function

    Public Function leerBlanco() As Boolean

        Dim leido As Boolean = True

        Try
            sensor.ReadWhite()
        Catch ex As Exception
            leido = False
        End Try

        Return leido
    End Function

    Public Function leerNegro() As Boolean

        Dim leido As Boolean = True

        Try
            sensor.ReadBlack()
        Catch ex As Exception
            leido = False
        End Try

        Return leido
    End Function

    Public Function medirMuestra()
        Dim resultado(PUNTOS_ESPECTRALES) As Single

        Try
            sensor.ReadSample(resultado)
        Catch ex As Exception

        End Try

        Return resultado
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class


