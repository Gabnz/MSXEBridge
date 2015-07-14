Imports System
Imports System.Runtime.InteropServices

Namespace Bridge

    Public Interface _Bridge
        <DispId(1)> Function Beep() As Boolean
        <DispId(2)> Function abrirPuerto() As Boolean
        <DispId(3)> Function cerrarPuerto() As Boolean
        <DispId(4)> Function leerBlanco() As Boolean
        <DispId(5)> Function leerNegro() As Boolean
        <DispId(6)> Function medirMuestra()
    End Interface

    'esta Guid es la que toma en cuenta Qt al momento de establecer el control del QAxObject
    <Guid("670F95E1-FEAD-46F7-B53F-6407A2B3B2A4"), _
     ClassInterface(ClassInterfaceType.None), _
     ProgId("MSXE.Bridge")>
    Public Class Bridge

        Implements _Bridge

        Public Bridge()

        Dim sensor As MiniScanXE.MSXE
        Dim log As MiniScanXE.MSLog
        Dim setup As MiniScanXE.MSSetup

        Const PUNTOS_ESPECTRALES As Integer = 31
        Public Sub New()
            sensor = New MiniScanXE.MSXE
            log = New MiniScanXE.MSLog
            setup = New MiniScanXE.MSSetup
        End Sub

        Public Function Beep() As Boolean Implements _Bridge.Beep

            Dim respuesta As Boolean = True

            Try
                sensor.Beep()
            Catch ex As Exception
                respuesta = False
            End Try

            Return respuesta
        End Function

        Public Function abrirPuerto() As Boolean Implements _Bridge.abrirPuerto

            Dim i As Integer = 1
            sensor.BaudRate = 9600
            Dim conectado As Boolean = False
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

        Public Function cerrarPuerto() As Boolean Implements _Bridge.cerrarPuerto

            Dim desconectado As Boolean = True

            Try
                sensor.CloseCommPort()
            Catch ex As Exception
                desconectado = False
            End Try

            Return desconectado
        End Function

        Public Function leerBlanco() As Boolean Implements _Bridge.leerBlanco

            Dim leido As Boolean = True

            Try
                sensor.ReadWhite()
            Catch ex As Exception
                leido = False
            End Try

            Return leido
        End Function

        Public Function leerNegro() As Boolean Implements _Bridge.leerNegro

            Dim leido As Boolean = True

            Try
                sensor.ReadBlack()
            Catch ex As Exception
                leido = False
            End Try

            Return leido
        End Function

        Public Function medirMuestra() Implements _Bridge.medirMuestra
            Dim resultado(PUNTOS_ESPECTRALES) As Single
            Try
                sensor.ReadSample(resultado)

                'el array resultante tiene 32 elementos, puesto que va desde cero (0) hasta 31
                'el metodo ReadSample llena el array desde uno (1) hasta 31, asi que hay que omitir el elemento cero (0) del array
                For i = 1 To UBound(resultado)
                    resultado(i - 1) = resultado(i)
                Next i

                ReDim Preserve resultado(UBound(resultado) - 1)
            Catch ex As Exception
                
            End Try

            Return resultado
        End Function
    End Class

End Namespace