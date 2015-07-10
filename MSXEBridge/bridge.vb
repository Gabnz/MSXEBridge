﻿Imports System
Imports System.Runtime.InteropServices

Namespace Bridge

    Public Interface _Bridge
        <DispId(1)> Function prueba() As String
        <DispId(2)> Function Beep() As Boolean
        <DispId(3)> Function abrirPuerto() As Boolean
        <DispId(4)> Function cerrarPuerto() As Boolean
        <DispId(5)> Function leerBlanco() As Boolean
        <DispId(6)> Function leerNegro() As Boolean
        <DispId(7)> Function medirMuestra() As Array
    End Interface

    'esta Guid es la que toma en cuenta Qt al momento de establecer el control del QAxObject
    <Guid("670F95E1-FEAD-46F7-B53F-6407A2B3B2A4"), _
     ClassInterface(ClassInterfaceType.None), _
     ProgId("MSXE.Bridge")>
    Public Class Bridge

        Implements _Bridge

        Public Bridge()

        Dim sensor As MiniScanXE.MSXE
        Const PUNTOS_ESPECTRALES As Integer = 31
        Public Sub New()
            sensor = New MiniScanXE.MSXE
        End Sub
        Public Function prueba() As String Implements _Bridge.prueba

            Return "Esta es una prueba de comunicacion."
        End Function

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

        Public Function medirMuestra() As Array Implements _Bridge.medirMuestra
            Dim resultado(PUNTOS_ESPECTRALES) As Single
            Dim flag As Boolean = True
            Try
                sensor.ReadSample(resultado)
            Catch ex As Exception
                flag = False
            End Try

            If (flag) Then
                'el array resultante tiene 32 elementos, puesto que va desde cero (0) hasta 31
                'el metodo ReadSample llena el array desde uno (1) hasta 31, asi que hay que omitir el elemento cero (0) del array
                'probar estas dos formas de omitir el primer elemento, para ver cual funciona adecuadamente

                'For i = 1 To UBound(resultado)
                'resultado(i - 1) = resultado(i)
                'Next i
                'ReDim Preserve resultado(UBound(resultado) - 1)

                resultado = resultado.Skip(0).ToArray
            End If

            Return resultado
        End Function
    End Class

End Namespace