Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports IWshRuntimeLibrary

Namespace Tarea2AYOD
    Partial Public Class btnLeerAccesoDirecto
        Inherits Form

        Private filePath As String = "datos.txt"
        Private dataFilePath As String = "datos.dat"
        Private indexFilePath As String = "index.dat"

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnEscribir_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim data As String = txtData.Text

            If Not String.IsNullOrEmpty(data) Then

                Try

                    Using writer As StreamWriter = New StreamWriter(filePath, True)
                        writer.WriteLine(data)
                    End Using

                    MessageBox.Show("Datos escritos en el archivo correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error al escribir en el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End Try
            Else
                MessageBox.Show("Ingrese datos antes de escribir en el archivo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub btnLeer_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try

                Using reader As StreamReader = New StreamReader(filePath)
                    Dim content As String = reader.ReadToEnd()
                    MessageBox.Show("Contenido del archivo:" & vbLf & content, "Lectura exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using

            Catch ex As Exception
                MessageBox.Show("Error al leer el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Try
        End Sub

        Private Sub btnEscribir2_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim data As String = txtData1.Text

            If Not String.IsNullOrEmpty(data) Then

                Try
                    Dim position As Long

                    Using dataFileStream As FileStream = New FileStream(dataFilePath, FileMode.Append, FileAccess.Write)

                        Using writer As StreamWriter = New StreamWriter(dataFileStream)
                            position = dataFileStream.Position
                            writer.WriteLine(data)
                        End Using
                    End Using

                    Using indexFileStream As FileStream = New FileStream(indexFilePath, FileMode.Append, FileAccess.Write)

                        Using indexWriter As BinaryWriter = New BinaryWriter(indexFileStream)
                            indexWriter.Write(position)
                        End Using
                    End Using

                    MessageBox.Show("Datos escritos en el archivo correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error al escribir en el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End Try
            Else
                MessageBox.Show("Ingrese datos antes de escribir en el archivo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub btnLeer2_Click(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Dim dataRecords As List(Of String) = New List(Of String)()

                Using indexFileStream As FileStream = New FileStream(indexFilePath, FileMode.Open, FileAccess.Read)

                    Using indexReader As BinaryReader = New BinaryReader(indexFileStream)

                        While indexFileStream.Position < indexFileStream.Length
                            Dim position As Long = indexReader.ReadInt64()

                            Using dataFileStream As FileStream = New FileStream(dataFilePath, FileMode.Open, FileAccess.Read)
                                dataFileStream.Seek(position, SeekOrigin.Begin)

                                Using dataReader As StreamReader = New StreamReader(dataFileStream)
                                    Dim data As String = dataReader.ReadLine()
                                    dataRecords.Add(data)
                                End Using
                            End Using
                        End While
                    End Using
                End Using

                Dim content As String = String.Join(Environment.NewLine, dataRecords)
                MessageBox.Show("Contenido del archivo:" & vbLf & content, "Lectura exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error al leer el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Try
        End Sub

        Private Sub btnCrearAccesoDirect_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim archivoObjetivo As String = txtRutaArchivo.Text
            Dim rutaAccesoDirecto As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MiAccesoDirecto.lnk")
            Dim accesoDirecto = TryCast(New WshShell().CreateShortcut(rutaAccesoDirecto), IWshShortcut)

            If accesoDirecto IsNot Nothing Then
                accesoDirecto.TargetPath = archivoObjetivo
                accesoDirecto.Save()
                MessageBox.Show("Acceso directo creado con éxito.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub btnLeerAccesoDirect_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim rutaAccesoDirecto As String = txtRutaAccesoDirecto.Text

            Try

                If System.IO.File.Exists(rutaAccesoDirecto) Then
                    Dim accesoDirecto = TryCast(New WshShell().CreateShortcut(rutaAccesoDirecto), IWshShortcut)

                    If accesoDirecto IsNot Nothing Then
                        MessageBox.Show($"Información del acceso directo:\n\n" & $"Destino: {accesoDirecto.TargetPath}\n" & $"Icono: {accesoDirecto.IconLocation}\n" & $"Trabaja en: {accesoDirecto.WorkingDirectory}", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    MessageBox.Show("El archivo de acceso directo no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End If

            Catch ex As Exception
                MessageBox.Show($"Error al intentar leer el acceso directo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Try
        End Sub

        Private Sub btnModificarAccesoDirecto_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim rutaAccesoDirecto As String = txtRutaAccesoDirecto.Text

            If System.IO.File.Exists(rutaAccesoDirecto) Then
                Dim accesoDirecto = TryCast(New WshShell().CreateShortcut(rutaAccesoDirecto), IWshShortcut)

                If accesoDirecto IsNot Nothing Then
                    accesoDirecto.TargetPath = txtNuevoDestino.Text
                    accesoDirecto.Save()
                    MessageBox.Show("Acceso directo modificado con éxito.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("El archivo de acceso directo no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End If
        End Sub
    End Class
End Namespace