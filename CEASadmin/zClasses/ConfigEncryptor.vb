Imports System.Configuration

Module ConfigEncryptor
    ' Método para encriptar la sección de connectionStrings
    Public Sub EncryptConnectionStrings()
        Try
            ' Cargar la configuración del archivo App.config
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

            ' Obtener la sección connectionStrings
            Dim section As ConfigurationSection = config.GetSection("connectionStrings")

            If section IsNot Nothing AndAlso Not section.SectionInformation.IsProtected Then
                ' Encriptar la sección utilizando el proveedor 'DataProtectionConfigurationProvider'
                section.SectionInformation.ProtectSection("DataProtectionConfigurationProvider")
                section.SectionInformation.ForceSave = True

                ' Guardar los cambios
                config.Save(ConfigurationSaveMode.Modified)
                Console.WriteLine("Sección 'connectionStrings' encriptada correctamente.")
            Else
                Console.WriteLine("La sección 'connectionStrings' ya está encriptada o no existe.")
            End If
        Catch ex As Exception
            Console.WriteLine("Error al encriptar: " & ex.Message)
        End Try
    End Sub

    ' Método para desencriptar la sección de connectionStrings
    Public Sub DecryptConnectionStrings()
        Try
            ' Cargar la configuración del archivo App.config
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

            ' Obtener la sección connectionStrings
            Dim section As ConfigurationSection = config.GetSection("connectionStrings")

            If section IsNot Nothing AndAlso section.SectionInformation.IsProtected Then
                ' Desencriptar la sección
                section.SectionInformation.UnprotectSection()
                section.SectionInformation.ForceSave = True

                ' Guardar los cambios
                config.Save(ConfigurationSaveMode.Modified)
                Console.WriteLine("Sección 'connectionStrings' desencriptada correctamente.")
            Else
                Console.WriteLine("La sección 'connectionStrings' no está encriptada o no existe.")
            End If
        Catch ex As Exception
            Console.WriteLine("Error al desencriptar: " & ex.Message)
        End Try
    End Sub
End Module
