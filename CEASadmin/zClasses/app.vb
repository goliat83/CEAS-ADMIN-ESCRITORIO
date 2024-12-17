Public Class app
	Public Function CheckForm(_form As Form) As Form

		For Each f As Form In Application.OpenForms
			If f.Name = _form.Name Then
				Return f

			End If
		Next

		Return Nothing

	End Function

End Class
