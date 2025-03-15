Option Strict Off
Option Explicit On
Friend Class frmAddIn
	Inherits System.Windows.Forms.Form
	Public VBInstance As VBIDE.VBE
	'UPGRADE_ISSUE: Connect objeto no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Public Connect As Connect
	
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Connect.Hide. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Connect.Hide()
	End Sub
	
	Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
		If Not IsNumeric(Me.txtLine.Text) Then
			MsgBox("Introduzca un número de línea correcto.")
		Else
			
			If CInt(Me.txtLine.Text) < 1 Or CInt(Me.txtLine.Text) > VBInstance.ActiveCodePane.CodeModule.CountOfLines Then
				MsgBox("Introduzca un número entre 1 y " & VBInstance.ActiveCodePane.CodeModule.CountOfLines)
			Else
				VBInstance.ActiveCodePane.SetSelection(CInt(Me.txtLine.Text), 1, CInt(Me.txtLine.Text), 1)
				'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Connect.Hide. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Connect.Hide()
			End If
		End If
		'    MsgBox "AddIn operation on: " & VBInstance.FullName
	End Sub
End Class