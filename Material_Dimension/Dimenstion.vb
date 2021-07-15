Imports Material_Dimension.CATIA_Property
Public Class Dimenstion
    Private CATIAFactory As CATIA_Property = New CATIA_Property
    Private Sub Dimenstion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CATIAFactory = CATIA_Property.SetInitialCATIA
        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
        Me.CheckBox1.Enabled = True

        If Me.CheckBox1.Checked = True Then
            Me.TextBox7.Enabled = False
            Me.TextBox8.Enabled = False
            Me.TextBox9.Enabled = False
        End If
        Me.TopLevel = True
        Me.TopMost = True
        If CATIAFactory.PartDocument Is Nothing Then
            MsgBox("Please Open a CATPart file!!")
            End
        End If
        Call CATMain()
    End Sub
    Sub CATMain()
        Call BoundingBox(CATIAFactory.PartDocument)

        Dim Bbox_dx As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dx.1").ValueAsString
            Dim Bbox_dy As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dy.1").ValueAsString
            Dim Bbox_dz As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dz.1").ValueAsString
            Dim Offset_Bbox_Max_X As String = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Max_X.1").ValueAsString
            Dim Offset_Bbox_Max_Y As String = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Max_Y.1").ValueAsString
            Dim Offset_Bbox_Max_Z As String = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Max_Z.1").ValueAsString
            Dim Offset_Bbox_Min_X As String = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_X.1").ValueAsString
            Dim Offset_Bbox_Min_Y As String = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_Y.1").ValueAsString
            Dim Offset_Bbox_Min_Z As String = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_Z.1").ValueAsString


            Me.TextBox1.Text = Bbox_dx
            Me.TextBox2.Text = Bbox_dy
            Me.TextBox3.Text = Bbox_dz
            Me.TextBox4.Text = Offset_Bbox_Max_X
            Me.TextBox5.Text = Offset_Bbox_Max_Y
            Me.TextBox6.Text = Offset_Bbox_Max_Z
            Me.TextBox7.Text = Offset_Bbox_Min_X
            Me.TextBox8.Text = Offset_Bbox_Min_Y
            Me.TextBox9.Text = Offset_Bbox_Min_Z
            Me.Label5.Text = CATIAFactory.PartDocument.Part.Name

    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged             '    +X
        Dim Offset_Bbox_Max_X As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Max_X.1")
        Dim Offset_Bbox_Min_X As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_X.1")
        If CheckBox1.Checked = False Then
            If Me.TextBox4.Text <> "" Then
                Offset_Bbox_Max_X.ValuateFromString(TextBox4.Text + "mm")
                CATIAFactory.PartDocument.Part.Update()
            End If
        Else
            Offset_Bbox_Max_X.ValuateFromString(TextBox4.Text + "mm")
            TextBox7.Text = TextBox4.Text
            Offset_Bbox_Min_X.ValuateFromString(TextBox7.Text + "mm")
            CATIAFactory.PartDocument.Part.Update()
        End If
        Dim Bbox_dx As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dx.2").ValueAsString
        Me.TextBox1.Text = Bbox_dx
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged         '    -X
        Dim Offset_Bbox_Min_X As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_X.1")
        If Me.TextBox7.Text <> "" Then
            Offset_Bbox_Min_X.ValuateFromString(TextBox7.Text + "mm")
            CATIAFactory.PartDocument.Part.Update()
        End If
        Dim Bbox_dx As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dx.2").ValueAsString
        Me.TextBox1.Text = Bbox_dx
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Dim Offset_Bbox_Max_Y As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Max_Y.1")
        Dim Offset_Bbox_Min_Y As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_Y.1")
        If CheckBox1.Checked = False Then
            If Text <> "" Then
                Offset_Bbox_Max_Y.ValuateFromString(TextBox5.Text + "mm")
                CATIAFactory.PartDocument.Part.Update()
            End If
        Else
            Offset_Bbox_Max_Y.ValuateFromString(TextBox5.Text + "mm")
            TextBox8.Text = TextBox5.Text
            Offset_Bbox_Min_Y.ValuateFromString(TextBox8.Text + "mm")
            CATIAFactory.PartDocument.Part.Update()

        End If
        Dim Bbox_dy As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dy.2").ValueAsString
        Me.TextBox2.Text = Bbox_dy
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Dim Offset_Bbox_Min_Y As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_Y.1")
        If Text <> "" Then
            Offset_Bbox_Min_Y.ValuateFromString(TextBox8.Text + "mm")
            CATIAFactory.PartDocument.Part.Update()
        End If
        Dim Bbox_dy As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dy.2").ValueAsString
        Me.TextBox2.Text = Bbox_dy
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Dim Offset_Bbox_Max_Z As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Max_Z.1")
        Dim Offset_Bbox_Min_Z As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_Z.1")
        If CheckBox1.Checked = False Then
            If Text <> "" Then
                Offset_Bbox_Max_Z.ValuateFromString(TextBox6.Text + "mm")
                CATIAFactory.PartDocument.Part.Update()
            End If
        Else
            Offset_Bbox_Max_Z.ValuateFromString(TextBox6.Text + "mm")
            TextBox9.Text = TextBox6.Text
            Offset_Bbox_Min_Z.ValuateFromString(TextBox9.Text + "mm")
            CATIAFactory.PartDocument.Part.Update()
        End If
        Dim Bbox_dz As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dz.2").ValueAsString
        Me.TextBox3.Text = Bbox_dz
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        Dim Offset_Bbox_Min_Z As KnowledgewareTypeLib.Parameter = CATIAFactory.PartDocument.Part.Parameters.Item("Offset_Bbox_Min_Z.1")
        If Text <> "" Then
            Offset_Bbox_Min_Z.ValuateFromString(TextBox9.Text + "mm")
            CATIAFactory.PartDocument.Part.Update()
        End If
        Dim Bbox_dz As String = CATIAFactory.PartDocument.Part.Parameters.Item("Bbox_dz.2").ValueAsString
        Me.TextBox3.Text = Bbox_dz
    End Sub
    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        If CheckBox1.Checked = True Then
            TextBox7.Enabled = False
            TextBox8.Enabled = False
            TextBox9.Enabled = False
            TextBox7.Text = TextBox4.Text
            TextBox8.Text = TextBox5.Text
            TextBox9.Text = TextBox6.Text
        Else
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox9.Enabled = True
        End If

    End Sub

End Class
