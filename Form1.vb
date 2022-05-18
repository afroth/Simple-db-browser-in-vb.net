'We need to bring in the SqlClient namespace
Imports System.Data.SqlClient
Imports System.ComponentModel

Public Class Form1
    '------------------------------------------------------------
    '-                File Name : Form1.vb                      - 
    '-                Part of Project: rothhw8                  -
    '------------------------------------------------------------
    '-                Written By: Adam F Roth                   -
    '-                Written On: April 2, 2022                 -
    '------------------------------------------------------------
    '- File Purpose:  fire when program is initially loaded
    '- Main form for program. Database is created when you run the
    '- program.
    '------------------------------------------------------------
    '- Program Purpose: allows user to insert, update,delete
    '- information into the databse table owners and view
    '- the vehicles that owner has.
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically)
    '- blnAddSave - boolean to keep track if the add button was
    '- clicked before the save button.
    '- DBAdaptOwners - Database adaptor to tables Owners.
    '- DBAdaptVehicles - Database adaptor to tables Vehicles.
    '- dsOwners - data set to hold data of table Owners.
    '- dsVehicles - data set to hold data of table Owners.
    '- intRowCounter - keeps track of row count for next incert.
    '- myConn - database connection.
    '- strDBPATH - string of databse location.
    '------------------------------------------------------------
    'Name of database
    Const strDBNAME As String = "Autos"

    'Name of the database server
    Const strSERVERNAME As String = "(localdb)\MSSQLLocalDB"

    'Path to database in executable
    Dim strDBPATH As String = Environment.CurrentDirectory &
                                  "\" & strDBNAME & ".mdf"

    'This is the full connection string
    Dim strCONNECTION As String = "SERVER=" & strSERVERNAME & ";DATABASE=" &
                 strDBNAME & ";Integrated Security=SSPI;AttachDbFileName=" &
                 strDBPATH

    'creating connection object
    Dim myConn As New SqlConnection(strCONNECTION)
    'datasets for each table
    Dim dsOwners As New DataSet
    Dim dsVehicles As New DataSet
    'Sql Adapter for each table
    Dim DBAdaptOwners As SqlDataAdapter
    Dim DBAdaptVehicles As SqlDataAdapter
    'Integer to keep track of thelatest row in the table
    Dim intRowCounter As Integer
    'Boolean to keep track if add button was clicked after the save button
    Dim blnAddSave As Boolean = False

    '************************************************************************************
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '------------------------------------------------------------
        '-   Subprogram Name:  - Form1_Load
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: fires on form load
        '- populating textboxes, creating database.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):
        '- autoDB a new object of class DOAcustomer to create the database
        '- strSQLCmd - string to hold the sql query for table Owners
        '- strSQLCmdVehicles - string to hold the sql query for table Vehicle
        '------------------------------------------------------------
        Dim autoDB As New DOAcustomer()
        autoDB.dataBaseInitialize()
        Dim strSQLCmd As String
        Dim strSQLCmdVehicles As String

        'Load up all courses since they will never change while the program runs
        strSQLCmd = "Select * From Owners"
        strSQLCmdVehicles = "Select * From Vehicles"

        DBAdaptOwners = New SqlDataAdapter(strSQLCmd, strCONNECTION)
        DBAdaptVehicles = New SqlDataAdapter(strSQLCmdVehicles, strCONNECTION)

        DBAdaptOwners.Fill(dsOwners, "Owners")

        ' binds the text boxes to the data set to view owner information
        fillInForm()
        gridRefresh()
        intRowCounter = dsOwners.Tables("Owners").Rows.Count + 1

    End Sub

    '************************************************************************************
    Private Sub fillInForm()
        '------------------------------------------------------------
        '-   Subprogram Name:  - fillInForm
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: Binds the specified text boxes to the 
        '- related columns from table Owners
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        ' Binding the data from Owners table columns to specified text boxes
        txtFirstName.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.FirstName"))
        txtLastName.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.LastName"))
        txtAddress.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.StreetAddress"))
        txtCity.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.City"))
        txtState.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.State"))
        txtZip.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.ZipCode"))
        txtId.DataBindings.Add(New Binding("Text", dsOwners,
                                         "Owners.TUID"))
    End Sub

    '************************************************************************************
    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnFirst_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: changes postion to row 0
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        'Called when we click the move to first button <| in Owners listing 
        BindingContext(dsOwners, "Owners").Position = 0
        'Refresh Grid
        gridRefresh()
    End Sub

    '************************************************************************************
    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnPrev_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: changes table row position to postion -1
        '-
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        'Called when we click the move to previous button << in Owners listing 
        BindingContext(dsOwners, "Owners").Position = (BindingContext(dsOwners,
                       "Owners").Position - 1)
        'Refresh Grid
        gridRefresh()
    End Sub

    '************************************************************************************
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnNext_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: changes table row position to postion +1
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        'Called when we click the move to previous button >> in Owners listing 
        BindingContext(dsOwners, "Owners").Position = (BindingContext(dsOwners,
                       "Owners").Position + 1)
        'Refresh Grid
        gridRefresh()
    End Sub

    '************************************************************************************
    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnLast_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: changes table row position to last row
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        'Called when we click the move to last button |> in Owners listing 
        BindingContext(dsOwners, "Owners").Position =
                      (dsOwners.Tables("Owners").Rows.Count - 1)
        'Refresh Grid
        gridRefresh()
    End Sub
    '************************************************************************************
    Private Sub enableFormText(blnIsEnabled As Boolean)
        '------------------------------------------------------------
        '-   Subprogram Name:  - enabledFormText
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: Enables or disables text boxes based on
        '- the value passed in.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- blnEnabled - holds the 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        ' enabled or diables the text boxes
        txtFirstName.Enabled = blnIsEnabled
        txtLastName.Enabled = blnIsEnabled
        txtAddress.Enabled = blnIsEnabled
        txtCity.Enabled = blnIsEnabled
        txtState.Enabled = blnIsEnabled
        txtZip.Enabled = blnIsEnabled
        btnSave.Visible = blnIsEnabled
        btnCancel.Visible = blnIsEnabled

    End Sub

    '************************************************************************************
    ' this is fired when the add or update buttons are clicked so the user cannot reclick them
    ' during an add or update
    Private Sub enableFormButtons(blnIsEnabled As Boolean)


        ' enabled or diables the text boxes
        btnAdd.Visible = blnIsEnabled
        btnUpdate.Visible = blnIsEnabled
        btnDelete.Visible = blnIsEnabled

    End Sub
    '************************************************************************************
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnAdd_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: enables text boxes adds new data to
        '- Owner table.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- blnisEnabled - holds the boolean value to enable or disable
        '- the txt boxes.
        '------------------------------------------------------------
        Dim blnisEnabled = True
        Dim blnButtonEnabled = False

        ' Sub call to enable typing into the text fields
        enableFormText(blnisEnabled)
        enableFormButtons(blnButtonEnabled)
        'Clear out the current edits
        BindingContext(dsOwners, "Owners").EndCurrentEdit()
        ' Add a new record in the record set dsOwners
        BindingContext(dsOwners, "Owners").AddNew()

        blnAddSave = True

        dgvVehicles.Visible = False
    End Sub
    '************************************************************************************
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnUpdate_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: enables the text boxes to be edited
        '- on the current data set position.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- blnisEnabled - holds the boolean value to enable or disable
        '- the txt boxes.                                           -
        '------------------------------------------------------------

        Dim blnisEnabled = True
        Dim blnButtonEnabled = False
        ' Sub call to enable typing into the text fields
        enableFormText(blnisEnabled)
        enableFormButtons(blnButtonEnabled)

        dgvVehicles.Visible = True

    End Sub
    '************************************************************************************
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnSave_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: saves a new data row into the table
        '- or updates the current row based on previous button pressed.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSQLCmd - string to hold sql query
        '- cmdBuilder - cmd builder to get have insert command.
        '------------------------------------------------------------
        'Set up the data adapter for Students...
        Dim strSQLCmd As String
        Dim cmdBuilder As SqlCommandBuilder
        strSQLCmd = "Select * From Owners Where TUID = '" &
                    Trim(txtId.Text) & "'"

        DBAdaptOwners = New SqlDataAdapter(strSQLCmd, strCONNECTION)
        cmdBuilder = New SqlCommandBuilder(DBAdaptOwners)
        DBAdaptOwners.InsertCommand = cmdBuilder.GetInsertCommand

        Dim blnIsEnabled = False
        Dim blnButtonEnabled = True
        ' if the add new button was hit right before this button
        If blnAddSave = True Then
            txtId.Text = intRowCounter
            intRowCounter = +1
            blnAddSave = False
        End If

        'Stop any current edits.
        BindingContext(dsOwners, "Owners").EndCurrentEdit()

        'Update the database
        myConn.Open()
        DBAdaptOwners.Update(dsOwners, "Owners")
        myConn.Close()

        'Update the dataset to correspond with database.
        dsOwners.AcceptChanges()
        'disable the text boxes and hide the save and cancel button
        dgvVehicles.Visible = True
        enableFormText(blnIsEnabled)
        enableFormButtons(blnButtonEnabled)
        'Refresh the data view grid
        gridRefresh()
    End Sub
    '************************************************************************************
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnCancel_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: cancel button stops current edit
        '- 
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- blnisEnabled - holds the boolean value to enable or disable
        '- the txt boxes.                                                   -
        '------------------------------------------------------------
        'boolean so we can hide buttons and disable text boxes
        Dim blnIsEnabled = False
        Dim blnButtonEnabled = True
        'Cancel any current edits.
        BindingContext(dsOwners, "Owners").CancelCurrentEdit()
        'disable the text boxes and hide the save and cancel button
        enableFormText(blnIsEnabled)
        enableFormButtons(blnButtonEnabled)
        ' made the data grid view visable
        dgvVehicles.Visible = True
    End Sub
    '************************************************************************************
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        '------------------------------------------------------------
        '-   Subprogram Name:  - btnDelete_Click
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        'Dim intID As Integer = txtId.Text
        Dim strSQLCmd As String
        'Dim strSqlQuery As String = "Delete * From Vehicles Where OwnerID = '" & txtId.Text & "'"
        Dim cmdBuilder As SqlCommandBuilder
        ' creating sqlCOmmand object
        Dim DBCmd As SqlCommand = New SqlCommand()
        'sql query saved into sql command object
        DBCmd.CommandText = "DELETE FROM Vehicles WHERE OwnerID = '" & txtId.Text & "'"
        strSQLCmd = "Select * From Owners Where TUID = '" &
         Trim(txtId.Text) & "'"

        cmdBuilder = New SqlCommandBuilder(DBAdaptOwners)

        'Cancel any current edits.
        BindingContext(dsOwners, "Owners").CancelCurrentEdit()

        If MessageBox.Show("Are you sure you want to Delete?", Me.Name, MessageBoxButtons.YesNo) = DialogResult.Yes Then
            'Remove the owner record in current position
            BindingContext(dsOwners, "Owners").RemoveAt(BindingContext(dsOwners, "Owners").Position)
            'Update the database
            myConn.Open()
            'Assigning connection to SQl command for deletion
            DBCmd.Connection = myConn
            'Executing sql command for deletion
            DBCmd.ExecuteNonQuery()
            'fill in the dataset for Vehicles to repop after deletion
            DBAdaptVehicles.Fill(dsVehicles, "Vehicles")
            'update the Oweners table to represent the dataset
            DBAdaptOwners.Update(dsOwners, "Owners")
            ' update the vehicles table to represent the data set
            DBAdaptVehicles.Update(dsVehicles, "Vehicles")
            'close the connection
            myConn.Close()
            'Refresh our data grid view
            gridRefresh()
        Else

        End If

    End Sub
    '************************************************************************************
    Private Sub gridRefresh()
        '------------------------------------------------------------
        '-   Subprogram Name:  - gridRefresh
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: Shows the datagrid viewer where
        '- the vehicle table OwnerID is equal to Owners TUID
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSqlQuery - string holds a sql query
        '------------------------------------------------------------

        'sql query to get the inforation from the vehicles table.
        Dim strSqlQuery = "Select * From Vehicles Where OwnerID = '" & txtId.Text & "'"
        ' open connection
        myConn.Open()
        ' clear anything in the data set.
        dsVehicles.Clear()
        ' assign the sql query to the adapter
        DBAdaptVehicles = New SqlDataAdapter(strSqlQuery, myConn)
        ' filling the data set with information from the query
        DBAdaptVehicles.Fill(dsVehicles, "Vehicles")
        ' filling the data grid view with the data from the dataset
        dgvVehicles.DataSource = dsVehicles.Tables("Vehicles")
        ' Hide specified Column so they dont show in data grid view
        dgvVehicles.Columns("TUID").Visible = False
        dgvVehicles.Columns("OwnerID").Visible = False
        ' closing connection
        myConn.Close()
    End Sub
End Class
