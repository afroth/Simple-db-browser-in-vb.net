'We will be working with SQL Server Database files
Imports System.Data.SqlClient

Public Class DOAcustomer
    '------------------------------------------------------------
    '-                File Name : DOAcustomer                   - 
    '-                Part of Project: rothhw8                  -
    '------------------------------------------------------------
    '-                Written By: Adam F Roth                   -
    '-                Written On: April 2, 2022                 -
    '------------------------------------------------------------
    '- File Purpose: creates a databse and populates the tables.
    '------------------------------------------------------------
    '- Program Purpose: creates the databse if not created 
    '- and populates the tables owners and vehicles
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically)
    '- strCONNECTION - string for database location.
    '- strDBPATH - path of where databse is created to.
    '------------------------------------------------------------
    Private Const strDBNAME As String = "Autos" 'Name of database
    Private Const strSERVERNAME As String = "(localdb)\MSSQLLocalDB"

    Private strDBPATH As String = Environment.CurrentDirectory &
                                  "\" & strDBNAME & ".mdf"

    Private strCONNECTION As String = "SERVER=" & strSERVERNAME & ";DATABASE=" &
                     strDBNAME & ";Integrated Security=SSPI;AttachDbFileName=" &
                     strDBPATH

    Public Sub dataBaseInitialize()
        '------------------------------------------------------------
        '-   Subprogram Name: dataBaseInitialize                    - 
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: function call to create database
        '- empty existing and repopulate it.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        'If the database doesn't exist, create it
        If Not (System.IO.File.Exists(strDBPATH)) Then
            createDB()
        End If
        'Make sure all tables are cleaned out each time we run this
        cleanOwnersTbl()
        cleanVehiclesTbl()

        'Put some data into the tables
        popOwnersTbl()
        popVehiclesTbl()


    End Sub


    Private Sub createDB()
        '------------------------------------------------------------
        '-   Subprogram Name: createDB                              - 
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: creates a databse if it does not
        '- and creates the tables owners and vehicles
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBConn  - sql connection for adding in data
        '- strSQLCmd - string to hold sql query
        '- DBCmd - sql COmmand to execure query.
        '------------------------------------------------------------
        'Let's build a SQL Server database from scratch
        Dim DBConn As SqlConnection
        Dim strSQLCmd As String
        Dim DBCmd As SqlCommand

        'All we need to do initially is just point at the server
        DBConn = New SqlConnection("Server=" & strSERVERNAME)

        'Let's write a SQL DDL Command to build the database
        'There are a lot of other parameters but we can let them default
        'All we need are these three
        strSQLCmd = "CREATE DATABASE " & strDBNAME & " On " &
                    "(NAME = '" & strDBNAME & "', " &
                    "FILENAME = '" & strDBPATH & "')"

        DBCmd = New SqlCommand(strSQLCmd, DBConn)

        Try
            'Open the connection and try running the command
            DBConn.Open()
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Database was successfully created", "",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            'If we can't build the database, we are dead in the water so bail...
            MessageBox.Show(ex.ToString())
            MessageBox.Show("Cannot build database!  Closing program down...")
            End
        End Try

        'We are currently pointing at the [MASTER] database, so we
        'need to close the connection and reopen it pointing at the
        'Registration database...
        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If

        'Now we need to use the full connection string with the Integrated 
        'Security line, et cetera
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        'Build the Tables one at a time

        'Build the Owners Table by writing the SQL DDL Command
        DBCmd.CommandText = "CREATE TABLE Owners (" &
                             "TUID INT, " &
                             "FirstName varchar(50), " &
                             "LastName varchar(50), " &
                             "StreetAddress varchar(50), " &
                             "City varchar(50), " &
                             "State varchar(50), " &
                             "ZipCode varchar(50), " &
                             "PRIMARY KEY CLUSTERED([TUID] ASC))"

        DBCmd.Connection = DBConn
        Try
            DBCmd.ExecuteNonQuery()
            ' MessageBox.Show("Created Owners Table")
        Catch Ex As Exception
            '  MessageBox.Show("Owners Table Already Exists")
        End Try

        'Build the Vehicles Table by writing the SQL DDL Command
        DBCmd.CommandText = "CREATE TABLE Vehicles (" &
                             "TUID INT, " &
                             "OwnerID INT, " &
                             "Make varchar(50), " &
                             "Model varchar(50), " &
                             "Color varchar(50), " &
                             "ModelYear INT, " &
                             "VIN nvarchar(50), " &
                             "PRIMARY KEY CLUSTERED([TUID] ASC))"
        DBCmd.Connection = DBConn
        Try
            DBCmd.ExecuteNonQuery()
            '  MessageBox.Show("Created Vehicles Table")
        Catch Ex As Exception
            '  MessageBox.Show("Vehicles Table Already Exists")
        End Try

        'We can check to see if we're open before trying to
        'issue a connection close
        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If

    End Sub

    Private Sub popOwnersTbl()
        '------------------------------------------------------------
        '-   Subprogram Name: popOwnersTbl                          - 
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose:  populates the owners table
        '-
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBConn  - sql connection for adding in data
        '- strSQLCmd - string to hold sql query                     -
        '------------------------------------------------------------
        Dim DBConn As SqlConnection
        Dim DBCmd As SqlCommand = New SqlCommand()

        'Now try to open up a connection to the database
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        DBCmd.CommandText = "INSERT INTO Owners ( TUID,FirstName, LastName, StreetAddress, City, State, ZipCode) " &
           "VALUES ('1','Tom','Thomas', '123 Elm', 'Saginaw', 'MI', '48604')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBCmd.CommandText = "INSERT INTO Owners ( TUID,FirstName, LastName, StreetAddress, City, State, ZipCode) " &
           "VALUES ('2','Jane','Jones', '456 Pine', 'Saginaw', 'MI', '48605')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBCmd.CommandText = "INSERT INTO Owners (TUID,FirstName, LastName, StreetAddress, City, State, ZipCode) " &
           "VALUES ('3','Bob','Fredericks', '789 Maple', 'Birch Run', 'MI', '48415')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBConn.Close()
        ' MessageBox.Show("Added Owners To owners Table")

    End Sub

    Private Sub popVehiclesTbl()
        '------------------------------------------------------------
        '-   Subprogram Name: popVehiclesTbl                        - 
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: populates the vehicles table
        '-
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBConn  - sql connection for adding in data
        '- strSQLCmd - string to hold sql query                     -
        '------------------------------------------------------------
        Dim DBConn As SqlConnection
        Dim DBCmd As SqlCommand = New SqlCommand()

        'Now try to open up a connection to the database
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        DBCmd.CommandText = "INSERT INTO Vehicles (TUID,OwnerID, Make, Model, Color, ModelYear, VIN) " &
           "VALUES ('1','1','Chevy', 'Camaro', 'Blue', '2018', '14XA1394394')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBCmd.CommandText = "INSERT INTO Vehicles (TUID,OwnerID, Make, Model, Color, ModelYear, VIN) " &
           "VALUES ('2','2','Ford', 'F-150', 'Red', '2017', '2A7764747236')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBCmd.CommandText = "INSERT INTO Vehicles (TUID,OwnerID, Make, Model, Color, ModelYear, VIN) " &
           "VALUES ('3','2','Dodge', 'Dart', 'Red', '2017', '56B6D7667')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBCmd.CommandText = "INSERT INTO Vehicles (TUID,OwnerID, Make, Model, Color, ModelYear, VIN) " &
           "VALUES ('4','3','Kia', 'Soul', 'Green', '2013', '1A1467464484')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBCmd.CommandText = "INSERT INTO Vehicles (TUID,OwnerID, Make, Model, Color, ModelYear, VIN) " &
           "VALUES ('5','3','Dodge', 'Viper', 'Yellow', '2014', '48J764E7633')"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()

        DBConn.Close()
        ' MessageBox.Show("Added Vehicles To Vehicles Table")
    End Sub

    Private Sub cleanOwnersTbl()
        '------------------------------------------------------------
        '-   Subprogram Name: cleanOwnersTbl                        - 
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: deletes all existing data in the oweners
        '- table
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBConn  - sql connection for adding in data
        '- strSQLCmd - string to hold sql query                     -
        '------------------------------------------------------------
        '
        Dim DBConn As SqlConnection
        Dim DBCmd As SqlCommand = New SqlCommand()

        'Now try to open up a connection to the database
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        'Use SQL DML to zap the contents of the table
        DBCmd.CommandText = "DELETE FROM Owners"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()
        ' MessageBox.Show("Deleted Everything In Registration")

        DBConn.Close()

    End Sub

    Private Sub cleanVehiclesTbl()
        '------------------------------------------------------------
        '-   Subprogram Name: cleanVehiclesTbl                      - 
        '------------------------------------------------------------
        '-                Written By: Adam F Roth                   -
        '-                Written On: April 3, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose: deletes all existing data in the vehicles
        '- table.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):
        '- 
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBConn  - sql connection for adding in data
        '- strSQLCmd - string to hold sql query                     -
        '------------------------------------------------------------
        Dim DBConn As SqlConnection
        Dim DBCmd As SqlCommand = New SqlCommand()

        'Now try to open up a connection to the database
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        'Use SQL DML to zap the contents of the table
        DBCmd.CommandText = "DELETE FROM Vehicles"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()
        ' MessageBox.Show("Deleted Everything In Registration")

        DBConn.Close()

    End Sub
End Class
