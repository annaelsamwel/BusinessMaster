' ***********************************************************************
' Assembly         : Corelibr
' Author           : Annael Samwel
' Created          : 08-18-2014
'
' Last Modified By : user
' Last Modified On : 05-08-2014
' ***********************************************************************
' <copyright file="DataManager.vb" company="">
'     Copyright ©  2011
' </copyright>
' <summary></summary>
' ***********************************************************************
Imports System.Windows.Forms

''' <summary>
''' Class DataManager.
''' </summary>
Public Class DataManager
#Region "Events"
    ''' <summary>
    ''' Occurs when [current changed].
    ''' </summary>
    Public Event CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs)
#End Region
#Region "Fields"
    ''' <summary>
    ''' The adapter
    ''' </summary>
    Private Adapter As SqlClient.SqlDataAdapter

    ''' <summary>
    ''' The _ data table
    ''' </summary>
    Private WithEvents _DataTable As DataTable
    ''' <summary>
    ''' The _ binding source
    ''' </summary>
    Private WithEvents _BindingSource As BindingSource

#End Region
#Region "Property"
    ''' <summary>
    ''' Gets the user identifier.
    ''' </summary>
    ''' <value>The user identifier.</value>
    Public ReadOnly Property UserID As Integer
        Get
            Return Devpp.Common.Login.UserID
        End Get
    End Property
    ''' <summary>
    ''' Gets the data table.
    ''' </summary>
    ''' <value>The data table.</value>
    Public ReadOnly Property DataTable() As DataTable
        Get
            Return _DataTable
        End Get
    End Property
    ''' <summary>
    ''' Gets the binding source.
    ''' </summary>
    ''' <value>The binding source.</value>
    Public ReadOnly Property BindingSource() As BindingSource
        Get
            Return _BindingSource
        End Get
    End Property
#End Region
#Region "Methods"
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DataManager"/> class.
    ''' </summary>
    Public Sub New()

        _DataTable = New DataTable
        _BindingSource = New BindingSource
        _BindingSource.DataSource = _DataTable
    End Sub
    ''' <summary>
    ''' Executes the sp.
    ''' </summary>
    ''' <param name="spName">Name of the sp.</param>
    ''' <param name="param">The parameter.</param>
    ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
    Public Function ExecuteSP(ByVal spName As String, ByVal ParamArray param() As Object) As Boolean
        Try

            _DataTable = New DataTable
            _BindingSource = New BindingSource
            _BindingSource.DataSource = _DataTable
            Adapter = New SqlClient.SqlDataAdapter(Devpp.Data.SQLSERVER.GetSPSQLCom(spName, param))
            Adapter.Fill(_DataTable)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

    End Function
    ''' <summary>
    ''' Refills this instance.
    ''' </summary>
    Public Sub Refill()
        Try
            _DataTable.Rows.Clear()
            Adapter.Fill(_DataTable)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

    ''' <summary>
    ''' Handles the CurrentChanged event of the _BindingSource control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
    Private Sub _BindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles _BindingSource.CurrentChanged
        RaiseEvent CurrentChanged(sender, e)
    End Sub

    ''' <summary>
    ''' Handles the ColumnChanging event of the _DataTable control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.Data.DataColumnChangeEventArgs"/> instance containing the event data.</param>
    Private Sub _DataTable_ColumnChanging(ByVal sender As Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles _DataTable.ColumnChanging
        e.Row.EndEdit()
    End Sub
    ''' <summary>
    ''' Writes the card.
    ''' </summary>
    ''' <param name="CardNo">The card no.</param>
    ''' <param name="Block">The block.</param>
    Public Sub WriteCard(ByVal CardNo As Integer, ByVal Block As String)

    End Sub
End Class