Imports Simphony.Simulation

Public Class Technician

    Inherits Resource
    Public Property TechSkill As String
    Public Property TechName As String
    Public Property TechID As Integer
    Public Property UnitCost As Double
    Public Property PreviousTask As RepairTask
    Public Property InitialTechLocation As String
    Public Property TotalTaskCost As Double
    'for Greedy algorithm
    Public Property CurrentTask As RepairTask
    'for Dijkstar
    Public Property Path As New List(Of RepairTask)
    Public Sub New()
        Me.Servers = 0
    End Sub

    Public Sub New(profession As String)
        Me.TechSkill = profession
        Me.Servers = 0
        'Me.TechName = Name

    End Sub


End Class
