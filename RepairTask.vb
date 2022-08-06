Imports Simphony.Simulation
Public Class RepairTask

    Inherits Entity
    Public Property TaskName As String
    Public Property TaskDescription As String
    Public Property Status As String

    Public Property Priority As Integer
    Public Property EsTime As Double
    Public Property ResourceN As Integer
    Public Property RequiredSkill As String
    Public Property UsedSkill As String
    Public Property TaskTech As Technician
    Public Property TaskFile As WaitingFile
    Public Property TechnicianID As String
    Public Property TechnicianName As String
    Public Property MustSkilled As Boolean
    Public Property ReportDate As Date
    Public Property AssignedTime As Date
    Public Property StartTime As Date
    Public Property EndTime As Date
    Public Property CostTech As Double = 0
    Public Property CostTravel As Double = 0
    Public Property CostDelay As Double = 0
    Public Property TotalCost As Double = 0
    Public Property RealTotalCost As Double = 0
    Public Property Type As String = ""

    'For GA Class
    Public Property TaskID As Integer
    Public Property SequenceNumber As Integer

    'For loss calculation
    Public Property Capital As Boolean
    Public Property PatientsNoPerHour As Integer
    Public Property PatientCost As Double

    Public Property LossPerHour As Double

    'For DES/CBR
    Public Property TaskGroupTech As New List(Of Technician)
    Public Property GetRes As Boolean

End Class
