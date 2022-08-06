
Imports Simphony.Simulation
Imports System.Globalization

Public Class Scenario

    Inherits DiscreteEventScenario

    Private ReadOnly MyEngine As New DiscreteEventEngine
    Private MyForm As Form1
    Private CaptureTechnicianResource As New Action(Of RepairTask)(AddressOf CaptureTechnician)
    Private TravelTechnicianEvent As New Action(Of RepairTask)(AddressOf TravelTechnician)
    Private RepairEvent As New Action(Of RepairTask)(AddressOf Repair)
    Private ReleaseTechnicianResource As New Action(Of RepairTask)(AddressOf ReleaseTechnician)

    Public Sub New(ByVal MyEngine As DiscreteEventEngine, ByVal InputForm As Form1)

        Me.MyEngine = MyEngine
        Me.MyForm = InputForm
        Alfa = MyForm.TBAlfa.Text
        Beta = MyForm.TBBeta.Text
        Gama = MyForm.TBGama.Text
    End Sub

    Public Overrides Function InitializeScenario() As Integer

        'StartPlan = DateTime.Parse(MyForm.MTBStartDate.Text).ToString("yyyy/MM/dd HH:mm", CultureInfo.GetCultureInfo("fa-Ir")) 'Read the Start that the user enter

        For Each tech As Technician In GroupTech

                Dim TempFile As New WaitingFile()
                TempFile.Name = tech.Name.Replace("Technician", "File")
                TempFile.IsBlocking = True
                tech.WaitingFiles.Add(TempFile)
                GroupFile.Add(TempFile)
            Next

        Return 1 'Number of runs
        'Throw New NotImplementedException()
    End Function

    Public Overrides Function InitializeRun(runIndex As Integer) As Double

        For Each tech In GroupTech
            tech.InitializeRun(runIndex)
        Next

        For Each file In GroupFile
            file.InitializeRun(runIndex)
        Next


        ' Define Entities
        ProjectTasks = ProjectTasks.OrderBy(Function(x) x.SequenceNumber).ToList
        For Each Newtask As RepairTask In ProjectTasks

            Newtask.TaskTech = GroupTech.Find(Function(p) p.TechName = Newtask.TechnicianName)
            Newtask.TaskTech.PreviousTask = Nothing
            MyEngine.ScheduleEvent(Newtask, CaptureTechnicianResource, 0)

        Next

        'Throw New NotImplementedException()

        Return 1000 'total simulation time
    End Function

    Private Sub CaptureTechnician(ByVal Newtask As RepairTask)

        MyEngine.RequestResource(Newtask, Newtask.TaskTech, Newtask.ResourceN, TravelTechnicianEvent, GroupFile.Find(Function(j) j.Name = Newtask.TaskTech.Name.Replace("Technician", "File")), 0)

    End Sub

    Private Sub TravelTechnician(ByVal Newtask As RepairTask)

        Dim TravelDuration As New Double
        Dim MyLoc As String

        Newtask.AssignedTime = StartPlan.AddMinutes(MyEngine.TimeNow)

        If IsNothing(Newtask.TaskTech.PreviousTask) Then
            MyLoc = Newtask.TaskTech.InitialTechLocation
        Else
            MyLoc = Newtask.TaskTech.PreviousTask.TaskName
        End If

        Dim newrow() As DataRow = tblDistances.Select("From='" & MyLoc & "' And To='" & Newtask.TaskName & "'")

        TravelDuration = newrow.First.Item("Distance") * 10 / (1 * 60)  'In Minutes

        MyEngine.ScheduleEvent(Newtask, RepairEvent, TravelDuration)

        Newtask.CostTravel = MyCostClass.CalCost(TravelDuration / 60, Newtask.TaskTech.UnitCost) 'Cost of travel should be corrected

    End Sub

    Private Sub Repair(NewTask As RepairTask)

        NewTask.StartTime = StartPlan.AddMinutes(MyEngine.TimeNow)

        NewTask.UsedSkill = NewTask.TaskTech.TechSkill

        If MyForm.CBProd.Checked Then
            If NewTask.RequiredSkill <> NewTask.UsedSkill Then
                NewTask.EsTime = NewTask.EsTime * (1 + MyForm.TBProdDec.Text / 100)
            End If
        End If

        MyEngine.ScheduleEvent(NewTask, ReleaseTechnicianResource, NewTask.EsTime)

        NewTask.CostTech = MyCostClass.CalCost(NewTask.EsTime / 60, NewTask.TaskTech.UnitCost)

    End Sub

    Private Sub ReleaseTechnician(NewTask As RepairTask)

        MyEngine.ReleaseResource(NewTask, NewTask.TaskTech, NewTask.ResourceN)

        NewTask.EndTime = StartPlan.AddMinutes(MyEngine.TimeNow)

        'CostDelay of PMs is 0, it may be different after defining a penalty function

        If NewTask.Type = "CM" Then
            If tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("CapitalOrNotCapital") = True Then
                NewTask.CostDelay = MyCostClass.CalCost((NewTask.EndTime - NewTask.ReportDate).TotalHours, tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("LossPerPatient") * tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("PatientNoPerHour") * tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("WorkingPercent") / 100)

            ElseIf tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("CapitalOrNotCapital") = False Then
                NewTask.CostDelay = MyCostClass.CalCost((NewTask.EndTime - NewTask.ReportDate).TotalHours, (tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("EquipmentPrice") / (tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("LifeExpectancy") * 12 * 30 * 24 * tblLossOfDamagedEquip.Rows.Find(NewTask.TaskName).Item("WorkingPercent") / 100)))
            End If

        ElseIf NewTask.Type = "PM" Then
            NewTask.CostDelay = 0
        End If

        NewTask.TotalCost = NewTask.Priority * (NewTask.CostTech * Beta + NewTask.CostTravel * Alfa + NewTask.CostDelay * Gama)

        'NewTask.TotalCost = NewTask.CostTech * Beta + NewTask.CostTravel * Alfa + NewTask.CostDelay * Gama
        NewTask.TaskTech.PreviousTask = NewTask
    End Sub

    Public Overrides Sub FinalizeRun(runIndex As Integer)
        AChromosomCost = 0
        For Each tech In GroupTech
            tech.FinalizeRun(runIndex, MyEngine.TimeNow)
        Next

        For Each file In GroupFile
            file.FinalizeRun(runIndex, MyEngine.TimeNow)
        Next

        For Each task As RepairTask In ProjectTasks

            AChromosomCost = AChromosomCost + task.TotalCost
        Next
        'Throw New NotImplementedException()
    End Sub

    Public Overrides Sub FinalizeScenario()

        'Throw New NotImplementedException()
    End Sub

End Class
