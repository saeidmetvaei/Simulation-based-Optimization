Imports System.Data.OleDb
Imports System.Data.Common
Imports Simphony.Simulation
Imports System.Text
Imports System.Net
Imports Newtonsoft.Json

Public Class GA

    Dim myform As Form1
    Private MyEngine As New DiscreteEventEngine

    Public Sub New(ByVal Myform As Form1)
        Me.myform = Myform
        GenerationNumber = 0
        GenerationTargetCount = Myform.TBGenerationN.Text 'the number of generation
        Population = Myform.TBPopulation.Text ' the count of chromosoms in one generation
        EliteSelectionPercent = Myform.TBElite.Text 'the percent of Elite
        MutationProbability = Myform.TBMutation.Text 'the probability of mutation
        MaxStopCounter = Myform.TBStopG.Text 'the max repetation of best total cost


        PrepareAccess()
        Generation.Clear()

        InitialGeneration()
        AnalyzeGeneration(Generation)

        Do While GenerationNumber < GenerationTargetCount And StopCounter < MaxStopCounter


            CrossOver()
            AnalyzeGeneration(Generation)


            ' Sorting Chromosomes by their total cost Beacuse we are going to select bests of them
            Generation = Generation.OrderBy(Function(x) x.total_cost).ToList

            'To Check if best chromosom is improving or not
            If Generation.Item(0).total_cost = LastBestCost Then
                StopCounter = StopCounter + 1
            Else
                LastBestCost = Generation.Item(0).total_cost
                StopCounter = 0
            End If
        Loop

        BestChromosome = Generation.Item(0)

        ExtractBestChromosomeData(BestChromosome)

        ' show result of Best Chromosome

        BestProject = BestProject.OrderBy(Function(x) x.SequenceNumber).ToList

        Myform.CostTxt.Text = BestChromosome.total_cost 'CalculateFitness(BestProject)

        UpdateAccess(GAOut)

    End Sub


    Private Sub InitialGeneration()

        GenerationNumber = GenerationNumber + 1

        For i = 1 To Population

            Dim MyChromosom As New Chromosome

            MyChromosom.ID = i
            MyChromosom.MadeBy = "Initial Generation"

            For j = 1 To GenomesCount
                Randomize()
                MyChromosom.total_genomes.Add(j, Rnd())
            Next

            Generation.Add(MyChromosom)

        Next

    End Sub

    Public Sub AnalyzeGeneration(ByVal MyGeneration As List(Of Chromosome))

        For Each MyChromose In MyGeneration

            MyChromose.Resource_genomes.Clear()
            MyChromose.Sequencey_genomes.Clear()

            ' extract resource genomes & sequence genomes
            For i = 1 To TasksCount
                MyChromose.Resource_genomes.Add(i, MyChromose.total_genomes(i))
            Next

            For i = 1 To TasksCount
                MyChromose.Sequencey_genomes.Add(i, MyChromose.total_genomes(TasksCount + i))
            Next

            ProjectTasks.Clear()

            ' To find out who is going to do each tasks
            For GenemoeNumber = 1 To MyChromose.Resource_genomes.Count
                Dim MyTask As New RepairTask
                MyTask.TaskID = GenemoeNumber

                ' other information of a task
                MyTask.TaskName = tblTasks.Rows(GenemoeNumber - 1).Item("EquipmentCode")
                MyTask.TaskDescription = tblTasks.Rows(GenemoeNumber - 1).Item("TaskDescription")
                MyTask.Status = tblTasks.Rows(GenemoeNumber - 1).Item("Status")
                MyTask.ReportDate = tblTasks.Rows(GenemoeNumber - 1).Item("ReportDate")
                MyTask.ResourceN = tblTasks.Rows(GenemoeNumber - 1).Item("ResourceTechN")
                MyTask.Priority = tblTasks.Rows(GenemoeNumber - 1).Item("Priority")
                MyTask.EsTime = tblTasks.Rows(GenemoeNumber - 1).Item("EstimatedRepairTime")
                MyTask.RequiredSkill = tblTasks.Rows(GenemoeNumber - 1).Item("ResourceSkill")
                MyTask.MustSkilled = tblTasks.Rows(GenemoeNumber - 1).Item("MustSkilled")
                MyTask.Type = tblTasks.Rows(GenemoeNumber - 1).Item("Type")
                'MyTask.Capital = tblLossOfDamagedEquip.Rows.Find(MyTask.TaskName).Item("CapitalOrNotCapital")
                'MyTask.PatientsNoPerHour = tblLossOfDamagedEquip.Rows.Find(MyTask.TaskName).Item("PatientsNoPerHour")
                'MyTask.PatientCost = tblLossOfDamagedEquip.Rows.Find(MyTask.TaskName).Item("LossPerPatient")

                'Resource of a task

                For i = 0 To ResourcesCount - 1

                    Select Case MyChromose.Resource_genomes.Item(GenemoeNumber)
                        Case i / ResourcesCount To (i + 1) / ResourcesCount
                            MyTask.TechnicianName = tblTechnicians.Rows(i).Item("TechnicianName")
                    End Select

                Next

                If myform.CBSkill.Checked Then

                    If MyTask.MustSkilled = True Then
                        Dim FilteredtblTechnicians As New DataTable
                        FilteredtblTechnicians = tblTechnicians.Select("TechnicianSkill='" & MyTask.RequiredSkill & "'").CopyToDataTable()
                        Dim FilteredResourcesCount As Integer = 0
                        FilteredResourcesCount = FilteredtblTechnicians.Rows.Count

                        For i = 0 To FilteredResourcesCount - 1

                            Select Case MyChromose.Resource_genomes.Item(GenemoeNumber)
                                Case i / FilteredResourcesCount To (i + 1) / FilteredResourcesCount
                                    MyTask.TechnicianName = FilteredtblTechnicians.Rows(i).Item("TechnicianName")
                            End Select

                        Next
                    End If
                End If

                    ProjectTasks.Add(MyTask)

            Next

            ' to extract the sequence of each task
            Dim mylist As New List(Of Double)

            For i = 1 To MyChromose.Sequencey_genomes.Count
                mylist.Add(MyChromose.Sequencey_genomes.Item(i))
            Next

            mylist.Sort()

            For GenomeNumber = 1 To MyChromose.Sequencey_genomes.Count
                'Dim mytask As New RepairTask
                ProjectTasks.Item(GenomeNumber - 1).SequenceNumber = mylist.IndexOf(MyChromose.Sequencey_genomes.Item(GenomeNumber))

            Next

            ' Result of Simulation for each chromosome

            MyChromose.total_cost = CalculateFitness(ProjectTasks)


            ' To Save Project Details on DB
            SaveOnDataTable(ProjectTasks, MyChromose)

        Next

    End Sub


    Private Sub CrossOver()

        Dim Parent1, Parent2 As Chromosome
        Dim Child1, Child2 As Chromosome
        Dim Parents As New List(Of Chromosome)
        Dim CrossPos As New List(Of Integer)


        For Each myChromosome As Chromosome In Generation
            Parents.Add(myChromosome)
        Next


        ' Sorting Chromosomes by their total cost Beacuse we are going to select bests of them
        Generation = Generation.OrderBy(Function(x) x.total_cost).ToList

        ' Choosing best Chromosomes 
        Dim EliteChromosomesList As New List(Of Chromosome)
        Dim EliteChromosomesCount As Integer = Int(EliteSelectionPercent * Population)

        For i = 0 To EliteChromosomesCount - 1
            EliteChromosomesList.Add(Generation.Item(i))
        Next


        ' Strat for creating new generation
        Generation.Clear()
        GenerationNumber = GenerationNumber + 1
        Dim GeneomID As Integer = 0


        ' Adding Elite Chromosomes to next generation
        For Each MyChromosome As Chromosome In EliteChromosomesList

            GeneomID = GeneomID + 1

            MyChromosome.ID = GeneomID
            MyChromosome.MadeBy = "Elite"

            Generation.Add(MyChromosome)

        Next

        ' Doing CrossOver and creating the rest of chromosomes
        Do While Generation.Count < Population

            Child1 = New Chromosome
            Child2 = New Chromosome

            Randomize()
            Parent1 = Parents.Item(Int(Rnd() * (Parents.Count)))
            Parents.Remove(Parent1)
            Randomize()
            Parent2 = Parents.Item(Int(Rnd() * Parents.Count))
            Parents.Remove(Parent2)

            For j = 1 To 3
                Randomize()
                CrossPos.Add(Int(Rnd() * (GenomesCount)))
            Next
            CrossPos.Sort()

            For GeneID = 1 To CrossPos.Item(0)
                Child1.total_genomes.Add(GeneID, Parent1.total_genomes(GeneID))
                Child2.total_genomes.Add(GeneID, Parent2.total_genomes(GeneID))
            Next

            For GeneID = CrossPos.Item(0) + 1 To CrossPos.Item(1)
                Child1.total_genomes.Add(GeneID, Parent2.total_genomes(GeneID))
                Child2.total_genomes.Add(GeneID, Parent1.total_genomes(GeneID))
            Next

            For GeneID = CrossPos.Item(1) + 1 To CrossPos.Item(2)
                Child1.total_genomes.Add(GeneID, Parent1.total_genomes(GeneID))
                Child2.total_genomes.Add(GeneID, Parent2.total_genomes(GeneID))
            Next

            For GeneID = CrossPos.Item(2) + 1 To GenomesCount
                Child1.total_genomes.Add(GeneID, Parent2.total_genomes(GeneID))
                Child2.total_genomes.Add(GeneID, Parent1.total_genomes(GeneID))
            Next

            GeneomID = GeneomID + 1
            Child1.MadeBy = "CrossOver"
            Child1.ID = GeneomID

            GeneomID = GeneomID + 1
            Child2.MadeBy = "CrossOver"
            Child2.ID = GeneomID

            Mutation(Child1)
            Mutation(Child2)

            Generation.Add(Child1)
            Generation.Add(Child2)

        Loop

    End Sub

    Private Sub Mutation(ByVal MyChild As Chromosome)

        Dim GenomeMutationProbability As Double

        For genome = 1 To MyChild.total_genomes.Count

            Randomize()
            GenomeMutationProbability = Rnd()

            If GenomeMutationProbability <= MutationProbability Then
                Randomize()
                MyChild.total_genomes.Item(genome) = Rnd()
                MyChild.MadeBy = "Mutation"
            End If

        Next

    End Sub


    Private Function CalculateFitness(projectTasks As List(Of RepairTask))


        Dim TotalCost As Double = 0

        Dim HospitalModel As New Model(MyEngine, myform)
        MyEngine.InitializeEngine()
        MyEngine.Simulate(HospitalModel)
        myform.TBTimeSim.Text = MyEngine.TimeNow
        TotalCost = AChromosomCost

        Return TotalCost
    End Function


    Private Sub PrepareAccess()

        Dim adapter1 As OleDbDataAdapter
        Dim adapter5 As OleDbDataAdapter

        mydataset.Tables.Clear()

        'to read tasks Table
        tblTasksPM = MyReadInfo.UncompleteTasksPM(myform)
        tblTasksCM = MyReadInfo.UncompleteTasksCM()



        'Solve GA for CM, PM or Both
        tblTasks.Clear()
        If myform.CBPM.Checked = True And myform.CBCM.Checked = True Then
            tblTasks = tblTasksCM
            For i = 0 To tblTasksPM.Rows.Count - 1
                Dim newrow = tblTasks.NewRow()
                newrow("EquipmentCode") = tblTasksPM.Rows(i).Item("EquipmentCode")
                newrow("TaskDescription") = tblTasksPM.Rows(i).Item("TaskDescription")
                newrow("Status") = tblTasksPM.Rows(i).Item("Status")
                newrow("ReportDate") = tblTasksPM.Rows(i).Item("ReportDate")
                newrow("ResourceTechN") = tblTasksPM.Rows(i).Item("ResourceTechN")
                newrow("Priority") = tblTasksPM.Rows(i).Item("Priority")
                newrow("EstimatedRepairTime") = tblTasksPM.Rows(i).Item("EstimatedRepairTime")
                newrow("ResourceSkill") = tblTasksPM.Rows(i).Item("ResourceSkill")
                newrow("MustSkilled") = tblTasksPM.Rows(i).Item("MustSkilled")
                newrow("Type") = tblTasksPM.Rows(i).Item("Type")

                tblTasks.Rows.Add(newrow)
            Next

        ElseIf myform.CBPM.Checked = False And myform.CBCM.Checked = True Then
            tblTasks = tblTasksCM
        ElseIf myform.CBPM.Checked = True And myform.CBCM.Checked = False Then
            tblTasks = tblTasksPM
        ElseIf myform.CBPM.Checked = False And myform.CBCM.Checked = False Then
            MsgBox("Please Select CM, PM, or both to satrt simulation")
            'Exit Function??????
        End If

        TasksCount = tblTasks.Rows.Count
        GenomesCount = 2 * TasksCount

        'to read distances Table
        tblDistances = MyReadInfo.Distances()

        ' to read Resources table
        tblTechnicians = MyReadInfo.Technicians()

        ResourcesCount = tblTechnicians.Rows.Count

        ' to read Loss of production table
        tblLossOfDamagedEquip = MyReadInfo.LossOfDamagedEquip()


        Try
            MyAccessConn = New OleDbConnection(connectionstring)
            MyAccessConn.Open()
            'to drop GA_Out Table
            adapter1 = New OleDbDataAdapter("DROP TABLE GA_Out;", MyAccessConn)
            adapter1.Fill(mydataset, "GA_Out")

            'to creat GA_Out Table
            Dim SQlScript As String = "CREATE TABLE GA_Out (Project_Number  INTEGER PRIMARY KEY"
            For i = 0 To TasksCount - 1
                SQlScript = SQlScript & "," & "Task" & i + 1 & "_Resource  CHAR (20)"
            Next
            For i = 0 To TasksCount - 1
                SQlScript = SQlScript & "," & "Task" & i + 1 & "_Sequence INTEGER"
            Next
            SQlScript = SQlScript & ", Project_Cost DOUBLE, Made_By CHAR (20));"

            adapter5 = New OleDbDataAdapter(SQlScript, MyAccessConn)
            adapter5.Fill(GAOut)

            Dim adapter As OleDbDataAdapter
            adapter = New OleDbDataAdapter("SELECT GA_Out.* FROM GA_Out;", MyAccessConn)
            adapter.Fill(GAOut)

            MyAccessConn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub


    Private Sub ExtractBestChromosomeData(ByVal MyChromose As Chromosome)

        ProjectTasks.Clear()

        ' To find out who is going to do each tasks
        For GenemoeNumber = 1 To MyChromose.Resource_genomes.Count
            Dim MyTask As New RepairTask
            MyTask.TaskID = GenemoeNumber
            ' other information of a task
            MyTask.TaskName = tblTasks.Rows(GenemoeNumber - 1).Item("EquipmentCode")
            MyTask.TaskDescription = tblTasks.Rows(GenemoeNumber - 1).Item("TaskDescription")
            MyTask.Status = tblTasks.Rows(GenemoeNumber - 1).Item("Status")
            MyTask.ReportDate = tblTasks.Rows(GenemoeNumber - 1).Item("ReportDate")
            MyTask.ResourceN = tblTasks.Rows(GenemoeNumber - 1).Item("ResourceTechN")
            MyTask.Priority = tblTasks.Rows(GenemoeNumber - 1).Item("Priority")
            MyTask.EsTime = tblTasks.Rows(GenemoeNumber - 1).Item("EstimatedRepairTime")
            MyTask.RequiredSkill = tblTasks.Rows(GenemoeNumber - 1).Item("ResourceSkill")
            MyTask.MustSkilled = tblTasks.Rows(GenemoeNumber - 1).Item("MustSkilled")
            MyTask.Type = tblTasks.Rows(GenemoeNumber - 1).Item("Type")

            For i = 0 To ResourcesCount - 1

                Select Case MyChromose.Resource_genomes.Item(GenemoeNumber)
                    Case i / ResourcesCount To (i + 1) / ResourcesCount
                        MyTask.TechnicianName = tblTechnicians.Rows(i).Item("TechnicianName")
                End Select

            Next

            If myform.CBSkill.Checked Then

                If MyTask.MustSkilled = True Then
                    Dim FilteredtblTechnicians As New DataTable
                    FilteredtblTechnicians = tblTechnicians.Select("TechnicianSkill='" & MyTask.RequiredSkill & "'").CopyToDataTable()
                    Dim FilteredResourcesCount As Integer = 0
                    FilteredResourcesCount = FilteredtblTechnicians.Rows.Count

                    For i = 0 To FilteredResourcesCount - 1

                        Select Case MyChromose.Resource_genomes.Item(GenemoeNumber)
                            Case i / FilteredResourcesCount To (i + 1) / FilteredResourcesCount
                                MyTask.TechnicianName = FilteredtblTechnicians.Rows(i).Item("TechnicianName")
                        End Select

                    Next
                End If
            End If

                ProjectTasks.Add(MyTask)
        Next

        ' to extract the sequence of each task
        Dim mylist As New List(Of Double)

        For i = 1 To MyChromose.Sequencey_genomes.Count
            mylist.Add(MyChromose.Sequencey_genomes.Item(i))
        Next

        mylist.Sort()

        For GenomeNumber = 1 To MyChromose.Sequencey_genomes.Count
            ProjectTasks.Item(GenomeNumber - 1).SequenceNumber = mylist.IndexOf(MyChromose.Sequencey_genomes.Item(GenomeNumber))
        Next





        Dim HospitalModel As New Model(MyEngine, myform)
        MyEngine.InitializeEngine()
        MyEngine.Simulate(HospitalModel)
        myform.TBTimeSim.Text = MyEngine.TimeNow
        'TotalCost = AChromosomCost

        'Connect to the Host
        Dim webClient As New WebClient()

        Try

            webClient.Headers("content-type") = "application/json"

            webClient.OpenRead("http://cmms-service.ir/MrHosSim/FormatTable.php")

            webClient.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        For Each task As RepairTask In ProjectTasks
            Dim s As String = JsonConvert.SerializeObject(task)

            'jsonString.Text = s


            'Dim webClient As New WebClient()
            Dim resByte As Byte()
            Dim resString As String
            Dim reqString() As Byte

            Try
                webClient.Headers("content-type") = "application/json"
                reqString = Encoding.Default.GetBytes(s)
                resByte = webClient.UploadData("http://cmms-service.ir/MrHosSim/FromVBToSimOut.php", "post", reqString)
                resString = Encoding.Default.GetString(resByte)

                'jsonString.Text = resString

                webClient.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next

        Try
            MyAccessConn = New OleDbConnection(ConnectionString)
            MyAccessConn.Open()

            Dim AdapterSchedule As New OleDbDataAdapter("SELECT SimOut.* FROM SimOut;", MyAccessConn)
            Dim AdapterDelete As New OleDbCommand("Delete SimOut.* FROM SimOut;", MyAccessConn)
            AdapterDelete.ExecuteNonQuery()

            Dim CID As New DataColumnMapping("TaskID", "TaskID")
            Dim CEquipmentCode As New DataColumnMapping("EquipmentCode", "EquipmentCode")
            Dim CTaskDescription As New DataColumnMapping("TaskDescription", "TaskDescription")
            Dim CStatus As New DataColumnMapping("Status", "Status")
            Dim CPriority As New DataColumnMapping("Priority", "Priority")
            Dim CPlannedStartT As New DataColumnMapping("PlannedStartTime", "PlannedStartTime")
            Dim CPlannedFinishT As New DataColumnMapping("PlannedFinishTime", "PlannedFinishTime")
            Dim CTechnecianQ As New DataColumnMapping("TechnecianQ", "TechnecianQ")
            Dim CTechName As New DataColumnMapping("TechName", "TechName")
            Dim CUsedSkill As New DataColumnMapping("UsedSkill", "UsedSkill")
            Dim CRequiredSkill As New DataColumnMapping("RequiredSkill", "RequiredSkill")
            Dim CTotalCost As New DataColumnMapping("TotalCost", "TotalCost")
            Dim CPercentComplete As New DataColumnMapping("PercentComplete", "PercentComplete")
            Dim CSequence As New DataColumnMapping("Sequence", "Sequence")
            Dim CAssignedTime As New DataColumnMapping("AssignedTime", "AssignedTime")
            Dim CRealTotalCost As New DataColumnMapping("RealTotalCost", "RealTotalCost")

            Dim dt As New DataTableMapping("Table", "SimResult")

            dt.ColumnMappings.Add(CID)
            dt.ColumnMappings.Add(CEquipmentCode)
            dt.ColumnMappings.Add(CPriority)
            dt.ColumnMappings.Add(CPlannedStartT)
            dt.ColumnMappings.Add(CPlannedFinishT)
            dt.ColumnMappings.Add(CTechnecianQ)
            dt.ColumnMappings.Add(CTechName)
            dt.ColumnMappings.Add(CUsedSkill)
            dt.ColumnMappings.Add(CRequiredSkill)
            dt.ColumnMappings.Add(CTotalCost)
            dt.ColumnMappings.Add(CPercentComplete)
            dt.ColumnMappings.Add(CSequence)
            dt.ColumnMappings.Add(CAssignedTime)
            dt.ColumnMappings.Add(CTaskDescription)
            dt.ColumnMappings.Add(CStatus)
            dt.ColumnMappings.Add(CRealTotalCost)

            AdapterSchedule.TableMappings.Add(dt)

            AdapterSchedule.Fill(SimResult)
            SimResult.Clear()
            myform.DataGridView1.DataSource = SimResult

            Dim cb As New OleDbCommandBuilder(AdapterSchedule)

            'For Each item In Tasks
            For Each task As RepairTask In ProjectTasks
                Dim taskschedule = SimResult.NewRow()

                taskschedule("TaskID") = task.TaskID
                taskschedule("EquipmentCode") = task.TaskName
                taskschedule("TaskDescription") = task.TaskDescription
                taskschedule("Status") = task.Status
                taskschedule("Priority") = task.Priority
                taskschedule("AssignedTime") = task.AssignedTime
                taskschedule("PlannedStartTime") = task.StartTime
                taskschedule("PlannedFinishTime") = task.EndTime
                taskschedule("TechnecianQ") = task.ResourceN
                taskschedule("TechName") = task.TechnicianName
                taskschedule("RequiredSkill") = task.RequiredSkill
                taskschedule("UsedSkill") = task.UsedSkill
                taskschedule("TotalCost") = task.TotalCost
                taskschedule("Sequence") = task.SequenceNumber
                taskschedule("RealTotalCost") = task.CostTech + task.CostTravel + task.CostDelay
                SimResult.Rows.Add(taskschedule)

            Next

            AdapterSchedule.Update(SimResult.GetChanges())
            SimResult.AcceptChanges()

            MyAccessConn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        'BestProject = ProjectTasks

    End Sub


    Dim ProjectCounter As Integer = 0

    Private Sub SaveOnDataTable(ByVal MyProject As List(Of RepairTask), ByVal MyChromosome As Chromosome)

        ProjectCounter = ProjectCounter + 1


        Try

            '    MyAccessConn.Open()
            'Dim adapter As OleDbDataAdapter
            'adapter = New OleDbDataAdapter("SELECT GA_Out.* FROM GA_Out;", MyAccessConn)
            'adapter.Fill(mydataset, "GA_Out")

            Dim myrow As DataRow

            myrow = GAOut.NewRow

            myrow.Item("Project_Number") = ProjectCounter

        For i = 0 To TasksCount - 1
                myrow.Item("Task" & i + 1 & "_Resource") = MyProject.Item(i).TechnicianName
            Next

            For i = 0 To TasksCount - 1
                Dim j = i + 1
                myrow.Item("Task" & j & "_Sequence") = MyProject.IndexOf(MyProject.Find(Function(p) p.Name = "RepairTask" & j))
            Next

            myrow.Item("Project_Cost") = MyChromosome.total_cost
        myrow.Item("Made_By") = MyChromosome.MadeBy

            GAOut.Rows.Add(myrow)



            '    Dim da As New OleDbCommandBuilder(adapter)

            'adapter.Update(mydataset.Tables("GA_Out"))

            '    MyAccessConn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub


    Private Sub UpdateAccess(MyDataTable As DataTable)

        Try

            MyAccessConn.Open()
            Dim adapter As OleDbDataAdapter
            adapter = New OleDbDataAdapter("SELECT GA_Out.* FROM GA_Out;", MyAccessConn)
            'adapter.Fill(mydataset, "GA_Out")
            Dim da As New OleDbCommandBuilder(adapter)

            adapter.Update(GAOut)

            MyAccessConn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub


End Class


