Public Class Cost

    Public Function CalCost(ByVal MyTime As Double, ByVal MyUnitCost As Double) As Double
        Dim MyCost As Double

        MyCost = MyTime * MyUnitCost

        Return MyCost
    End Function

End Class
