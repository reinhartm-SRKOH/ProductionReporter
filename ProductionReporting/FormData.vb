Public Class FormData

    ' Used by all reporters.
    Private associateValue As String
    Private prodLineValue As String
    Private partNumberValue As String

    ' Used by production reporter.
    Private cavityCountValue As String
    Private shotsValue As String
    Private freeShotsValue As String
    Private scrappedPiecesValue As String
    Private productionQtyValue As String

    ' Used by downtime reporter.
    Private downtimeCodeValue As String
    Private downtimeReasonValue As String
    Private downtimeQtyValue As String

    ' Used by scrap reporter.
    Private scrapCodeValue As String
    Private scrapReasonValue As String
    Private scrapQtyValue As String


    ' Used only when reporting, not on the form.
    Private shiftValue As String
    Private runDateValue As String
    Private runtimeValue As String




    Public Property Associate() As String
        Get
            Return associateValue
        End Get
        Set(ByVal value As String)
            associateValue = value
        End Set
    End Property

    Public Property ProdLine() As String
        Get
            Return prodLineValue
        End Get
        Set(ByVal value As String)
            prodLineValue = value
        End Set
    End Property

    Public Property PartNumber() As String
        Get
            Return partNumberValue
        End Get
        Set(ByVal value As String)
            partNumberValue = value
        End Set
    End Property



    Public Property CavityCount() As String
        Get
            Return cavityCountValue
        End Get
        Set(value As String)
            cavityCountValue = value
        End Set
    End Property

    Public Property Shots() As String
        Get
            Return shotsValue
        End Get
        Set(value As String)
            shotsValue = value
        End Set
    End Property

    Public Property FreeShots() As String
        Get
            Return freeShotsValue
        End Get
        Set(value As String)
            freeShotsValue = value
        End Set
    End Property

    Public Property ScrappedPieces() As String
        Get
            Return scrappedPiecesValue
        End Get
        Set(value As String)
            scrappedPiecesValue = value
        End Set
    End Property

    Public Property ProductionQty() As String
        Get
            Return productionQtyValue
        End Get
        Set(ByVal value As String)
            productionQtyValue = value
        End Set
    End Property




    Public Property DowntimeCode() As String
        Get
            Return downtimeCodeValue
        End Get
        Set(ByVal value As String)
            downtimeCodeValue = value
        End Set
    End Property

    Public Property DowntimeReason() As String
        Get
            Return downtimeReasonValue
        End Get
        Set(value As String)
            downtimeReasonValue = value
        End Set
    End Property

    Public Property DowntimeQty() As String
        Get
            Return downtimeQtyValue
        End Get
        Set(ByVal value As String)
            downtimeQtyValue = value
        End Set
    End Property




    Public Property ScrapCode() As String
        Get
            Return scrapCodeValue
        End Get
        Set(ByVal value As String)
            scrapCodeValue = value
        End Set
    End Property

    Public Property ScrapReason() As String
        Get
            Return scrapReasonValue
        End Get
        Set(value As String)
            scrapReasonValue = value
        End Set
    End Property

    Public Property ScrapQty() As String
        Get
            Return scrapQtyValue
        End Get
        Set(ByVal value As String)
            scrapQtyValue = value
        End Set
    End Property




    Public Property Shift() As String
        Get
            Return shiftValue
        End Get
        Set(ByVal value As String)
            shiftValue = value
        End Set
    End Property

    Public Property RunDate() As String
        Get
            Return runDateValue
        End Get
        Set(ByVal value As String)
            runDateValue = value
        End Set
    End Property

    Public Property Runtime() As String
        Get
            Return runtimeValue
        End Get
        Set(ByVal value As String)
            runtimeValue = value
        End Set
    End Property

End Class
