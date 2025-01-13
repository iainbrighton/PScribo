$script:pscriboDocument = Document 'TestDocument' { }

$script:processes = @(
    [PSCustomObject]@{
        Name = "TestProcess1"
        Id = 1
        CPU = 10
        Memory = 100MB
    },
    [PSCustomObject]@{
        Name = "TestProcess2"
        Id = 2
        CPU = 20
        Memory = 200MB
    }
)

$script:services = @{
    Service1 = [PSCustomObject]@{
        Name = "TestService1"
        Status = "Running"
        StartType = "Automatic"
    }
    Service2 = [PSCustomObject]@{
        Name = "TestService2"
        Status = "Stopped"
        StartType = "Manual"
    }
}

Export-ModuleMember -Variable pscriboDocument, processes, services