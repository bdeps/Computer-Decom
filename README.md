# Computer-Decom
Script to decommission a computer (disable, move, append a description)

# Usage
    .Required Parameters
    
    .Parameter $Identity
    Takes computerGUIDS. Either a list or a single guid

    .Parameter $computerDomain
    Takes the domain the computers are in. All computers in the list needs to be from the same domain

    .Parameter $newDescription
    Takes the a string that contains an addition or replacement to the current description

    .Parameter $targetOUPath
    Takes an OU path where the computer objects should be moved to "OU=DisabledComputers, DC=Contoso, DC=com"

    .Parameter $appendToCurrentDescription
    TRUE if you want to append $newDescription to the end of the existing one. False if you want to completely replace the current description

    .Not Required Parameters
    
    .Parameter $orderID
    OrderID if running from an automatic service order. Automatically appended at the end of the description

    .Parameter $orderDate
    OrderDate if running from an automatic service order order. Automatically appended at the end of the description

# Examples
    .EXAMPLE
    Decom-Computers -Identity "ExampleGUID1,ExampleGUID2" -computerDomain "contoso.com" -newDescription "Test Description" -appendToCurrentDescription $true -orderID "123"
    
 # Planned Improvements
    Add options to clear AD-Groups
    Add options to append the OU-Path before the move (save in an attribute or description?)
    
