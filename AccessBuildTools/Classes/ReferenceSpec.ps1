class ReferenceSpec {
    <#
.SYNOPSIS
Creates an instance of the ReferenceSpec class.

.PARAMETER Name
The name of the reference.

.PARAMETER ApplyTheme
Specify, if the reference should have the database theme applied.

.EXAMPLE
New-ReferenceSpec -Name "myreference" -ApplyTheme $false
#>
    ReferenceSpec([string]$Name, [switch]$ApplyTheme) {
        $this.Name = $Name
        $this.ApplyTheme = $ApplyTheme
    }
    [string]$Name
    [bool]$ApplyTheme
}

function New-ReferenceSpec {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [switch]$ApplyTheme
    )
    [ReferenceSpec]::new($Name, $ApplyTheme)
}

Export-ModuleMember -Function New-ReferenceSpec