class Connection {
    Connection([string]$DSN, [string]$User, [string]$Password) {
        $this.DSN = $DSN
        $this.User = $User
        $this.Password = $Password
    }

    [string]$DSN
    [string]$User
    [string]$Password
}

<#
.SYNOPSIS
Creates an instance of the Connection class.

.PARAMETER DSN
The ODBC DSN used to connect.

.PARAMETER User
The database user for login.

.PARAMETER Password
The database password.

.EXAMPLE
New-Connection -DSN "mydsn" -User "dba" -Password "***"
#>
function New-Connection {
    param(
        [Parameter(Mandatory)]
        [string]$DSN,
        [Parameter(Mandatory)]
        [string]$User,
        [Parameter(Mandatory)]
        [string]$Password
    )
    [Connection]::new($DSN, $User, $Password)
}

Export-ModuleMember -Function New-Connection