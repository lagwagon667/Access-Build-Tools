class ReferenceSpec {
    ReferenceSpec([string]$Name, [bool]$ApplyTheme) {
        $this.Name = $Name
        $this.ApplyTheme = $ApplyTheme
    }
    [string]$Name
    [bool]$ApplyTheme
}