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