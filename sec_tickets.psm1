class SecTicket
{
    [String] $ticketnumber
    [String] $pLevel
    [String] $category
    [String] $compromise
    [String] $Activitydate
    [String] $DDateTime
    [String] $NDateTime
    [String] $FireEmail
    [string] $source
    [string] $name
    [string] $filepath
    [string] $file
    [string] $signature
    [string] $hostname

 SecTicket ([String]$ticketnumber, [String]$pLevel, [string]$category, [string]$compromise, [string]$Activitydate, [string]$DDateTime, [string]$NDateTime, [string]$FireEmail, [string]$source, [string]$name, [string]$filepath, [string]$file, [string]$signature, [string]$hostname)
    {
        $this.ticketnumber = $ticketnumber
        $this.pLevel = $pLevel
        $this.category = $category
        $this.compromise = $compromise
        $this.Activitydate = $Activitydate
        $this.DDateTime = $DDateTime
        $this.NDateTime = $NDateTime
        $this.FireEmail = $FireEmail
        $this.source = $source
        $this.name = $name
        $this.filepath = $filepath
        $this.file = $file
        $this.signature = $signature
        $this.hostname = $hostname
        }
}
