$accountName=$args[0]

if($args.count -lt 1){
 "Usage: ./trustForDelegation.ps1 <accountname>"
}

$TRUSTED_FOR_DELEGATION = 524288;

$gc="GC://" + $([adsi] "LDAP://RootDSE").Get("RootDomainNamingContext")

$filter = "(cn=$accountName)"
$domain = New-Object System.DirectoryServices.DirectoryEntry($gc)
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.SearchRoot = $domain
$searcher.Filter = $filter
$results = $searcher.FindAll()
if($results.count -eq 0){ "User Not Found"; }else{
 foreach ($result in $results){
  $dn=[string]$($result.properties["adspath"]).replace("GC://","LDAP://")
  $account=New-Object System.DirectoryServices.DirectoryEntry($dn)
  "Trusting $($account.cn) for Delegation..."
  $uac=$account.userAccountControl[0] -bor $TRUSTED_FOR_DELEGATION
  $account.userAccountControl[0]=$uac
  $result=$account.CommitChanges()
 }
}
