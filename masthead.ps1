Param(
  [Parameter(Mandatory = $true, Position = 1)]
  [string]$UserName,
  [Parameter(Mandatory = $true, Position = 2)]
  [SecureString]$Password,
  [Parameter(Mandatory = $true, Position = 3)]
  [string]$MastheadInstallerSite
)

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

Import-Module "$DistributionFolder\CreateSite\modules\Microsoft.SharePoint.Client.dll"

$ListTitle = "masthead-app-settings"
$caName = "Masthead";
$caTitle = "s-masthead-spx";
$caDescription = "Masthead for sharepoint";
$caLocation = "ClientSideExtension.ApplicationCustomizer";
$caComponentId = "27b0cb87-695b-4405-ae63-9db7d67e1029"

$caClassicName = "Masthead Classic";
$caClassicTitle = "masthead-classic";
$caClassicDescription = "Masthead classic action";
$caClassicLocation = "ScriptLink";
$caClassicSequence = 4884;
$caClassicScriptpt1 = 'function masthedClassicRetrieve() {var request = new XMLHttpRequest();request.open("GET","';
$caClassicScriptpt2 = '/_api/lists/getbytitle(''masthead-app-settings'')/items?$filter=Title eq ''classicScript''",true);request.onreadystatechange = function(){if (request.readyState === 4 && request.status === 200){var json = JSON.parse(request.response);var script = document.createElement("script");script.type = "text/javascript";script.src = json.value[0].URL + "masthead-classic.js";document.getElementsByTagName("body")[0].appendChild(script);var link = document.createElement("link");link.type = "text/css";link.rel = "stylesheet"; link.href= json.value[0].URL + "styles.css";document.getElementsByTagName("head")[0].appendChild(link);}}; request.withCredentials = true;request.setRequestHeader("Accept", "application/json");request.send();}masthedClassicRetrieve();masthedClassicRetrieve = null;'

Function Get-SPOCredentials([string]$UserName, [SecureString]$Password) {
  return New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $Password)
}

Function Get-List([Microsoft.SharePoint.Client.ClientContext]$Context, [String]$ListTitle) {
  $list = $context.Web.Lists.GetByTitle($ListTitle)
  $context.Load($list)
  $context.ExecuteQuery()
  $list
}

Function Get-Context-For-Site([string]$siteURL, [string]$UserName, [SecureString]$Password) {
  $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
  $context.Credentials = Get-SPOCredentials -UserName $UserName -Password $Password
  return $context
}

$adminContext = Get-Context-For-Site -siteURL $MastheadInstallerSite -UserName $UserName -Password $Password

Function Get-Masthead-Actions-From-Context([Microsoft.SharePoint.Client.ClientContext]$Context) {

  $siteActions = $Context.Site.UserCustomActions
  $webActions = $Context.Web.UserCustomActions

  $Context.Load($siteActions)
  $Context.Load($webActions)
  $Context.ExecuteQuery()

  $mastheadWeb = $webActions | Where-Object {$_.ClientSideComponentId -eq $caComponentId -or $_.Sequence -eq $caClassicSequence}
  $mastheadSite = $siteActions | Where-Object {$_.ClientSideComponentId -eq $caComponentId -or $_.Sequence -eq $caClassicSequence}

  if ($mastheadWeb -ne $null -And $mastheadSite -ne $null) {
    $mastheadWeb + $mastheadSite
  }
  if ($mastheadWeb -ne $null){
    $mastheadWeb
  }
  if ($mastheadSite -ne $null) {
    $mastheadSite
  }

}
Function Install-To-Site([string]$TargetSite) {

  $siteContext = Get-Context-For-Site -siteURL $TargetSite -UserName $UserName -Password $Password

  $siteWeb = $siteContext.Web
  $siteContext.Load($siteWeb)
  $siteContext.ExecuteQuery()
  $existingActions = $siteWeb.UserCustomActions
  $siteContext.Load($existingActions)
  $siteContext.ExecuteQuery()

  $caMasthead = $existingActions.Add()
  $caMasthead.Name = $caName
  $caMasthead.Title = $caTitle
  $caMasthead.Group = ""
  $caMasthead.Description = $caDescription
  $caMasthead.Location = $caLocation
  $caMasthead.ClientSideComponentId = $caComponentId
  $caMasthead.Update()

  $siteContext.ExecuteQuery()

  Try {
    $caClassic = $existingActions.Add()
    $caClassic.Name = $caClassicName
    $caClassic.Title = $caClassicTitle
    $caClassic.Group = ""
    $caClassic.Description = $caClassicDescription
    $caClassic.Location = $caClassicLocation
    $caClassic.Sequence = $caClassicSequence
    $caClassic.ScriptBlock = $caClassicScriptpt1 + $adminContext.Url + $caClassicScriptpt2
    $caClassic.Update()
    $siteContext.ExecuteQuery()
  }
  Catch {
    "Error adding classic version of masthead. If you are on a modern site, this is expected."
  }

  $settingsList = Get-List -Context $adminContext -ListTitle $ListTitle
  $adminContext.Load($settingsList)
  $adminContext.ExecuteQuery();

  $query = New-Object Microsoft.SharePoint.Client.CamlQuery
  $query.ViewXml = "<View>
    <RowLimit></RowLimit>
</View>"
  $listItems = $settingsList.GetItems($query)
  $adminContext.Load($listItems)
  $adminContext.ExecuteQuery();

  $instance = $listItems | Where-Object {$_["URL"] -eq $TargetSite.ToLower()}

  if ($instance -eq $null) {
    $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $mastheadURLItem = $settingsList.AddItem($listItemInfo)
    $mastheadURLItem["URL"] = $TargetSite.ToLower()
    $mastheadURLItem["Title"] = "masthead-url";
    $mastheadURLItem.Update()
    $settingsList.Update()
    try {
      $adminContext.ExecuteQuery()
    }
    catch {
      "Error adding to list"
    }
  }
  else {
    "Already on list, exiting..."
  }
}

Function Uninstall-From-Site([string]$TargetSite) {

  $siteContext = Get-Context-For-Site -siteURL $TargetSite -UserName $UserName -Password $Password

  $actions = Get-Masthead-Actions-From-Context -Context $siteContext

  Foreach ($action in $actions) {
    $action.DeleteObject()
  }
  $siteContext.Web.Update()
  $siteContext.ExecuteQuery()

  $settingsList = Get-List -Context $adminContext -ListTitle $ListTitle
  $adminContext.Load($settingsList)
  $adminContext.ExecuteQuery();

  $query = New-Object Microsoft.SharePoint.Client.CamlQuery
  $query.ViewXml = "<View>
    <RowLimit></RowLimit>
</View>"
  $listItems = $settingsList.GetItems($query)
  $adminContext.Load($listItems)
  $adminContext.ExecuteQuery();

  $instance = $listItems | Where-Object {$_["URL"] -eq $TargetSite.ToLower()}

  Foreach ($item in $instance) {
    $item.DeleteObject()
  }

  $adminContext.Web.Update()
  $adminContext.ExecuteQuery()
}

Function Install-To-Site-And-Subsites([string]$TargetSite) {
  Use-On-Subsites -TargetSite $TargetSite -RecursiveFunction ${function:Install-To-Site}
}

Function Uninstall-From-Site-And-Subsites([string]$TargetSite) {
  Use-On-Subsites -TargetSite $TargetSite -RecursiveFunction ${function:Uninstall-From-Site}
}

Function Use-On-Subsites([string]$TargetSite, [scriptblock]$RecursiveFunction) {
  $RecursiveFunction.Invoke($TargetSite)

  $siteContext = Get-Context-For-Site -siteURL $TargetSite -UserName $UserName -Password $Password

  $TargetSite -match 'https:\/\/(.*?\.sharepoint.com)'
  $originalDomain = $Matches[1]

  $web = $siteContext.Web
  $siteContext.Load($web)
  $siteContext.ExecuteQuery()
  $Webs = $siteContext.Web.Webs
  $siteContext.Load($Webs)
  $siteContext.ExecuteQuery()

  Foreach ($site in $Webs) {
    if  ($site.Url -match $originalDomain) {
      Use-On-Subsites -TargetSite $site.Url -RecursiveFunction $RecursiveFunction
    }
  }

}
