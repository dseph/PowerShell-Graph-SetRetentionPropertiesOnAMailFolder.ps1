# PowerShell-Graph-SetRetentionPropertiesOnAMailFolder.ps1
<#
# By Daniel Bagley, Microsoft Ltd. 2026. Use at your own risk.  No warranties are given.
# Thanks to Chris Pollit for helping with teh adnim side.

 DISCLAIMER:
THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
 
============================================================
 
Extended Properties used:
  PR_POLICY_TAG       0x3019 Binary
  PR_RETENTION_FLAGS  0x301D Integer
  PR_RETENTION_PERIOD 0x301A Integer

Summary:
  PATCH /users/{user}/mailFolders/{folderId}
  singleValueExtendedProperties:
      "Binary 0x3019"
      "Integer 0x301D"
      "Integer 0x301A"

.NOTES:
    This is not a documented/supported Purview retention-label
    folder API. It is simply the Graph extended-property mapping  of the MAPI properties used by an EWS sample.
    This sample, instructions and information has not been throughly tested and is based upon articles and samples
    around EWS and PowerShell setting the retention properties on mail folders. Be sure to modify and test througly
    and making this code your own before using in production.

PERMISSIONS:
    Mail.ReadWrite - Application permission. Don't forget to do an admin grant.

If you want to experiment with the same MAPI retention properties that EWS used (PR_POLICY_TAG, PR_RETENTION_FLAGS, PR_RETENTION_PERIOD), Graph 
supports extended properties on some mailbox objects. However, Microsoft does not document a supported mailbox-folder retention-label API. The 
classic EWS approach stamps these properties directly on the folder.

For Exchange Online in 2026:

Use Purview retention labels for item-level retention where possible.
Use Retention Policies assigned at the mailbox level when folder-level targeting is not required.
If you need true folder-level MRM tag stamping, EWS remains the historical method, though you should evaluate Microsoft's EWS retirement guidance.
https://learn.microsoft.com/en-us/purview/retention?tabs=table-overriden

===============================
Before you try this script.

Before this script can be used, there needs to hav been a retention policy assigned to the mailbox and before that a retention policy
needs to have bene created.

# First, connect to Exchange Online.
Connect-ExchangeOnline
# Create a Retention Tag - if you need one:
 
New-RetentionPolicyTag `
    -Name "Delete After 3 Years" `
    -Type Personal `
    -RetentionEnabled $true `
    -AgeLimitForRetention 1095 `
    -RetentionAction PermanentlyDelete
    
# Create a retention policy:
Connect-ExchangeOnline
New-RetentionPolicy `
    -Name "Corporate Retention Policy" `
    -RetentionPolicyTagLinks "Delete After 3 Years"

# Apply the Policy to a Mailbox
Set-Mailbox `
    -Identity "user@contoso.com" `
    -RetentionPolicy "Corporate Retention Policy"

# Verify the Assignment
Get-Mailbox user@contoso.com 
Select-Object DisplayName,RetentionPolicy

# Show All Tags in the Assigned Policy
$Mailbox = Get-Mailbox user@contoso.com
$Policy = Get-RetentionPolicy $Mailbox.RetentionPolicy
$Policy.RetentionPolicyTagLinks |
    ForEach-Object {
        Get-RetentionPolicyTag $_
    } |
    Format-Table `
        Name,
        Type,
        RetentionAction,
        AgeLimitForRetention,
        RetentionEnabled -AutoSize

-------------
# Helpful:
-------------
# To list policies on a mailbox:
Connect-ExchangeOnline
$mailbox = "user@contoso.com"
Get-Mailbox $mailbox |
    Select-Object DisplayName,PrimarySmtpAddress,RetentionPolicy

# Get the Policy and All Associated Retention Tags
Connect-ExchangeOnline
$mailbox = "user@contoso.com"
$MailboxObject = Get-Mailbox $mailbox
$RetentionPolicy = Get-RetentionPolicy $MailboxObject.RetentionPolicy
$RetentionPolicy
$RetentionPolicy.RetentionPolicyTagLinks |
    ForEach-Object {
        Get-RetentionPolicyTag $_
    } |
    Format-Table `
        Name,
        Type,
        RetentionAction,
        AgeLimitForRetention,
        RetentionEnabled -AutoSize

The managed folder assistant will work only if the TotalDeletedItemSize  plus the TotalItemSize is over 10 megabytes. Run the following command and add those values together to check:
get-mailboxstatistics username | select totalitemsize,totaldeleteditemsize

# Check if ELC retention is active - it handles retention 
$mbs=get-mailbox user@contoso.com | select elcprocessingdisabled,retentionholdenabled, retentionpolicy, displayname, smtpaddress, userprincipalname
if ($mbx.elcprocessingdisable -eq $true) { write-host "Retention won't run"}
if ($mbx.retentionholdenabled -eq $true) { write-host "Retention won't run"}

============================================================
#>

# --------------------------------
# Configuration:
# --------------------------------

# Authentication:
$TenantId     = "<tenant-id>"       # TODO: Set
$ClientId     = "<app-id>"          # TODO: Set
$ClientSecret = "<client-secret>"   # TODO: Set

# this is the target mailbox and folder:
$Mailbox      = "user@contoso.com"         # TODO: Set
$FolderId     = "<graph-mailFolder-id>"    # TODO: Set

# Same values from the EWS sample
$PolicyTagGuid    = [Guid]"92186ff7-7f4d-4efa-a09b-c6620a8278f0"  # TODO: Set - This is the policy tag already registered on the mailbox.
 
# ------------------------------------------------------------
# Get app-only token using client credentials
# Permissions likely required:
#   Mail.ReadWrite application permission
#
# If you are testing delegated auth instead, replace this token
# acquisition block with your delegated OAuth flow.
# ------------------------------------------------------------

$TokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

$TokenBody = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

$TokenResponse = Invoke-RestMethod `
    -Method POST `
    -Uri $TokenUri `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $TokenBody

$AccessToken = $TokenResponse.access_token

# ------------------------------------------------------------
# Convert the retention policy tag GUID to the byte array format
# used by the MAPI binary property.
#
# Graph extended property Binary values are supplied as base64.
# ------------------------------------------------------------

$PolicyTagBase64 = [Convert]::ToBase64String($PolicyTagGuid.ToByteArray())
$RetentionFlags = 1  # TODO: Set
$RetentionPeriod = 30  # TODO: Set  

# ------------------------------------------------------------
# PATCH the mailFolder with the extended properties
# ------------------------------------------------------------

 

$GraphUri = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailFolders/$FolderId"

$BodyObject = @{
    singleValueExtendedProperties = @(
        @{
            id    = "Binary 0x3019"
            value = $PolicyTagBase64
        },
        @{
            id    = "Integer 0x301D"
            value = "$RetentionFlags"
        },
        @{
            id    = "Integer 0x301A"
            value = "$RetentionPeriod"
        }
    )
}

$JsonBody = $BodyObject | ConvertTo-Json -Depth 10

$Headers = @{
    Authorization        = "Bearer $AccessToken"
    "Content-Type"       = "application/json"
    "client-request-id"  = [Guid]::NewGuid().ToString()
    "return-client-request-id" = "true"
}

try {
    Invoke-RestMethod `
        -Method PATCH `
        -Uri $GraphUri `
        -Headers $Headers `
        -Body $JsonBody

    Write-Host "PATCH completed." -ForegroundColor Green
}
catch {
    Write-Host "PATCH failed." -ForegroundColor Red

    if ($_.Exception.Response) {
        $StatusCode = [int]$_.Exception.Response.StatusCode
        Write-Host "HTTP Status: $StatusCode" -ForegroundColor Yellow

        $Reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $ResponseBody = $Reader.ReadToEnd()
        Write-Host "Response Body:" -ForegroundColor Yellow
        Write-Host $ResponseBody
    }
    else {
        Write-Host $_.Exception.Message
    }
}
