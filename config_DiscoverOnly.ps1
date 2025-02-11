<#
.SYNOPSIS
    Configuration file for msuserstats.ps1
#>

<###################################
    BASICS
####################################>
# Entra Tenant ID
$CONF_TENANT_ID = "ece9bbd7-ece7-4242-a428-72ca97e9ff7a"

# Include all users from Active Directory
$CONF_INCLUDE_ACTIVE_DIRECTORY = $true

# Query Exchange Online for mailbox types
$CONF_QUERY_EXCHANGE_ONLINE = $true

# Days before a user is marked as inactive without a sign-in
$CONF_INACTIVE_DAYS = 90

<###################################
    ADVANCED
####################################>
# Groups with MFA enabled users - if users are not enabled for MFA by default
$CONF_MFA_ENABLED_SECURITY_GROUPS = @()

# Groups containing user with MFA exceptions
$CONF_MFA_EXCEPTION_SECURITY_GROUPS = @()

# Create user review list for all company entities. Set $CONF_AD_DN_COUNTRY_LEVEL and/or $CONF_AD_DN_ENTITY_LEVEL to find entity from AD OU. 
$CONF_CREATE_ENTITY_INACTIVE_USERLISTS = $false
$CONF_ENTITY_FILE = ""
# Path to create Entity files
$CONF_ENTITY_OUTPUT_PATH = $CONF_MY_WORKDIR
# Discover AD admin accounts by searching the AD sAMAccountname with -match
$CONF_AD_ADMIN_SEARCH = @('adm')
# Discover EID admin accounts by searching the EID UPN with -match
$CONF_EID_ADMIN_SEARCH = @('adm')

# Discover account type by searching the AD distinguished name (DN). Contains is used to search inside DN - all lower case
$CONF_ACCOUNT_TYPE_SEARCH = @{ 
    "External User" = ""
    "Service Account" = ""
    "Internal User" = ""
}

# Search groups where users are member of. Membership will be checked for all groups listed here.
# Groups are selected using a Contains search on DisplayName using this setting
$CONF_GROUP_MEMBERSHIP = @()
# To limit the amount of groups to search for, enter a common prefix like "sec-"
$CONF_GROUP_MEMBERSHIP_FILTER_PREFIX = ""

# Cleanup Exceptions
$CONF_FILE_CLEANUPEXCEPTIONLIST = ""

# Determine country and organization unit from AD distinguished name if you use OUs in AD
# Set 0 to disable
# Example DN: CN=Mustermann\, Max,OU=Users,OU=Sub,OU=DE,DC=europe,DC=company,DC=org
# Use the DN and count from left to right along equal signs to find the position of country. Adjust if needed to your AD OU structure.
$CONF_AD_DN_COUNTRY_LEVEL = 0
# Use the DN and count from left to right along equal signs to find the position of the OU. Adjust if needed to your AD OU structure. 
$CONF_AD_DN_ENTITY_LEVEL = 0

<###################################
    OPTIONS
####################################>
# Working directory of this script
$CONF_MY_WORKDIR = $PWD.Path

# Delimiter to use in CSV files
$CONF_CSV_DELIMITER = ";"

# All Users temporary working file
$CONF_FILE_ALL = "$($CONF_MY_WORKDIR)/all_users.csv"

# Selected users temporary working file
$CONF_FILE_REQUESTEDTYPE = "$($CONF_MY_WORKDIR)/user_type_$($UserType)"

# Export final file with requested users as Excel sheet in addition to this directory. It will always be exported to $CONF_MY_WORKDIR
$CONF_EXCEL_EXPORT_PATH = ""

# Marker/Keyword for inactive user category
$CONF_INACTIVE_KEYWORD = "Inactive"
