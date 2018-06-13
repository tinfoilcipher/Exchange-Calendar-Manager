# ExchangeCalendarManager
Exchange Calendar Access Manager for Exchange Online 365 and Exchange On Premise 2010 and Higher

## OUTLINE
The purpose of this script is to remove the need for users to handle their calendar access, and to remove the painful exercise of managing
mass permissions from the shell manually or from the archaic method of Outlook/OWA.

## USE - GENERAL
On execution, you will be prompted to chose Exchange Online or On Premise with 1 or 2 as an input. This selects the correct connection.<br>
<br>
After connection, you will be asked if you're **GRANTING** or **REMOVING** access with 1 or 2<br>
<br>
For either choice you'll then be asked to provide a list of users to grant/remove permissions FOR. Input can be entered in the following
formats<br>
**user1,user2,user3**<br>
**user1, user2, user3**<br>
**user1; user2, user3**<br>
**Any combination of the above**<br>
Either the Exchange Aliases, primary SMTPs or userPrincipalNames should be valid inputs.<br>
<br>
You will then be asked for a list of Calendars to grant/revoke these permissions to/from, input is in the same format as above<br>
<br>
Finally an access level needs to be specified using input 1-5<br>
**1 - Owner (Full Access)**<br>
**2 - Editor (Read Write - Edit All)**<br>
**3 - Author (Read Write - Edit Own)**<br>
**4 - Reviewer (Read Only)**<br>
**5 - Contributor (Write Only)**<br>

## USE - ON PREMISE
Should work with any Exchange Server on premise where the following is true:<br>
1. Autodiscover is configured<br>
2. Remote Powershell is enabled and you have permissions<br>
3. The server has an internally resolvable FQDN<br>
Though if the above conditions aren't met, your Exchange is probably in a seriously unhealthy state to begin with.<br>
<br>
To execute, enter the FQDN of your CAS as the $strOnPremServer Constant and execute.<br>

## USE - OFFICE 365
No change to constants are needed, you will however be prompted to enter the credentials for a user that has Exchange Admin rights
to the tenancy you are connecting to.
