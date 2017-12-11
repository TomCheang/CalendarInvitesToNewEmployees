# CalendarInvitesToNewEmployees

## Send upcoming company events to your new employees

If you're part of a rapidly growing company, you want those employees to have
upcoming company events on their calendar.  I prepared this automation for the
Comms team because they had a segment of employees that were unaware of
events that every employee has on their calendars.

The solution I devised for the team has been working without much intervention
for over two years.  The logic works as such:
1. Search Comms mailbox for future events sent to the target distribution list
2. if future events found:
  a. add the new employees to the recipient list
  b. Send out the updated event, but only to newly added recipients
  c. removed newly added recipients, save cal item but send no update messages
      (this action restores item to original state for the next run)

Originally, a copy of the invite was saved to each newly provisioned mailbox but
this proved problematic when the Comms team needed to send out a cancellation.
With the newly devised workflow above, any updates or cancellations to the calendar item
will be received by all members of your targeted distribution list.  This script
also detects whether your mailboxes are on-premises or are in Exchange Online.

**Prerequisites:**
1. A mailbox which contains calendar items sent to your targeted DL
2. ApplicationImpersonation RBAC or FullAccess rights to Comms mbx
3. specific properties to help you identify invitations, if any
4. A static or dynamic Distribution List
4. A way to determine newly provisioned mailboxes such as a custom attribute or
   a filter for WhenMailboxCreated

Once you have your prequisites configured, run this daily job by main function:

```
  $splMainFx = @{
    "CommsMbx" = "comms@yourcompany.com"
    "CreateDateAddDays" = '-1'  # age of mailboxes you want to target
  }
Add-EventsToNewEmployeeMailboxes @splMainFx
```
