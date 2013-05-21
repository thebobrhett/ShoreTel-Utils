ShoreTel-Utils
==============

3rd party web pages to monitor phone system usage

Intended audience:
  Accounting - Provides a set of tools to monitor system usage for allocating system costs to individual cost centers without having to use account codes.
  Human Resources - Provides a set of tools to monitor for phone system abuse.
  Safety - Provides a history of 911 or other emergency related calls.
  IT - Provides a set of tools for monitoring system usage relative to system capacity so that bandwidth can be maintained accordingly.

These pages are written in ASP and are intended to be used on an intranet in conjunction with a ShoreTel phone system.
I found the reporting features built in to the ShoreTel server lacking so I created my own.
This application reads the ShoreTel databases to report usage of the system
It requires read only access to the ShoreTel database server across the network.

It reports the peak system usage (how many lines were in use simultaneously) by day.
It reports any calls to 911 within the past 30 days.
In our case we have an internal emergency response number (for onsite EMS or Fire Brigade). Calls to this number are reported.
In our case we have phones in the elevators that autodial when lifted. Calls from these phones are reported.
In our case we have an information line in case of local emergency (hurricane, snow, etc) to provide workers with operational status of the business. Calls to this number are reported.

It reports the top 5 (configurable) for the past 30 days (configurable) in the following categories:
  Most frequently called numbers, inbound
  Most frequently calling numbers, inbound
  Most frequently called numbers, outbound
  Most frequently calling numbers, outbound
  Most frequently called numbers, internal
  Most frequently calling numbers, internal
  Most talk time
  Longest duration calls
  Longest hold times, inbound calls only
  Longest duration long distance calls
  Most frequently calling numbers, long distance

There is a statistics section that shows:
  Cost of the phone system per month (derived from manually entered billing data)
  Sum of talk time for past 30 days in minutes
  Cost of talk time per minute (50% of cost divided by minutes)*
  Total number of extensions in system
  Cost of an extension per month (50% of cost divided by extensions)*
* We allocated the cost of phone system usage to cost centers based on 50% usage and 50% number of extensions.

There is a cost center overview section that shows number of extensions, usage in minutes, and the resulting allocation (based on the aforementioned calculation) per cost center*
* I used the "Extension Lists" feature of the ShoreWare Director to group extensions into cost centers

All of these sections provide links to detail pages so that usage history of an individual extension or group can be examined.
