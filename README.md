Generates calendar events from Google Form entries.
- In this specific case, reservation requests for several vehicles are processed at once, and are either denied or turned into calendar events.

REQUIREMENTS:
- One existing Google Sheet connected to several Google Forms. First sheet must be the master sheet (not linked to any form); last sheet must be a completely empty worksheet. Order between first and last is irrelevant.
- The master sheet must have formulas inserted into columns H and I to concatenate dates and times.
- for example:
```
	=if(not(isblank(A2)),concatenate(text(E2,"mm/dd/yyyy")&" "&text(G2,"hh:mm:ss")),"")
```
- The master sheet must have up to column L, and first row as labels.
- An existing Google Calendar.
- Links in main() must be updated with the appropriate Google Sheet and Google Calendar to be accessed by the user.

FEATURES:
- Combines data from all form-linked sheets into single master sheet.
- Notes conflicts between events, denies new entries that conflict with existing ones.
- Generates unique request numbers for each entry.
- (PENDING) Sends emails to requestors notifying them that the reservation was either accepted or denied.
