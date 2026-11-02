# Calendar flows (from Julián's old Calendly + Asana board, saved Nov 1 2026)

Old Calendly event types (calendly.com/julian-131):

| Event | Length | Old Asana notes / questions |
|---|---|---|
| Book Home Consultation | 1 hr | Seller pitch: "You set your price, we'll advertise your home… thank you for the opportunity." 7 intake questions. |
| Book Zoom Presentation | 30 min | Remote version of the consultation, for clients wherever they live. 7 questions. |
| Schedule Home Tour | 30 min | "As your Scheduling Liaison I'll set up your itinerary…", tour with Julián or self-guided. 5 questions. |
| Schedule a Phone Call | 30 min | "I'm never too busy to take a phone call…" 6 topic options. |
| Let's Grab Coffee | 45 min | "Coffee's on me." Location list; "What would you like to order?" |
| Let's Grab Lunch / Brunch | 1 hr | Late lunch around 1–2 pm; menu list. |
| Meet for Drinks | 1 hr | "3 W's: Wine, Whiskey and Golf." Favorite spots; drink of choice. |
| Let's Go Golfing | 5 hr | 9 or 18, tee time booked?, where to book. |

Plan for the website calendar (`bookFlow`/`pickTime`, Apps Script `bookingSlots`/`bookSlot`):
- Public choices (keep it to 3, fewer choices book more): **Game-plan consultation** (in person or Zoom, 30–45 min),
  **Home tour** (30 min), **Phone call** (15–30 min). Each logs to Follow Up Boss with its answers.
- Relationship meetings (coffee, lunch, golf) stay as private links for sphere/VIPs and past clients, not on the
  public page. Drop "drinks" from anything public.
- Keep the useful questions (topic, timeline, pre-approval, agent status, price range, where they live now);
  drop menus and drink orders.

Old intake questions worth keeping (from the Calendly screenshots):
- Seller consultation (1 hr, in person or Zoom): address, phone, target list price, how much is owed, when to list,
  how they heard about Julián. Description: "You set your price and we'll advertise your home… the memories you've made
  will always remain." (Drop the old IS Luxury / Ivan Sher branding and the 702.875.1382 number; use 702-498-2571.)
- Home tour (30 min): location/MLS #, other addresses to add, tour type (drive by / in person / virtual tour), what they
  like about the home, how they heard. "As your scheduling liaison I'll set up your itinerary… any day 9am-6pm."
- Golf (Julián's favorite meeting place, ties to his podcast The Course of Business): 9 or 18, tee time booked?, where to
  book. Offer as a private invite for past clients, sphere and serious buyers/sellers, not a public booking button.

## Julián's decisions (Nov 2, 2026) — build from these
- **No coffee or lunch meetings anymore.** In person = golf course meetings only (plus home tours/showings).
- **Golf** (private invite link, not a public button): Mon 8am–3pm, Tue 8am–3pm, Thu 11am–4pm, Fri 8am–3pm (PT).
  Ask "9 or 18?" (old Calendly blocked 5 hrs for 18; 9 holes ≈ 2.5 hrs), tee time booked?, where to book.
- **Zoom / Google Meet** for out-of-town clients: Mon–Sat, no longer than 45 min (old Calendly Zoom was 30 min;
  use 30 min default, 45 max). Use the old Zoom Presentation description/questions, updated (no IS Luxury, 702-498-2571).
- **Home tours / showings**: already live (Request a showing → pick a time, Mon–Sat 9–6, 30 min).
- **Phone call**: 30 min, current hours (Mon–Sat 9–6).
- Every booking: Google Calendar event (Meet link for virtual), FUB appointment + note with answers, confirmation text,
  15-minute reminder text (remindUpcoming already handles any event with a phone in the notes).
