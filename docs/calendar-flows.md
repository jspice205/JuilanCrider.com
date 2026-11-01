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
