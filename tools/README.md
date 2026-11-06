# Quick property cards

Copies Jessica's card (`sa/`) for a real agent. Only the agent details, listings and
photos change. Design, booking, calculator, lead form and preview ribbon stay the same.

1. Make a blank sheet: `python3 tools/quick_card.py --blank agents.xlsx`
2. Fill the **Agents** tab (one row per agent) and the **Listings** tab (one row per home).
3. Put photos in `photos/<slug>/`: `agent.jpg`, optional `hero.jpg` (else the first
   listing photo is used), and the listing photos named in the Listings tab.
4. Build: `python3 tools/quick_card.py agents.xlsx photos/`
   (`--only <slug>` builds one agent, `--force` rebuilds an existing folder).

Each card lands in `<slug>/` and opens at `property.cevyq.com/<slug>/` once merged.

Nothing made up is carried over from Jessica: her reviews, sold homes, sales figures,
ticker and licence numbers are cleared. Only what is in the sheet shows. Every card is
a 7 day preview (ribbon, then lock), not shown on Google, counted on Card Desk as
`property-<slug>`, and has a privacy line in the small print. Country SA, US or AU
sets the money, calculator and price (R1,500, $75, A$110).
