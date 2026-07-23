# MTM Tab — Quick Reference for Property Managers

*A practical guide to the `MTM` tab's four buttons: what's safe, what's not, and what to check before you click Refresh.*

## What this tab is for

The `MTM` tab tracks all your month-to-month units in one place, separately from the monthly renewal sheets. The top block (**Confirmed**) is built straight from the Yardi Rent Roll — it's the official record. The block below it (**Pending (Manual)**) is a heads-up list of units you've flagged as going MTM on a recent month sheet, shown here *before* Yardi's Rent Roll has caught up and confirmed it. Some columns on this tab are rebuilt fresh from your data every time you click Refresh; others are yours to type into and are carried forward automatically. Knowing which is which is the point of this guide.

---

## Before You Refresh — Checklist

- [ ] Do you know which units you *don't* want to lose from Pending? Check **Keep/Lock** on those rows first — a plain refresh can silently drop a Pending row that's aged out of its 3-month window.
- [ ] Have you already typed in any Notes or Last Increase dates you want kept? They carry forward automatically **only if the unit is still MTM in the Yardi Rent Roll you're about to import** — if it's not, that row and everything on it is deleted.
- [ ] Are you mid-edit on any cell, or mid-click on a checkbox? Finish it (press Enter or click elsewhere) before clicking any of the four MTM buttons.
- [ ] Is this the correct, current Yardi Rent Roll export? Refresh only knows what's in the file you select.
- [ ] If you're doing a big batch of changes, consider saving a copy of the workbook first, purely as a snapshot you can compare against if something looks off afterward.

---

## Do's

- **Run Refresh MTM Tracker regularly.** It's the only thing that moves a unit from Pending into the official Confirmed section once Yardi's Rent Roll actually shows it as MTM.
- **Check Keep/Lock on a Pending row the moment you're not 100% sure this month's Import Monthly Data has already been run.** The Pending window is based on the last month you formally imported, not today's date — a unit flagged MTM on a very recent sheet can otherwise be invisible to Pending (live and on refresh) until that import happens, then quietly age back out again a few months later if it's still sitting in Pending.
- **Type directly into the columns meant for manual entry** — in Confirmed: Last Increase and Notes; in Pending: Expected Increase Date, New MTM Rate, and Notes. These are preserved across every Refresh and Sort.
- **Resolve a unit by changing its Renewal Status on the month sheet** (e.g., to "Renewed") once you know the outcome. This is the only thing that permanently removes a Keep/Lock-protected row from Pending — leaving the status blank does not count as resolved.
- **Let a unit disappear from Pending on its own once it shows up as its own row in Confirmed above.** That's expected, not a bug — its Expected Increase Date becomes the new row's Next Increase, and its New MTM Rate/Notes get folded into that row's Notes automatically.
- **Click Refresh before Sort if you want the freshest numbers.** Sort only reorders and re-labels what's already sitting on the sheet — it never pulls anything new from Yardi.
- **If Import Selected MTM tells you some units were skipped**, know that their Import checkbox is left checked (not cleared) — fix the reported issue (e.g., the target month sheet doesn't exist yet, or the floor plan isn't set up) and just click Import again; you don't need to re-check anything.

## Don'ts

- **Don't assume editing Last Increase updates Next Increase.** It doesn't, automatically. Once a unit has a Next Increase date on record, both Refresh and Sort carry that exact date forward untouched forever — even after you change Last Increase. To force it to recalculate, clear/delete the Next Increase cell for that row yourself, then click Refresh MTM Tracker.
- **Don't type directly into the Import checkbox cell** (Confirmed's rightmost column) **or the Keep/Lock checkbox cell** (Pending's rightmost column) — click the checkbox control itself. Those cells are deliberately formatted so their real value never displays as text, so typing something into one directly can look fine on screen while behaving unpredictably, and it won't visually match the checkbox until the next Refresh or Sort resyncs it.
- **Don't grey-fill and bold a cell in column A of the Confirmed section** to highlight or flag a row manually. That exact combination — grey background plus bold text — is what the tool uses internally to detect where the Confirmed section ends and Pending begins. Using it for any other purpose can confuse Import, Sort, Clear, and Refresh about which rows belong to which section.
- **Don't click Clear Selection MTM unless you actually mean to uncheck every Import checkbox in the Confirmed section.** It clears all of them in one shot — there's no per-row undo once you've clicked past it. (It only touches Import checkboxes — Notes, Last Increase, and Pending's Keep/Lock boxes are untouched by it.)
- **Don't rely on Excel's Undo to save you after a Refresh, once you've saved and closed the workbook.** If a unit is no longer flagged MTM in the Yardi Rent Roll file you select, Refresh MTM Tracker deletes its entire Confirmed row — including any Notes or Last Increase you typed on it — and nothing is kept anywhere else on the sheet. This is permanent the moment you click Refresh.
- **Don't assume a duplicate row in Pending is harmless.** If the same unit ever ends up listed twice there, running Sort MTM Tracker will silently delete the extra copy, keeping only the first one it finds — if the two rows had different Notes, Expected Increase Date, or Keep/Lock settings, whichever version was on the deleted row is gone for good.
- **Don't click Refresh, Import, Clear, or Sort while you're still typing in a cell or mid-click on a checkbox.** All four rebuild every checkbox on the sheet from scratch every time they run, so triggering one on a sheet that's mid-edit is asking for a mismatch between what's checked and what the sheet shows.

---

## The Keep/Lock Checkbox — When and Why

Keep/Lock is the newest feature on this tab (added in the July 2026 update) and the one most worth understanding before you rely on it.

**What it is:** a checkbox in the rightmost column of the Pending (Manual) section, one per row.

**What checking it does:** keeps that specific row in Pending through every future refresh, even if:
- it falls outside the normal rolling 3-month window that Pending otherwise watches, or
- it graduates and shows up as its own row in the Confirmed section above — you'll then see the same unit listed in *both* sections at once. That's intentional, not a glitch, as long as it's locked.

**What removes a locked row:** only one thing — marking that unit's Renewal Status as something other than "MTM" (e.g., "Renewed") on a month sheet the tracker is watching. A blank status does not count as resolved and will not clear it.

**Why the window can silently drop a unit in the first place:** the Pending window isn't based on today's date — it's based on whichever month you most recently ran Import Monthly Data for. If that import is behind schedule, or if a unit has been sitting in Pending for a while without graduating to Confirmed, its original month sheet can eventually fall outside the 3-month window the tracker is currently watching. When that happens on a refresh, the row is deleted with no warning message — the sheet just quietly no longer lists it.

**When to check it:**
- Right after you mark a unit MTM on the current month's sheet, if you're not sure this month's formal import has already run.
- For any unit you want to keep visible in Pending as a running reminder, even once it becomes a real Confirmed row.

**Important:** once a Pending row has already disappeared, checking a box does nothing — there's no row left to check. To bring it back, re-set that unit's Renewal Status to "MTM" again on a month sheet inside the current window, and check Keep/Lock this time so it doesn't happen again.

---

## Confirmed vs. Pending — What Survives a Refresh

### Confirmed (top section)

| Column | Rebuilt fresh from Yardi every refresh | Carried forward / manual |
|---|---|---|
| Unit, Name, Floor Plan, Lease Expiry, Current Rent, Market Rent | Yes | |
| Last Increase | | Yes — your entry, always carried forward |
| Next Increase | | Yes — carried forward untouched once set; only computed fresh if blank |
| Status (Eligible / Not Yet Eligible) | Yes — recalculated from Next Increase, informational only | |
| Notes | | Yes — your entry, never overwritten; an auto-note may be added *only* if Notes is blank and Last Increase is blank on a lease that's 15+ months overdue |
| Import checkbox | | Yes — carried forward if the row already existed; new rows start unchecked |

**If a unit drops out of Yardi's MTM list entirely, its whole Confirmed row — including Notes and Last Increase — is deleted on the next refresh with nothing kept.**

### Pending (Manual) (bottom section)

| Column | Rebuilt fresh from the month sheet every refresh | Carried forward / manual |
|---|---|---|
| Unit, Name, Floor Plan, Current Rent, Source Month | Yes (or restored by Keep/Lock if the unit aged out but stayed locked) | |
| Market Rent | Yes — pulled fresh from the Rent Roll; falls back to its last known value only if the unit isn't found this time | |
| Expected Increase Date, New MTM Rate, Notes | | Yes — your entries, carried forward every rebuild |
| Keep/Lock checkbox | | Yes — carried forward; defaults to unchecked for a brand-new row |

---

## If Something Looks Wrong

Don't panic — a missing row usually has a boring explanation. Work through these before assuming data was lost:

1. **Check the Confirmed section first.** A unit that vanished from Pending most likely graduated there — that's the normal, expected outcome, and its Expected Increase Date/New MTM Rate/Notes should have carried over into the new row's Next Increase and Notes.
2. **If it's not in Confirmed either**, and it had been sitting in Pending for a while without Keep/Lock checked, it likely aged out of the rolling 3-month window. The underlying "MTM" flag is presumably still sitting on its original month sheet — the tracker just isn't watching that sheet anymore.
3. **To bring it back**, re-set that unit's Renewal Status to "MTM" on a current, in-window month sheet, then check its Keep/Lock box this time so it doesn't happen again.
4. **If a Confirmed row disappeared entirely**, check the Yardi Rent Roll file you selected for that refresh — if the unit no longer shows as MTM there, the tool has no way to know it should keep the row, and any manual Notes/Last Increase on it are gone.
5. **If a checkbox looks unchecked when you thought you'd checked it (or vice versa)**, consider whether you clicked a button while the sheet was mid-edit — try the operation again on an idle sheet.
6. **When in doubt, don't run another Refresh/Sort/Import/Clear on top of a state that already looks wrong.** Save a copy of the workbook first so you have something to compare against, then investigate from there.
