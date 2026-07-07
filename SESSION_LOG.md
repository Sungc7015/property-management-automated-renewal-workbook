# Session Log

## [2026-07-07] Pending section rework + Confirmed/Pending data flow fixes — v2.9.0

Follow-up session to v2.8.0, same day. Expanded Pending (Manual) from 5 to 9 columns (Expected Increase Date, New MTM Rate, Market Rent, Notes), with Market Rent now auto-imported from the Rent Roll and New MTM Rate also picked up live from month sheets' col S ("MTM Rate"). Fixed "Sort MTM Tracker" corrupting the live MTMTable header's formatting/position by switching it to sort via the Table's own `Sort` object; fixed `NextIncreaseDate` being off by a day (now exactly `=EDATE(...,12)`); stopped both Sort and Refresh from overwriting an existing Next Increase. Fixed a real bug where Pending would re-import a unit indefinitely once flagged "MTM" on any month sheet in its 3-month scan window, even after a later sheet showed it resolved. Added a check so a unit already picked up as a real Confirmed row this refresh is excluded/dropped from Pending instead of reappearing after being manually deleted per the on-sheet reminder note.

That last fix initially caused a regression the user caught in testing (needed Pending entries disappearing) — traced via a second Fable audit pass to 3 real bugs in the accompanying Pending→Confirmed data carryover, not the exclusion logic itself: New MTM Rate/Notes were being silently lost on the transition (not just Expected Increase Date), Expected Increase Date was being written to the wrong column and double-projected 12 months late, and a Market Rent value could leak from one unit to the next in the rebuild loop due to a stale VBA loop variable. All fixed and confirmed by the user before this push. Two full Fable-model audit passes ran this session in total (per the multi-model-audit practice for high-stakes VBA with no runtime available to verify).

Documented in README.md (version bump + changelog row).

**Artifact:** `modMTM.bas`, `modDynamic.bas`, `modReaders.bas`, `README.md` — commits `4b89529` (feature/fixes), `6adfa35` (version bump), `d57a25f` (README docs), pushed to `origin/master`.

## [2026-07-07] Pending (Manual) MTM tracker feature — v2.8.0

Added a "Pending (Manual)" section to the MTM tracker sheet so units manually flagged "MTM" on month sheets (ahead of the rent roll confirming it) surface immediately, via a live change-trigger plus a refresh-driven rebuild, scoped to a rolling 3-month window anchored to the last import so old years' month sheets never resurface stale statuses. Planned, implemented, audited by three independent models (Fable/Opus/Sonnet, reconciled by orchestrator — caught a real row-math bug and an anchor-name crash risk before Excel testing), fixed one cosmetic date-formatting issue the user found in testing, then committed and pushed as v2.8.0.

Also documented the feature in README.md (new "Pending (Manual) Section" under MTM Tracker Workflow, plus a cross-reference from the Renewal Status column reference); checked PULLING_REPORTS.md, no changes needed there.

**Artifact:** `modMTM.bas`, `modDynamic.bas`, `modImport.bas`, `modSheetUtils.bas`, `README.md` — commits `8129a83` (feature), `36998ab` (version bump), `738e10d` (session log), `b20a49a` (README docs), all pushed to `origin/master`.
