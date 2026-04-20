# Main Menu — Design Notes & Deferred Decisions

Purpose: capture the navigation-UX options for the HDOS Tools menu so we can revisit this decision without re-deriving it. The menu itself works; what's open is *how the user moves between the menu and the tools*.

---

## Current (temporary) default

- Menu cards use `target="_blank"` — each tool opens in a **new browser tab**.
- Tool pages have **no** "Back to Menu" link. The menu tab stays open in the background; the user returns by switching tabs or closing the tool tab.

Where the default lives:
- Menu cards: [main.py](main.py), look for `href="/store-report/" target="_blank"`.
- Tool pages do not contain any back-link markup.

This was chosen because the user's original request was "the menu should open the tool in a new tab so the menu is still available if you need it" — which this satisfies with the minimum UI surface on each tool page.

---

## Navigation options considered

### Option A — New tab only *(current default)*
- Menu uses `target="_blank"`.
- No back link or close button on tool pages.

**Pros:** simplest. Menu is always there via tab switching. No code changes to tool HTML.
**Cons:** users who close the tool tab mid-task to "go home" have no visible path back. On a phone/tablet browser, tab switching is clunkier than a link.

### Option B — Same tab + "← Back to Menu" link on every tool
- Menu uses plain `href="/store-report/"` (no `target`).
- Each tool's HTML has a `<a href="/">← Back to Menu</a>` in its header.

**Pros:** one-tab experience, familiar back-link pattern, works well on mobile.
**Cons:** menu tab is no longer "pinned" — navigating to a tool replaces it. If a tool is mid-upload and the user clicks Back, the work is lost.

### Option C — New tab AND back link (belt-and-suspenders)
- Menu uses `target="_blank"`, AND each tool has a "← Back to Menu" link.

**Pros:** accommodates both mental models. Menu tab stays alive; if a user *does* close it, the link gets them back.
**Cons:** slight UI clutter. Clicking the back link in a tool tab leaves the *original* menu tab still open too, so now there are two menu tabs.

### Option D — New tab + "Close this tab" button
- Menu uses `target="_blank"`, each tool has a "Close this tab" button that calls `window.close()`.

**Pros:** nice tidy-up gesture after finishing a task.
**Cons:** `window.close()` only works on tabs opened *by script*, not by an `<a target="_blank">` click in most modern browsers. So this button silently does nothing in most cases. Not recommended.

---

## Suggested decision framework

Pick based on who the primary user is and how they'll actually use this:

- **Manager using it at a desktop/laptop once a week**, short sessions: **Option A** (current) is fine. They'll learn tab-switching quickly.
- **Manager using it on a tablet or frequently switching between tools**: **Option B** (same tab + back link) is more ergonomic.
- **Multiple people, uncertain habits, want to be safe**: **Option C** (both) is the hedge.

---

## Future considerations

- **Third tool and beyond:** the current `grid-template-columns: 1fr 1fr` in [main.py](main.py) splits into two columns on desktop. With 3–4 tools, consider `repeat(auto-fit, minmax(280px, 1fr))` so the layout grows gracefully.
- **Menu status hints:** could the menu show useful live info? Examples: "Template: ✓ loaded" (green check) or "Last report: 2026-04-14" — would need a small Flask endpoint and a bit of JS on the menu page. Not urgent.
- **Keyboard shortcuts:** e.g., `1` / `2` keys to jump into tools. Nice-to-have.
- **Tool icons:** simple SVGs or emoji on each card would make the menu more scannable.

---

## Decision log (fill in when revisited)

- [ ] Chose option: ______________
- [ ] Date of decision: ______________
- [ ] Reasoning: ______________
