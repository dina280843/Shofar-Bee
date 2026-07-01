# The Shofar Bee 🐝

A daily word puzzle from Jerusalem, in the style of NYT Spelling Bee — by **Dina Raveh** and **Shalom Brenner**.

**Live site:** https://dina280843.github.io/Shofar-Bee/
**Substack:** https://theshofarbee.substack.com/

Each puzzle gives 7 letters (one is the required center letter) and challenges players to find as many valid words as possible, with bonus points for finding the pangram — a word using all 7 letters.

---

## How the site is structured

- **`index.html`** is always the *current, live* puzzle — what visitors see when they land on the site.
- **`day1.html`, `day2.html`, ...** are archived past puzzles, one file per day.
- There is no `day6.html` while Day 6 is live — it only gets created once Day 6 is archived and Day 7 takes over `index.html`.
- Every page's header nav links to all other days, plus a **Play Today** button that always points to `index.html#game`.

## Posting a new puzzle (e.g. going from Day 6 to Day 7)

1. **Archive today's puzzle.** Copy the current `index.html` to a new file named for its day number (e.g. `day6.html`).
2. **In the new archive file** (`day6.html`), update its nav bar:
   - Add a link to itself's *former* position isn't needed — but add a forward link to the new current day (e.g. `<a href="index.html">Day 7 →</a>`), and make sure "Play Today" points to `index.html#game`.
3. **Update every other archive page's nav** (`day1.html` – `day5.html`) to add a link to the newly-archived day, and repoint the "today" link/label to the new current day.
4. **Edit `index.html`** with the new puzzle's content:
   - Hero section: day number, date, puzzle title, tagline, and the 7 letter tiles (center letter gets the `center-letter` class).
   - "Today's Inspiration" quote card: a short quote + source relevant to the puzzle's theme.
   - "How to Play" rules block: update the center-letter callout.
   - The `PUZZLE` config object near the bottom (`<script>` section): `day`, `date`, `theme`, `center`, `outer` (the 6 non-center letters), and the full `words` list (Set of valid uppercase words).
   - The hardcoded pangram string used to highlight the answer in the "Reveal All Words" list (search for the old pangram word and replace it with the new one).
5. **Upload the changed/new files** via GitHub's **Add file → Upload files**, or push via git if working locally. Don't forget to scroll down and click **Commit changes** — files dragged in but not committed won't save.
6. Give it a minute for GitHub Pages to rebuild, then check the live site.

## Puzzle word lists

Word lists are generated/verified using a Spelling-Bee-style solver before being added to a puzzle's `PUZZLE.words` set, to confirm valid words, pangram(s), and total possible score.

## Notes

- Yiddish words are allowed in gameplay (shlep, chutzpah, etc.) even if not in the strict solved word list, at the editors' discretion.
- No proper nouns.
- Minimum word length is 4 letters.
