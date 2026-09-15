# motion-lab — row 1 motion studies (2026-09-16)

Not deployed, not pushed by clasp (subdirectory). Sources for the row-1 banner motion work.

- `moves.html` — round one: typewriter · odometer · decode · conveyor · assembly · film title (canvas)
- `level.html` — signature moves (firing · gear · loan · scan · torque) + next level
  (kit hero · film blur via `moveBlur` · scanplus · loanplus · sheen). Superset of the old sig.html.
- `cap*.js` — Playwright frame capture → `<dir>/NNN.png` + `manifest.json`
  (⚠ they write next to themselves; `capsig.js` expects a `sig.html` — point it at `level.html`)
- `encode.py` — manifest → verified GIF (reserved eBay colours, decoded back and compared)

Canvas is 799x133 = A1:E2 (A–C 260 wide, row 1 68 + row 2 65). eBay logo is embedded as a data
URI because a file:// image taints the canvas and blocks export.
Review pages: Six Moves https://claude.ai/artifact/Lnf6FjQQ1Sd548JfRr3fg6 ·
Signature https://claude.ai/artifact/FXxs9CryVvxTy5hSMkFAE5 · Next level https://claude.ai/artifact/JM3oWYZsbHfSYYjE77VSpr
