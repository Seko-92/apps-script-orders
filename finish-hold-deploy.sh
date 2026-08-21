#!/usr/bin/env bash
# Finishes the 2026-08-21 hold round once 1Password can be approved again.
# The SERVER half is already live (clasp push needs no agent); this is the board
# deploy plus the signed commit. Safe to re-run — every step verifies itself.
set -euo pipefail
cd "$(dirname "$0")"
export SSH_AUTH_SOCK="$HOME/.1password/agent.sock"

# ⚠ PRE-FLIGHT: ask the AGENT, not ssh. A failure here otherwise surfaces as
# "Permission denied (publickey)" — which sends you at ssh config and the server,
# and both are fine. The real answer is usually that the 1Password DESKTOP APP is
# not running: the socket file survives the app closing, so its presence proves
# nothing. Same class as the 2026-08-20 note that git's error names the wrong
# layer. Diagnosed once, encoded here so nobody re-derives it.
if ! ssh-add -l >/dev/null 2>&1; then
  echo "✗ The 1Password SSH agent is not answering on $SSH_AUTH_SOCK"
  echo
  ssh-add -l 2>&1 | sed 's/^/    /'
  echo
  echo "  'Connection refused'  → the 1Password DESKTOP APP is not running. Launch it."
  echo "  'agent refused'       → it is running but locked, or the key needs approval."
  echo
  echo "  Nothing was deployed and nothing was committed. Re-run when it answers."
  exit 1
fi

LOCAL=$(wc -c < FloorBoard.html)
echo "→ scp FloorBoard.html  ($LOCAL bytes)"
scp FloorBoard.html hetzner:/opt/hq-app/index.html
sleep 2
SERVED=$(curl -s https://hq.yassinqurabi.com/ | wc -c)
# ⚠ A byte match alone never proves the RIGHT file landed — grep the served copy.
[ "$LOCAL" = "$SERVED" ] || { echo "✗ MISMATCH $LOCAL vs $SERVED — do not commit"; exit 1; }
for pat in unlockAudio holdAcked ESCALATED htItems liftStrip holdFireLift; do
  n=$(curl -s https://hq.yassinqurabi.com/ | grep -c "$pat")
  [ "$n" -gt 0 ] || { echo "✗ '$pat' missing from the served board"; exit 1; }
  printf "   %-14s %s\n" "$pat" "$n"
done
echo "✓ board verified $LOCAL == $SERVED"

echo "→ signed commit"
git add -A
git commit -F - <<'MSG'
feat(hold): the loop closes both ways — lifted is announced, and the gate moves
to where labels are bought

Round three, all four items from the floor.

⭐ 1 · NOBODY WAS EVER TOLD A HOLD WAS LIFTED. The strip vanishing was the only
signal, and it only reached somebody already looking — so a box stayed set aside
until a human noticed, and the person who lifted it (remotely, on a phone) had no
way to know the floor ever found out. That is the same open loop this feature was
built to close, pointing backwards.

⚠ LIFTING IS STILL JUST DELETING THE WORD FROM THE NOTE. No button, no new state,
nothing to rot — that property was worth protecting, so the board DETECTS the
lift rather than being told about it.

⚠⚠ AND THE ONE FALSE POSITIVE IS EXCLUDED STRUCTURALLY, NOT GUESSED AT. An order
can also leave the held list because its ROWS were removed — and n8n deletes
shipped rows at ~1 AM Houston, which is off-hours. Gating the announcement to
working hours rules that out by construction rather than by a heuristic about why
something vanished. It also only fires for orders THIS SESSION saw held, so a
reloaded board can never announce history as news — which matters more here than
it does for the arrival beacon, because "OK TO SHIP" is an instruction.

Green, one gentle chime (the board's existing "resolved" phrase), and it retires
itself after two minutes: a lift is an EVENT, unlike the hold itself, which is a
STATE and persists until someone deals with it.

⚠ THE EMPTY-LIST BRANCH NEEDED THE CHECK TOO — and that is the commonest lift of
all. There is usually exactly one hold, so clearing it empties the list, and an
early return before the check would have shipped a feature that worked for every
case except the normal one.

⭐ 2 · THE ACKNOWLEDGE BUTTON MOVED TO FULFILLMENT, ABOVE PRINT PICK LIST — the
user's placement, and it is better than mine. A hold exists to stop a label being
bought, so the control belongs physically between the picker and the button that
starts that. It was in the Alerts card, which is where the COUNT belongs but not
the action. No new card; the alert row stays where it was.

⚠ The Fulfillment card is collapsed by default, and a gate inside a collapsed
card is not a gate — so it opens itself when a hold lands. On the RISE only,
exactly like the API-quota card: opening on every poll would fight an operator
who deliberately closed it. The button carries the count, and the hint warns
BEFORE the label rather than explaining after.

⭐ 3 · TWO GATES. The standing rule is ALERT ONCE PER CROSSING and the operative
word is CROSSING: PREPARING → SHIPPED on an unanswered hold IS one — a label now
exists and money is committed. The second message reads as an escalation, not a
repeat. It cannot become noise, because it only fires when somebody SHIPPED an
order carrying an unacknowledged hold, which is the exact disaster this prevents.

⭐ 4 · THE TAKEOVER NAMES WHAT IS IN THE BOX. An order id identifies a box to the
SYSTEM; it does not identify one to a person in an outbound area with fifteen
boxes in front of them, and a hold that sends you to a computer to look up what
you are holding has spent most of the time it saved. qty · SKU · shelf per line,
capped at 6. Free — the scan already read those columns.

Plus the two from earlier in the round: the siren now SOUNDS once off-hours
(silent was wrong in both directions), and the AudioContext is unlocked on the
first gesture of the page's life — a latent hole under the arrival beacon too,
which only became load-bearing when an ALARM depended on it.

diag-holdstop 61 → 67 · check-hold-sidebar 14 → 18 · test-holds 54 ·
test-hold-escalation 40 · check-sidebar green · node --check clean.

⚠ TWO HARNESS BUGS, both accusing working code: an inverted assertion, and a
"stayed silent" check that was really measuring the previous section's teardown —
swapping the whole held list between sections manufactures a lift, which
correctly chimes. Real holds change one at a time. Sixth and seventh instances in
this project: suspect the harness first.
MSG
git push origin main
echo
echo "✓ done — $(git log --oneline -1)"
echo "  signature: $(git cat-file commit HEAD | grep -c gpgsig)"
