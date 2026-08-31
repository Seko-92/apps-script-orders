#!/usr/bin/env bash
# ship-masthead.sh — scp the faces, then verify against WHAT THE SERVER SERVES.
#
#   VER=v4 ./ship-masthead.sh              ship + verify v4
#   VER=v4 PREV=v3 ./ship-masthead.sh      ...and sweep v3 ONLY after v4 verifies
#
# ⚠ A deploy is not done because the command returned; it is done when the server says
#   the right bytes are there. Two silent scp failures are already on this project's record.
# ⚠ Never read $? through a pipe — that is tail's exit code, not scp's (2026-08-21).
#
# ⚠⚠ THIS SCRIPT USED TO HARDCODE THE VERSION IN THREE PLACES — the scp glob, the
#    byte-verify filename and the verify URL — plus a fourth for the sweep. By 2026-08-31
#    the live set was v3 and v2 had already been removed from the server, so running it
#    would have RE-UPLOADED the stale 168-file v2 set, then byte-verified against the very
#    files it had just wrongly created, and printed ✅. A check that manufactures its own
#    passing condition is worse than no check.
#
# ⚠⚠ AND IT SWEPT THE OLD SET BEFORE VERIFYING THE NEW ONE. That is backwards: a failed
#    scp plus a completed sweep leaves the server with NO usable faces, and a missing file
#    under /mast/ answers 200 with the Floor Board's HTML — so every face silently becomes
#    the text chip. The sweep is now opt-in AND runs only after the verify passes.
set -uo pipefail
export SSH_AUTH_SOCK="$HOME/.1password/agent.sock"
cd "$(dirname "$0")/renders/mast" || exit 1

VER="${VER:-v4}"
PREV="${PREV:-}"
REMOTE=/opt/hq-app/mast
BASE=https://hq.yassinqurabi.com/mast

if ! timeout 10 ssh-add -l >/dev/null 2>&1; then
  echo "✗ the 1Password SSH agent is not answering."
  echo "  'Connection refused' = the DESKTOP APP is not running (not a pending approval)."
  echo "  Start + unlock 1Password, then re-run this script."
  exit 1
fi

# ⚠ Derive the state list from what is actually ON DISK rather than hardcoding it. The old
#   list named 'held', a state the verdict chain can never produce, so the verify reported a
#   false failure on a face we deliberately no longer wire up.
shopt -s nullglob
LOCAL=( *-h11-"$VER".png )
if [ ${#LOCAL[@]} -eq 0 ]; then
  echo "✗ no *-h11-$VER.png here. Render first:  VER=$VER node shoot-masthead.js"
  exit 1
fi
STATES=()
for f in "${LOCAL[@]}"; do STATES+=( "${f%%-h11-$VER.png}" ); done
COUNT=$( ls *-"$VER".png 2>/dev/null | wc -l )
echo "shipping $COUNT file(s) · version $VER · states: ${STATES[*]}"

ssh hetzner "mkdir -p $REMOTE" || { echo "✗ cannot reach the VPS"; exit 1; }

# ⚠ PNG, not GIF, for the pinned face: =IMAGE() shows only a GIF's first frame. The GIFs
#   ride along for any future surface that scrolls, where insertImage() DOES animate.
scp *-"$VER".png "hetzner:$REMOTE/"
SCP_RC=$?                                   # captured BEFORE any pipe
[ $SCP_RC -eq 0 ] || { echo "✗ scp failed (rc=$SCP_RC)"; exit 1; }
scp *.gif "hetzner:$REMOTE/" 2>/dev/null

echo
echo "=== byte-verify against the served copy (sampled across the day) ==="
fails=0
for f in "${STATES[@]}"; do
  for h in 03 08 11 17 21; do
    l=$(stat -c%s "$f-h$h-$VER.png" 2>/dev/null)
    d=$(curl -s -o /dev/null -w '%{size_download}' "$BASE/$f-h$h-$VER.png")
    if [ -n "$l" ] && [ "$l" = "$d" ]; then printf "  ✓ %-6s h%s  %s bytes\n" "$f" "$h" "$d"
    else printf "  ✗ %-6s h%s  local=%s served=%s\n" "$f" "$h" "$l" "$d"; fails=$((fails+1)); fi
  done
done

# ⚠ A 303,167-byte "image" is the Floor Board's HTML, not a face. Caddy's try_files serves
#   it for any unmatched path, so a MISSING file answers 200 — size alone can look fine if
#   you only eyeball it. Call it out by name.
echo
one=$(curl -s -o /dev/null -w '%{size_download}' "$BASE/${STATES[0]}-h11-$VER.png")
if [ "$one" -gt 200000 ]; then
  echo "  ⚠⚠ served size ${one}B looks like the board's HTML — the file is MISSING, not wrong."
  fails=$((fails+1))
fi

if [ $fails -ne 0 ]; then
  echo "✗ $fails sample(s) did not land — NOT sweeping anything."
  exit 1
fi
echo "✅ $COUNT faces live at $BASE/  (version $VER)"

# ---- the sweep, LAST and opt-in ------------------------------------------------------
if [ -n "$PREV" ]; then
  if [ "$PREV" = "$VER" ]; then
    echo "✗ refusing to sweep PREV=$PREV — that is the version just shipped."
    exit 1
  fi
  echo
  echo "sweeping superseded set $PREV ..."
  ssh hetzner "rm -f $REMOTE/*-$PREV.png" && echo "  ✓ $PREV removed"
else
  echo
  echo "  (no PREV set — superseded versions left in place. Pass PREV=vN to sweep one.)"
fi
