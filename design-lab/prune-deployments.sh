#!/usr/bin/env bash
# =====================================================================================
# prune-deployments.sh — delete every Apps Script deployment EXCEPT the live one.
#
# ⚠⚠ NEVER RUN `clasp undeploy --all`. It deletes the live deployment too, which
#    instantly breaks n8n, Zoho's two webhook rules, Telegram and the Floor Board —
#    every one of them POSTs to the /exec URL that deployment owns.
#
# ⚠ CLAUDE.md's own note: archiving can itself cause edge turbulence, so do this during
#   a planned pause, not while the floor is picking.
#
# Safety, in order:
#   1. the live id is read from Secrets.js, never typed
#   2. it REFUSES to run if that id is not in the deployment list (Secrets has drifted —
#      deleting anything then would be guesswork)
#   3. the full list is written to a timestamped backup BEFORE anything is deleted
#   4. @HEAD is never touched — it is the dev deployment, not a published one
#   5. dry run unless you pass --yes
#   6. stops on the first failure rather than carrying on
#
# Usage:  ./prune-deployments.sh          # dry run, changes nothing
#         ./prune-deployments.sh --yes    # actually delete
# =====================================================================================
set -uo pipefail
cd "$(dirname "$0")/.." || exit 1

GO="${1:-}"
STAMP=$(date +%Y%m%d-%H%M%S)
BACKUP="design-lab/deployments-backup-$STAMP.txt"

LIVE=$(grep -oP 'AKfycb[A-Za-z0-9_-]{20,}' Secrets.js | tail -1)
if [ -z "$LIVE" ]; then echo "✗ could not read the live deployment id from Secrets.js"; exit 1; fi

echo "live deployment (Secrets.WEB_APP_URL): $LIVE"
echo

clasp list-deployments > "$BACKUP" 2>/dev/null
if ! grep -qF "$LIVE" "$BACKUP"; then
  echo "✗ REFUSING: the live id is not in the deployment list."
  echo "  Secrets.js and the project have drifted — resolve that first."
  exit 1
fi
echo "full list backed up to $BACKUP"
echo

# every published deployment except the live one and @HEAD
mapfile -t DOOMED < <(grep '^- AKfycb' "$BACKUP" | grep -vF "$LIVE" | grep -v '@HEAD' \
                      | sed 's/^- \([A-Za-z0-9_-]*\) .*/\1/')

echo "keep:   1  (the live one, @$(grep -F "$LIVE" "$BACKUP" | grep -oP '@\d+' | head -1 | tr -d '@'))"
echo "keep:   $(grep -c '@HEAD' "$BACKUP")  (@HEAD — the dev deployment)"
echo "delete: ${#DOOMED[@]}"
echo

if [ "$GO" != "--yes" ]; then
  echo "── DRY RUN — nothing will be deleted. Re-run with --yes to proceed. ──"
  printf '  would delete %s\n' "${DOOMED[@]}" | head -8
  [ "${#DOOMED[@]}" -gt 8 ] && echo "  … and $(( ${#DOOMED[@]} - 8 )) more"
  exit 0
fi

ok=0; failed=0
for id in "${DOOMED[@]}"; do
  if clasp undeploy "$id" >/dev/null 2>&1; then
    ok=$((ok+1)); printf "  ✓ %s  (%d/%d)\n" "$id" "$ok" "${#DOOMED[@]}"
  else
    failed=$((failed+1)); echo "  ✗ FAILED on $id — stopping."; break
  fi
  sleep 1
done

echo
echo "deleted $ok, failed $failed"
echo "── verifying the live deployment survived ──"
clasp list-deployments 2>/dev/null | grep -F "$LIVE" || echo "✗✗ LIVE DEPLOYMENT MISSING — restore from $BACKUP"
