/**
 * light.js — ONE day-light curve, shared by the face and the sky.
 *
 * ⚠ It lived in BOTH masthead.html and sky.html and I flagged the drift risk in a comment
 *   rather than removing it. Two copies of one rule is how A-9 sorted after A-50 in three
 *   files until August. It is a file now, loaded by both.
 *
 * ⚠ The first curve used cos((h-13)/24) and hit ZERO at 07:00 and 19:00, so 7am rendered
 *   as full night with stars. Houston is light well before the floor opens. This is a
 *   sin() arc across 05:00 → 21:00 peaking at 13:00 — a 16-hour day, which is what a
 *   Gulf-coast day actually looks like.
 */
(function (root) {
  function dayLight(h) {
    var sun    = Math.max(0, Math.sin(((h - 5) / 16) * Math.PI));   // 0 at 05, 1 at 13, 0 at 21
    var golden = Math.max(0, 1 - Math.abs(sun - 0.35) * 3.2);       // low-but-present sun
    var dusk   = h >= 15;                                            // ember side of the day
    var g      = 0.055 + sun * 0.055;
    return {
      ground: 'rgb(' + Math.round(255 * (g + 0.012)) + ',' +
                       Math.round(255 * (g + 0.006)) + ',' +
                       Math.round(255 * (g * 0.86)) + ')',
      sky:    sun < 0.06 ? '#2b3a5e'
            : golden > 0.45 ? (dusk ? '#ff6a1f' : '#ffa23d')
            : '#ffe9a8',
      skyO:  (0.10 + sun * 0.30 + golden * 0.34),

      // ⚠⚠ THESE TWO WERE MISSING FOR THE WHOLE LIFE OF THE v2/v3 SETS, and their
      //    absence is most of why the banner "didn't look how I wanted".
      //
      //    masthead.html read L.horiz and L.mark and set them as custom properties.
      //    setProperty STRINGIFIES undefined, so the property was not left unset —
      //    it was SET to the token `undefined`. That matters: var(--markDim, 1)'s
      //    fallback only fires for an UNSET property, never for one holding garbage.
      //    Both substitutions produced `opacity: undefined`, which is invalid at
      //    computed-value time, so opacity reset to its INITIAL value — 1.
      //
      //    Consequences on all 144 shipped faces: .horizon rendered at FULL opacity
      //    at every hour (a hard coloured bar along the bottom edge — blue at night,
      //    orange at golden hour, cream at midday), and .mark never dimmed at all,
      //    contradicting its own comment.
      //
      // ⚠ A missing property is not a no-op. It is worse than a wrong value, because
      //   the fallback you wrote to protect against it cannot fire.

      // The identity never fades below 62% — the ground takes the hour, the mark does
      // not. This curve is specified by masthead.html's own comment, which was written
      // and then never implemented.
      mark:  0.62 + sun * 0.38,

      // The horizon line is a GLOW, not a rule: faintest at night, strongest at golden
      // hour when a real horizon burns, moderate at noon when the light is overhead and
      // the edge goes flat. Capped so it can never become the hard bar it used to be.
      horiz: Math.min(0.55, 0.08 + sun * 0.22 + golden * 0.35),

      sun: sun, golden: golden,
      night: sun < 0.10
    };
  }
  root.dayLight = dayLight;
})(typeof window !== 'undefined' ? window : globalThis);
