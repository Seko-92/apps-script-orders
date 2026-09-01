/**
 * palette.js — one table, five verdicts.
 *
 * The verdict is computed ONCE in __SparkData!A6 and read by the face image, the D1
 * headline and now the dial. Nothing here may re-derive it; this file only says what
 * each verdict LOOKS like.
 *
 * ⚠ COLOUR MEANS THE SAME THING HERE AS EVERYWHERE ELSE ON THE SHEET:
 *     brand yellow #ffd400 = ACT ON THIS       (PREPARING, paid shipping, the divider)
 *     warm red    #ff8a5c = SOMETHING IS LATE  (the Floor Board's own 3h redline tone)
 *     alarm red   #ff6b6b = SOMETHING IS BROKEN (BRAND.redAlert, the HAND low-stock ink)
 *     cool grey   #7e8894 = QUIET, NOTHING TO DO (MASTHEAD.restAccent, shared with the
 *                                                 resting sparkline so they agree)
 *     soft green  #7ec98a = HEALTHY (the Floor Board's live lamp)
 *
 * FOUR INK ROLES, and they must not collapse into each other:
 *     accent  the state colour — rings, pips, arc, hands
 *     lead    the ONE flank value that matters right now (rest's countdown, busy's queue)
 *     dim     the other flank values — present, readable, not competing
 *     label   the small caps under each value
 *   ⚠ The first render had `dim` and `label` at the same grey, so "88 SHIPPED" read as one
 *     undifferentiated block. A value and its caption must never share a tone.
 *
 *   'clear' is deliberately NOT yellow. Yellow is the sheet's action colour and an empty
 *   queue is the one state that asks for nothing — painting it yellow is how yellow stops
 *   meaning "act". This is the same ruling that moved red off NOT FOUND.
 */
'use strict';

// The band row 1 sits on. The dial's right edge fades into this so the merge reads as one
// object rather than a rectangle pasted onto the banner.
const BAND = '#1a1a1a';

const PALETTE = {
  rest: {
    accent: '#7e8894',
    lead:   '#ffd400',
    dim:    '#aeb6c0',
    ink:    '#ececec',
    label:  '#6e7681',
    rim:    '#39352d',
    tick:   '#4a453b',
    track:  '#2f2c26',
    bg:     ['#181713', '#1e1c17', '#241f16'],
    word:   'RESTING'
  },
  clear: {
    accent: '#7ec98a',
    lead:   '#7ec98a',
    dim:    '#9fb0a5',
    ink:    '#ececec',
    label:  '#6f7a72',
    rim:    '#333b34',
    tick:   '#414a42',
    track:  '#39443b',
    bg:     ['#161815', '#1a1d19', '#1e231d'],
    word:   'ALL CLEAR'
  },
  busy: {
    accent: '#ffd400',
    lead:   '#ffd400',
    dim:    '#b6bac2',
    ink:    '#ececec',
    label:  '#8a8f98',
    rim:    '#3a3529',
    tick:   '#4a4433',
    track:  '#403a2a',
    bg:     ['#181713', '#1e1c17', '#241f16'],
    word:   'WORKING'
  },
  late: {
    accent: '#ff8a5c',
    lead:   '#ff8a5c',
    dim:    '#d3b6aa',
    ink:    '#ececec',
    label:  '#a1867b',
    rim:    '#4a3128',
    tick:   '#553a2e',
    track:  '#4b352c',
    bg:     ['#1a1512', '#201814', '#281a14'],
    word:   'PAST THE LINE'
  },
  stale: {
    accent: '#ff6b6b',
    lead:   '#ff6b6b',
    dim:    '#cfaaaa',
    ink:    '#ececec',
    label:  '#a08585',
    rim:    '#452a2a',
    tick:   '#513131',
    track:  '#472e2e',
    bg:     ['#191313', '#1d1616', '#231818'],
    word:   'NO SIGNAL'
  }
};

/** Unknown verdicts fall back to 'clear' — never to a blank canvas. */
function paletteFor(verdict) {
  return PALETTE[verdict] || PALETTE.clear;
}

module.exports = { PALETTE, paletteFor, BAND };
