// eBay Feedback Notify (v1) → node "Build Summary" — paste over the whole code.
// 2026-10-07: Telegram rejected a summary ("can't parse entities") because an item
// title carried the pack-size note "* 4": a lone * opens bold text that never closes.
// eBay usernames with "_" do the same. Everything that comes from eBay is escaped
// for Telegram's (legacy) Markdown before it goes into the message.
const esc = s => String(s == null ? '' : s).replace(/([_*`\[])/g, '\\$1');

const feedbacks = $input.all().map(it => it.json).filter(f => f && f.feedbackId);
const sellerUser = 'hqmotorservice'; // your eBay username - edit if different
const url = `https://www.ebay.com/fdbk/feedback_profile/${sellerUser}`;

let text;
if (feedbacks.length === 0) {
  text = `✅ No feedback waiting for a reply right now.`;
} else {
  let bodyTxt = '';
  for (let i = 0; i < feedbacks.length; i++) {
    const f = feedbacks[i];
    bodyTxt += `${i + 1}. ${esc('[' + f.type + ']')} ${esc(f.buyer)}\n   “${esc(f.comment)}”\n   (${esc(f.itemTitle)})\n\n`;
  }
  text = `📨 Feedback to reply (${feedbacks.length}):\n\n${bodyTxt}🔗 Reply on eBay: ${url}`;
}
return [{ json: { text: text } }];
