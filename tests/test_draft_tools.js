const assert=require('node:assert/strict');
const t=require('../web/office/draft-tools.js');
// HTML: escaping, paragraphs, line breaks and real lists
const html=t.toHtml('Hi <Jo> & co,\n\nPoints:\n- one\n- two\n\n1. first\n2. second\nLine A\nLine B');
assert.ok(html.includes('Hi &lt;Jo&gt; &amp; co,'));
assert.ok(html.includes('<ul style="margin:0 0 11pt 0"><li>one</li><li>two</li></ul>'));
assert.ok(html.includes('<ol style="margin:0 0 11pt 0"><li>first</li><li>second</li></ol>'));
assert.ok(html.includes('Line A<br>Line B'));
assert.ok(!/<script/i.test(t.toHtml('<script>alert(1)</script>')));
// Review rules
const ids=text=>t.review(text).map(r=>r.id);
assert.deepEqual(ids('Kind regards'),[]);
assert.ok(ids('Please send names by [confirm deadline].').includes('placeholder'));
assert.equal(t.review('[confirm a] and [TBC b]').find(r=>r.id==='placeholder').message.startsWith('2 placeholders'),true);
assert.ok(ids('The door code is 4821.').includes('security-code'));
assert.ok(ids('Alarm code: 1234').includes('security-code'));
assert.ok(ids('Your PIN is 0042').includes('security-code'));
assert.ok(ids('password: Winter2026!').includes('security-password'));
assert.ok(ids('Card number 123456789').includes('security-card'));
assert.ok(ids('Note CCTV is offline on level 2').includes('security-ops'));
assert.ok(!ids('The code of conduct applies from 2026.').includes('security-code'));
assert.deepEqual(t.review('The door code is 4821.',{securityGuard:false}).map(r=>r.id),[]);
assert.ok(ids('Please find attached the report.').includes('attachment'));
assert.ok(!t.review('Please find attached the report.',{hasAttachments:true}).some(r=>r.id==='attachment'));
assert.ok(ids('We will refund the fee.').includes('commitment'));
assert.ok(ids('Please reply by Friday.').includes('deadline'));
assert.equal(t.review('door code 4821').find(r=>r.id==='security-code').level,'stop');
// History
const h=t.history(3);h.push({v:1});h.push({v:2});h.push({v:3});h.push({v:4});
assert.equal(h.size,3);assert.equal(h.current().v,4);assert.equal(h.back().v,3);h.push({v:5});
assert.equal(h.size,3);assert.equal(h.current().v,5);assert.equal(h.forward().v,5);h.update({v:6});assert.equal(h.current().v,6);
// Names
assert.equal(t.firstName('Citizen, Jordan'),'Jordan');assert.equal(t.firstName('Jordan Citizen (Contractor)'),'Jordan');assert.equal(t.firstName('jo@example.com'),'');
console.log('Draft tools: HTML formatting, send-safety checks, security guard, history and names passed.');
