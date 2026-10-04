/* gf.js - in-page Google Flights engine. No API, no keys: it runs INSIDE a google.com/travel/flights tab
 * and reads the same data the page itself renders (the ds:1 data block of the results HTML).
 * Inject once per tab (paste this whole file into the browser JS tool; check first with:  window.GF && GF.v), then drive it:
 *   GF.prime()                                   -> ONCE per tab: loads the full list like a person (click 'View more flights')
 *   GF.check()                                   -> sanity (right origin, ILS)
 *   GF.plan({from,to,minNights,maxNights})       -> pair count + time estimate (no network)
 *   GF.start('exact',{origin,dest,dep,ret,expectN})  -> background job, then poll GF.status()   (expectN = 'N results returned' shown by the page)
 *   GF.start('scan',{origin,dest,from,to,minNights,maxNights})
 *   GF.start('exactMany',{origin,dest,pairs:[[dep,ret],...]})   (deep dive on several pairs)
 *   GF.status() / GF.result(i) / GF.parts()      -> compact JSON text (chunks of 28000 chars)
 * origin/dest: IATA (TLV), "/m/..." id, or an English city name ("Budapest", "London").
 */
(() => {
  const main = () => {
  const G = (window.GF = window.GF || {});
  G.v = 8; G.fails = 0;
  const E = new TextEncoder(), sleep = ms => new Promise(r => setTimeout(r, ms));
  const vi = n => { const o = []; while (n > 127) { o.push((n & 127) | 128); n >>>= 7; } o.push(n); return o; };
  const ld = (f, b) => [...vi((f << 3) | 2), ...vi(b.length), ...b];
  const S = (f, s) => ld(f, [...E.encode(s)]);
  const V = (f, n) => [...vi(f << 3), ...vi(n)];
  const pl = (f, p) => ld(f, [...V(1, p.t), ...S(2, p.id)]);
  const b64 = a => btoa(String.fromCharCode(...a)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const P2 = n => String(n).padStart(2, '0');
  const D = a => (a ? `${a[0]}-${P2(a[1])}-${P2(a[2])}` : null);
  const T = a => (a ? `${P2(a[0] || 0)}:${P2(a[1] || 0)}` : null);

  // tfs = Google's own URL state. legs: [{date,from,to,sel:[{from,date,to,air,num}]}]
  G.tfs = legs => b64([...V(1, 28), ...V(2, 2),
    ...legs.flatMap(l => ld(3, [...S(2, l.date),
      ...(l.sel || []).flatMap(x => ld(4, [...S(1, x.from), ...S(2, x.date), ...S(3, x.to), ...S(5, x.air), ...S(6, x.num)])),
      ...pl(13, l.from), ...pl(14, l.to)])),
    ...V(8, 1), ...V(9, 1), ...V(14, 1), 0x82, 0x01, 0x0b, 0x08, ...Array(9).fill(255), 1, ...V(19, 1)]);
  const url = legs => `https://www.google.com/travel/flights/search?tfs=${G.tfs(legs)}&tfu=EgYIABABGAA&hl=en&curr=ILS`;
  const fl = d => [...((d[2] && d[2][0]) || []), ...((d[3] && d[3][0]) || [])];

  // Pull the ds:1 block out of the HTML. Fast path: cut at ", sideChannel"; fallback: bracket matching.
  const ds1 = t => {
    const i = t.indexOf("key: 'ds:1'"); if (i < 0) return null;
    const s = t.indexOf('data:', i) + 5, e = t.indexOf(', sideChannel', s);
    try { return JSON.parse(t.slice(s, e)); } catch (x) { /* fall through */ }
    let depth = 0, j = s, inS = false;
    for (; j < t.length; j++) {
      const c = t[j];
      if (inS) { if (c === String.fromCharCode(92)) j++; else if (c === '"') inS = false; }
      else if (c === '"') inS = true;
      else if (c === '[' || c === '{') depth++;
      else if (c === ']' || c === '}') { depth--; if (depth === 0) { j++; break; } }
    }
    return JSON.parse(t.slice(s, j));
  };
  // Google sometimes answers with a light error page (d[0]===13, ~1.8MB, no flights): that is retried, not trusted.
  G.get = async u => {
    for (let k = 0; k < 4; k++) {
      try {
        const t = await fetch(u, { credentials: 'include' }).then(r => r.text());
        const d = ds1(t);
        if (d && d[0] !== 13 && (d[2] !== undefined || d[3] !== undefined)) { G.fails = 0; return { d, ils: t.includes('₪') }; }
      } catch (e) { /* retry */ }
      G.fails++; await sleep(1200 * (k + 1));
    }
    return null;
  };

  // Google serves the short list (about 11) until the page itself has loaded the full list once ("View more flights").
  // prime() clicks that button the way a person would, so that every later request returns the full list.
  G.prime = async () => {
    const find = () => [...document.querySelectorAll('button,[role=button]')].find(x => /View more flights/i.test(x.innerText || x.getAttribute('aria-label') || ''));
    const b = find();
    const n = () => { const m = document.body.innerText.match(/(\d+) results? returned/); return m ? +m[1] : null; };
    if (!b) return JSON.stringify({ clicked: false, n: n(), note: 'no View-more button: the list is already complete (or this is not a results page)' });
    b.click();
    for (let i = 0; i < 40; i++) { await sleep(500); if (n() > 11) break; }
    return JSON.stringify({ clicked: true, n: n() });
  };

  // one direction = one record (legs, layovers, total minutes, price as given by Google)
  const rec = x => {
    const i = x[0];
    const legs = i[2].map(l => [l[3], l[6], D(l[20]), T(l[8]), D(l[21]), T(l[10]), l[11], l[22][0], l[22][1], l[17] || '', l[22][3] || '']);
    const lay = (i[13] || []).map(y => [y[0], y[1], y[2], y[5] || '']);
    return { id: legs.map(l => l[7] + l[8]).join('+') + '@' + legs[0][2] + 'T' + legs[0][3], al: (i[1] || []).join(', '),
      legs, lay, tot: i[9], price: x[1] && x[1][0] ? x[1][0][1] : null };
  };

  // resolve names -> Google place ids (city ids cover every airport of the city)
  G.resolve = async (origin, dest, date) => {
    const isId = p => /^[A-Z]{3}$/.test(p) ? { t: 1, id: p } : /^\/[mg]\//.test(p) ? { t: 2, id: p } : null;
    let o = isId(origin), d = isId(dest), title = '';
    if (!o || !d) {
      const q = `Flights from ${isId(origin) ? origin : origin} to ${dest} on ${date}`;
      const t = await fetch(`https://www.google.com/travel/flights?q=${encodeURIComponent(q)}&curr=ILS&hl=en`, { credentials: 'include' }).then(r => r.text());
      const m = t.match(/\["([^"]+)",(\d)\]\]\],\[\[\["([^"]+)",(\d)\]\]\]/);
      title = (t.match(/<title>([^<]*)/) || [])[1] || '';
      if (!m) throw new Error('cannot resolve place: ' + origin + ' / ' + dest);
      const mp = (id, ty) => ({ t: ty === '0' ? 1 : 2, id });
      o = o || mp(m[1], m[2]); d = d || mp(m[3], m[4]);
    }
    return { o, d, title };
  };

  G.plan = ({ from, to, minNights = 3, maxNights = 10, depDow = null, retDow = null }) => {
    const u = s => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)), f = ms => new Date(ms).toISOString().slice(0, 10);
    const pairs = [], day = 864e5;
    for (let a = u(from); a <= u(to); a += day) for (let n = minNights; n <= maxNights; n++) {
      const b = a + n * day; if (b > u(to)) break;
      if (depDow && !depDow.includes(new Date(a).getUTCDay())) continue;
      if (retDow && !retDow.includes(new Date(b).getUTCDay())) continue;
      pairs.push([f(a), f(b)]);
    }
    return { pairs, n: pairs.length, estSec: Math.round(pairs.length * 0.7 + 8) };
  };

  const pool = async (items, conc, fn, job) => {
    let k = 0;
    await Promise.all(Array.from({ length: conc }, async () => {
      while (k < items.length && job.state === 'running') {
        const it = items[k++]; await sleep(80 + Math.random() * 220);
        await fn(it); job.done++;
        if (G.fails >= 8) { job.state = 'blocked'; job.err = 'Google stopped answering (rate limit/captcha?). Wait, then resume with a slower pace.'; }
      }
    }));
  };

  const ex = async (o, job) => {
    const P = job.place || (job.place = await G.resolve(o.origin, o.dest, o.dep));
    const a = { date: o.dep, from: P.o, to: P.d }, b = { date: o.ret, from: P.d, to: P.o };
    const base = await G.get(url([a, b]));
    if (!base) throw new Error('no data for ' + o.dep + '..' + o.ret);
    const all = fl(base.d).map(rec).sort((x, y) => (x.price ?? 1e9) - (y.price ?? 1e9));
    const cap = o.maxOut || 40;   // every nonstop outbound, then the cheapest others up to the cap
    const outs = [...all.filter(x => x.legs.length === 1), ...all.filter(x => x.legs.length > 1)].filter((x, i) => x.legs.length === 1 || i < cap);
    const out = {}, ret = {}, itins = [], warn = [];
    if (!base.ils) warn.push('currency symbol not shown: check that prices are in ILS');
    if (o.expectN && all.length < o.expectN) warn.push('partial outbound list: got ' + all.length + ' of ' + o.expectN + ' shown on the page (run GF.prime() and repeat)');
    job.total += outs.length;
    await pool(outs, o.conc || 3, async x => {
      out[x.id] = x;
      if (x.price == null) { itins.push([x.id, null, null]); return; }
      const sel = x.legs.map(l => ({ from: l[0], date: l[2], to: l[1], air: l[7], num: l[8] }));
      const r = await G.get(url([{ ...a, sel }, b]));
      const R = r ? fl(r.d).map(rec) : [];
      if (!R.length) { warn.push('no return options for ' + x.id); return; }
      R.forEach(y => { ret[y.id] = y; itins.push([x.id, y.id, y.price]); });
    }, job);
    return { dep: o.dep, ret: o.ret, found: all.length, used: outs.length, out, ret_: ret, itins, warn };
  };

  // compact wire format (keeps what the agent must copy to disk small):
  //   out/ret_ : arrays of [legs, layovers, totalMinutes];  itins: [outIndex, retIndex(-1 = none), totalPriceILS|null]
  //   leg      : [from,to,depMMDD,depHHMM,arrMMDD,arrHHMM,minutes,airlineCode,flightNo]
  //   layover  : [minutes, arrivalAirport, departureAirport]
  const NM = {}, CITY = {};
  const cz = r => {
    r.legs.forEach(l => { NM[l[7]] = l[10] || NM[l[7]] || l[7]; });
    r.lay.forEach(y => { CITY[y[1]] = y[3]; });
    return [r.legs.map(l => [l[0], l[1], l[2].slice(5).replace('-', ''), l[3].replace(':', ''), l[4].slice(5).replace('-', ''), l[5].replace(':', ''), l[6], l[7], l[8]]),
      r.lay.map(y => [y[0], y[1], y[2]]), r.tot];
  };
  const trim = (e, cap) => {
    const nonstop = k => e.out[k[0]].legs.length === 1 && e.ret_[k[1]] && e.ret_[k[1]].legs.length === 1;
    const priced = e.itins.filter(k => k[2] != null).sort((x, y) => x[2] - y[2]).filter((k, _, a) => k[2] <= a[0][2] * 5); // drop >5x the cheapest (noise)
    const keep = [...priced.filter(nonstop), ...priced.filter(k => !nonstop(k))].slice(0, cap)
      .concat(e.itins.filter(k => k[2] == null).slice(0, 10));
    const O = [], R = [], oi = new Map(), ri = new Map();
    const ix = (m, arr, id, r) => { if (!m.has(id)) { m.set(id, arr.length); arr.push(cz(r)); } return m.get(id); };
    const dd = new Map();  // same out+ret pair can come back more than once (fare variants): keep the cheapest
    keep.map(k => [ix(oi, O, k[0], e.out[k[0]]), k[1] ? ix(ri, R, k[1], e.ret_[k[1]]) : -1, k[2]])
      .forEach(t => { const id = t[0] + '_' + t[1], old = dd.get(id); if (!old || (t[2] != null && (old[2] == null || t[2] < old[2]))) dd.set(id, t); });
    const its = [...dd.values()];
    return { dep: e.dep, ret: e.ret, outFound: e.found, outUsed: e.used, itinFound: e.itins.length, out: O, ret_: R, itins: its, warn: e.warn.slice(0, 20) };
  };
  const wrap = (kind, job, extra) => ({ kind, at: new Date().toISOString(), place: job.place, names: NM, cities: CITY, ...extra });

  const kinds = {
    async exact(o, job) { job.total = 1; job.done = 1; job.res = wrap('exact', job, { exacts: [trim(await ex(o, job), o.maxItins || 120)] }); },
    async exactMany(o, job) {
      const exs = [];
      for (const [dep, ret] of o.pairs) { exs.push(trim(await ex({ ...o, dep, ret }, job), o.maxItins || 120)); if (job.state !== 'running') break; }
      job.res = wrap('exactMany', job, { exacts: exs });
    },
    async scan(o, job) {
      const pl_ = o.pairs ? { pairs: o.pairs } : G.plan(o);
      job.place = await G.resolve(o.origin, o.dest, pl_.pairs[0][0]);
      job.total = pl_.pairs.length; const rows = [], failed = [];
      const a = (d, from, to) => ({ date: d, from, to });
      await pool(pl_.pairs, o.conc || 3, async ([dep, ret]) => {
        const r = await G.get(url([a(dep, job.place.o, job.place.d), a(ret, job.place.d, job.place.o)]));
        if (!r) { failed.push([dep.slice(5).replace('-', ''), ret.slice(5).replace('-', '')]); return; }
        const R = fl(r.d).map(rec), pr = R.filter(x => x.price != null), ns = pr.filter(x => x.legs.length === 1);
        const m = xs => xs.length ? Math.min(...xs.map(x => x.price)) : null;
        // row = [dep, ret, cheapest total, cheapest total with a NONSTOP OUTBOUND, #outbound options]
        rows.push([dep.slice(5).replace('-', ''), ret.slice(5).replace('-', ''), m(pr), m(ns), R.length]);
      }, job);
      rows.sort((x, y) => (x[2] ?? 1e9) - (y[2] ?? 1e9) || (x[0] + x[1]).localeCompare(y[0] + y[1]));
      job.res = wrap('scan', job, { year: +pl_.pairs[0][0].slice(0, 4), planned: pl_.pairs.length, failed, pairs: rows });
    },
  };

  G.start = (kind, o) => {
    const job = (G.job = { kind, state: 'running', done: 0, total: 0, t0: Date.now(), res: null, err: null, place: null });
    kinds[kind](o, job).then(() => { if (job.state === 'running') job.state = 'done'; job.sec = Math.round((Date.now() - job.t0) / 1000); })
      .catch(e => { job.state = 'error'; job.err = String(e && e.message || e); });
    return 'started ' + kind;
  };
  G.status = () => { const j = G.job; return j ? JSON.stringify({ kind: j.kind, state: j.state, done: j.done, total: j.total, sec: Math.round((Date.now() - j.t0) / 1000), err: j.err, parts: j.res ? G.parts() : 0 }) : 'no job'; };
  // self-verifying text: {"chk":"<fnv1a of the whole text with chk=00000000>", ...}; build_excel.py re-checks it,
  // so a mistake while copying the result into a file is detected instead of silently corrupting the Excel.
  const txt = () => {
    const j = G.job; if (j._t) return j._t;
    const t = JSON.stringify({ chk: '00000000', ...j.res });
    let h = 2166136261; for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
    return (j._t = t.replace('00000000', h.toString(16).padStart(8, '0')));
  };
  G.parts = () => Math.ceil(txt().length / 28000);
  G.result = (i = 0) => txt().slice(i * 28000, (i + 1) * 28000);
  G.check = () => JSON.stringify({ origin: location.origin, ok: location.origin === 'https://www.google.com', v: G.v });
  return 'GF ready v' + G.v;
  };
  return main();
})();
