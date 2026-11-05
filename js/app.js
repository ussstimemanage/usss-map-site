"use strict";
// 公演データのExcel（index.html と同じ場所に置く）。更新するときはこのファイルを差し替える
const DATA_FILE = 'USSS_Live_Archive.xlsx';

/* ---------- 定数 ---------- */
const C = 60, G = 3;
const MEMBERS_DEFAULT = ['うらたぬき','志麻','となりの坂田。','センラ'];
const MEMBER_STYLE = {'うらたぬき':['U','urata'],'志麻':['S','shima'],'となりの坂田。':['S','sakata'],'センラ':['S','senra']};
// 絞り込み「種類」のボタン。幅に余裕があればこの4行のまま固定し、
// どれか1行でも入りきらない幅になったら、同じ順番のまま幅に合わせて折り返す。ここにない種類はボタンを出さない
const TYPE_ROWS = [
  ['春ツアー','夏ツアー','ハロパ','しまさかBD','ワンマンライブ'],
  ['そまうさ','XYZ','ETA','スタプラ','ハニワ','裏留守華音'],
  ['フェス','ゲスト出演','ツーマン','単発','ファンミ','イベント'],
  ['海外公演','代替公演']
];
const SPECIAL = new Set(['overseas','online']);
const REGIONS = [
  {id:'hokkaido',name:'北海道'},{id:'tohoku',name:'東北'},{id:'kanto',name:'関東'},{id:'chubu',name:'中部'},
  {id:'kinki',name:'近畿'},{id:'chugoku',name:'中国'},{id:'shikoku',name:'四国'},{id:'kyushu',name:'九州・沖縄'},
  {id:'overseas',name:'海外',special:true},{id:'online',name:'配信',special:true}
];
// [id, 略称, エリア, 列, 行, 幅, 高さ]
const PREF_DEFS = [
  ['hokkaido','北海道','hokkaido',11,0,2,2],
  ['aomori','青森','tohoku',11,2],['akita','秋田','tohoku',10,3],['iwate','岩手','tohoku',11,3],
  ['yamagata','山形','tohoku',10,4],['miyagi','宮城','tohoku',11,4],['fukushima','福島','tohoku',10,5,2,1],
  ['ibaraki','茨城','kanto',12,6,1,2],['tochigi','栃木','kanto',11,6],['gunma','群馬','kanto',10,6],
  ['saitama','埼玉','kanto',10,7],['chiba','千葉','kanto',11,7,1,2],['tokyo','東京','kanto',10,8],['kanagawa','神奈川','kanto',10,9],
  ['niigata','新潟','chubu',9,4,1,2],['toyama','富山','chubu',8,5],['ishikawa','石川','chubu',7,4,1,2],['fukui','福井','chubu',7,6],
  ['yamanashi','山梨','chubu',9,7],['nagano','長野','chubu',9,6],['gifu','岐阜','chubu',8,6],['shizuoka','静岡','chubu',9,8],['aichi','愛知','chubu',8,7],
  ['mie','三重','kinki',8,8],['shiga','滋賀','kinki',7,7],['kyoto','京都','kinki',6,6,1,2],['osaka','大阪','kinki',6,8],
  ['hyogo','兵庫','kinki',5,6,1,2],['nara','奈良','kinki',7,8],['wakayama','和歌山','kinki',6,9,2,1],
  ['tottori','鳥取','chugoku',4,6],['shimane','島根','chugoku',3,6],['okayama','岡山','chugoku',4,7],['hiroshima','広島','chugoku',3,7],['yamaguchi','山口','chugoku',2,7],
  ['tokushima','徳島','shikoku',4,10],['kagawa','香川','shikoku',4,9],['ehime','愛媛','shikoku',3,9],['kochi','高知','shikoku',3,10],
  ['fukuoka','福岡','kyushu',1,8],['saga','佐賀','kyushu',0,8],['nagasaki','長崎','kyushu',0,9],['kumamoto','熊本','kyushu',1,9],
  ['oita','大分','kyushu',2,8],['miyazaki','宮崎','kyushu',2,9],['kagoshima','鹿児島','kyushu',1,10],['okinawa','沖縄','kyushu',0,12],
  ['overseas','海外','overseas',8,11,2,1],['online','配信','online',11,11,2,1]
].map(([id,short,region,x,y,w=1,h=1]) => ({id,short,region,x,y,w,h,
  full: SPECIAL.has(id) || id==='hokkaido' ? short : id==='tokyo' ? '東京都' : (id==='osaka'||id==='kyoto') ? short+'府' : short+'県'}));
const PREF = new Map(PREF_DEFS.map(p => [p.id,p]));
const PREF_LOOKUP = new Map();
PREF_DEFS.forEach(p => { PREF_LOOKUP.set(p.full,p.id); PREF_LOOKUP.set(p.short,p.id); });
['国外','overseas','Overseas','OVERSEAS'].forEach(k => PREF_LOOKUP.set(k,'overseas'));
['オンライン','配信ライブ','online','Online','ONLINE'].forEach(k => PREF_LOOKUP.set(k,'online'));
const REGION = new Map(REGIONS.map(r => [r.id,r]));

/* ---------- データ読み込み（Excel → 内部形式） ---------- */
// PARSE-START
function norm(s){ return String(s ?? '').normalize('NFKC').replace(/\r\n?/g,'\n').trim(); }
function pick(row, names){
  for (const n of names){ if (row[n] !== undefined && String(row[n]).trim() !== '') return row[n]; }
  return '';
}
function prefIdOf(s){ const t = norm(s); return PREF_LOOKUP.get(t) || PREF_LOOKUP.get(t.replace(/[県府]$/,'')) || null; }
function parseDate(v, XLSXlib){
  if (v instanceof Date && !isNaN(v)){ const d = new Date(v.getTime() + 12*3600e3); return ymd(d.getFullYear(), d.getMonth()+1, d.getDate()); }
  if (typeof v === 'number' && XLSXlib){ const o = XLSXlib.SSF.parse_date_code(v); return o ? ymd(o.y,o.m,o.d) : ''; }
  const m = norm(v).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? ymd(+m[1],+m[2],+m[3]) : '';
}
// 開演時刻：Excelの時刻(小数)・Date・文字列（2行書きは「夜\n昼」のまま保持）
function parseTime(v){
  const hm = (h,m) => `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  if (v instanceof Date && !isNaN(v)) return hm(v.getHours(), v.getMinutes());
  if (typeof v === 'number'){ const t = Math.round((v % 1) * 1440); return hm(Math.floor(t/60) % 24, t % 60); }
  return norm(v).split('\n').map(x => x.trim()).filter(Boolean).join('\n');
}
function ymd(y,m,d){ return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`; }
function splitList(v){ return norm(v).split(/[,、，\/／・\n]+/).map(s=>s.trim()).filter(Boolean); }
function parseSeats(v){
  return String(v ?? '').split(/\n|\|/).map(s=>s.trim()).filter(Boolean).map(s => {
    const i = s.indexOf('=');
    return i > 0 ? {label:s.slice(0,i).trim(), src:s.slice(i+1).trim()} : {label:'', src:s};
  });
}
// venueRows / showRows: 1行=1オブジェクト（見出し→値）
function buildData(venueRows, showRows, XLSXlib){
  const warnings = [];
  const venues = new Map(), byName = new Map();
  venueRows.forEach((r,i) => {
    const id = norm(pick(r,['会場ID','ID','id']));
    const name = norm(pick(r,['会場名','会場']));
    if (!id && !name) return;
    const prefRaw = pick(r,['都道府県','県']);
    const pref = prefIdOf(prefRaw);
    if (!pref) warnings.push(`会場シート ${i+2}行目：都道府県「${norm(prefRaw)}」を判別できません`);
    const capRaw = norm(pick(r,['最大キャパ','キャパ','収容人数']));
    const v = { id: id || name, name: name || id, pref,
      address: norm(pick(r,['住所'])), capacity: /^[\d,\s]+(人|席|名)?$/.test(capRaw) ? Number(capRaw.replace(/[^\d]/g,'')) : capRaw,
      access: norm(pick(r,['アクセス'])), seats: parseSeats(pick(r,['座席表画像','座席表'])),
      url: norm(pick(r,['公式サイト','URL','url'])), memo: norm(pick(r,['メモ','備考'])) };
    venues.set(v.id, v); byName.set(v.name, v);
  });
  const shows = [];
  showRows.forEach((r,i) => {
    const date = parseDate(pick(r,['日付','開催日','公演日']), XLSXlib);
    const title = norm(pick(r,['公演名','ツアー・イベント名','タイトル']));
    if (!date && !title) return;
    if (!date){ warnings.push(`公演シート ${i+2}行目：日付が読めません`); return; }
    const vid = norm(pick(r,['会場ID']));
    const vname = norm(pick(r,['会場名','会場']));
    const venue = venues.get(vid) || byName.get(vname) || byName.get(vid);
    if (!venue){ warnings.push(`公演シート ${i+2}行目：会場「${vid || vname}」が会場シートにありません`); return; }
    let members = splitList(pick(r,['出演メンバー','メンバー','出演']));
    if (members.length === 1 && members[0] === '全員') members = MEMBERS_DEFAULT.slice();
    const note = norm(pick(r,['備考','メモ']));
    // 「開催中止」は0公演、「昼夜2公演」は2公演として数える
    const cancelled = note.includes('開催中止');
    const weight = cancelled ? 0 : (/昼夜\s*2\s*公演/.test(note) ? 2 : 1);
    shows.push({ id: norm(pick(r,['公演ID','ID'])) || `row${i+2}`, date, year: date.slice(0,4), title,
      type: norm(pick(r,['種類','ライブの種類','区分'])) || 'その他', venueId: venue.id, members,
      start: parseTime(pick(r,['開演','開演時間'])), note, cancelled, weight });
  });
  return {venues, shows, warnings};
}
function fromWorkbook(wb, XLSXlib){
  const find = kw => wb.SheetNames.find(n => n.includes(kw));
  const vs = find('会場'), ss = find('公演');
  if (!vs || !ss) throw new Error('「公演」と「会場」という名前のシートが見つかりません');
  const opt = {defval:'', raw:true};
  return buildData(XLSXlib.utils.sheet_to_json(wb.Sheets[vs],opt), XLSXlib.utils.sheet_to_json(wb.Sheets[ss],opt), XLSXlib);
}
// PARSE-END

// サンプル用の座席図（Excel側では images/xxx.png のような画像パスやURLを書きます）
function seatSrc(src){
  if (!src.startsWith('SEAT:')) return src;
  const label = src.slice(5);
  let blocks = '';
  for (let r=0;r<3;r++) for (let c=0;c<4;c++){
    const x=70+c*68, y=72+r*48;
    blocks += `<rect x="${x}" y="${y}" width="60" height="40" rx="3" fill="#cfe0ea" stroke="#9fb8c8"/><text x="${x+30}" y="${y+25}" font-size="13" text-anchor="middle" fill="#14293b" font-family="sans-serif">A${r*4+c+1}</text>`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect width="400" height="300" fill="#f6f8fa"/><rect x="130" y="16" width="140" height="32" rx="4" fill="#14293b"/><text x="200" y="37" font-size="14" fill="#fff" text-anchor="middle" font-family="sans-serif">STAGE</text>${blocks}<rect x="14" y="60" width="44" height="180" rx="4" fill="#ecdcb4"/><rect x="342" y="60" width="44" height="180" rx="4" fill="#ecdcb4"/><text x="36" y="154" font-size="12" text-anchor="middle" fill="#5a4a20" font-family="sans-serif">西</text><text x="364" y="154" font-size="12" text-anchor="middle" fill="#5a4a20" font-family="sans-serif">東</text><text x="200" y="284" font-size="12" fill="#566a7a" text-anchor="middle" font-family="sans-serif">${label}（サンプル図）</text></svg>`;
  return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
}

/* ---------- 状態 ---------- */
const state = {
  data: null, source: 'sample',
  f: {q:'', year:'', types:new Set(), members:new Set()},
  region: null, openPrefs: new Set(), limit: 60, filtered: []
};
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmtDate = d => d.replace(/-/g,'.');
const venueOf = s => state.data.venues.get(s.venueId);
const regionOfShow = s => { const v = venueOf(s); return v && v.pref ? PREF.get(v.pref).region : null; };
const sumW = list => list.reduce((a,s) => a + s.weight, 0);
const noteHtml = n => esc(n).replace(/開催中止/g, '<span class="cancel">開催中止</span>');
// 今年のデータがあれば今年、なければデータ内でいちばん新しい年
function defaultYear(d){
  const ys = new Set(d.shows.map(s => s.year)), cur = String(new Date().getFullYear());
  return ys.has(cur) ? cur : ([...ys].sort().pop() || '');
}
let allMembers = MEMBERS_DEFAULT.slice();

function setData(d, source, label){
  d.shows.forEach(s => {
    const v = d.venues.get(s.venueId);
    const pf = v && v.pref ? PREF.get(v.pref) : null;
    s._hay = [s.title,s.type,s.members.join(' '),s.note,s.year,v&&v.name,v&&v.address,pf&&pf.full].join(' ').normalize('NFKC').toLowerCase();
  });
  d.shows.sort((a,b) => b.date.localeCompare(a.date) || a.id.localeCompare(b.id));
  state.data = d; state.source = source;
  const extra = [...new Set(d.shows.flatMap(s => s.members))].filter(m => !MEMBERS_DEFAULT.includes(m));
  allMembers = MEMBERS_DEFAULT.concat(extra);
  state.f = {q:'', year:defaultYear(d), types:new Set(), members:new Set()}; $('#q').value = '';
  $('#dataStatusText').textContent = label;
  $('#dataStatus').classList.toggle('ok', source !== 'sample');
  const wb = $('#warnBox');
  if (d.warnings.length){
    wb.hidden = false;
    wb.innerHTML = `<details class="warn"><summary>読み込み時の注意 ${d.warnings.length}件（該当行は表示していません）</summary><ul>${d.warnings.slice(0,50).map(w=>`<li>${esc(w)}</li>`).join('')}</ul></details>`;
  } else { wb.hidden = true; wb.innerHTML = ''; }
  buildFilters(); render(); fitMap();
  if (typeof refreshDatePage === 'function') refreshDatePage();
}

/* ---------- 絞り込み ---------- */
function matches(s){
  const f = state.f;
  if (f.year && s.year !== f.year) return false;
  if (f.types.size && !f.types.has(s.type)) return false;
  for (const m of f.members) if (!s.members.includes(m)) return false;
  if (f.q){ const words = f.q.normalize('NFKC').toLowerCase().split(/\s+/).filter(Boolean); if (!words.every(w => s._hay.includes(w))) return false; }
  return true;
}
function buildFilters(){
  const d = state.data;
  const years = [...new Set(d.shows.map(s => s.year))].sort().reverse();
  $('#year').innerHTML = '<option value="">すべて</option>' + years.map(y => `<option value="${y}">${y}年</option>`).join('');
  const present = new Set(d.shows.map(s => s.type));
  const chip = t => `<button type="button" class="chip${present.has(t) ? '' : ' none'}" data-type="${esc(t)}" aria-pressed="false">${esc(t)}</button>`;
  $('#typeChips').innerHTML = TYPE_ROWS.map(row => `<div class="trow">${row.map(chip).join('')}</div>`).join('');
  layoutTypes();
  $('#memberChips').innerHTML = allMembers.map(m => `<button type="button" class="chip" data-member="${esc(m)}" aria-pressed="false">${MEMBER_STYLE[m] ? `<span class="cdot" style="--mc:var(--m-${MEMBER_STYLE[m][1]})"></span>` : ''}${esc(m)}</button>`).join('');
}
function syncChips(){
  document.querySelectorAll('[data-type]').forEach(b => b.setAttribute('aria-pressed', state.f.types.has(b.dataset.type)));
  document.querySelectorAll('[data-member]').forEach(b => b.setAttribute('aria-pressed', state.f.members.has(b.dataset.member)));
  $('#year').value = state.f.year;
  const n = state.f.types.size + state.f.members.size, fc = $('#fcount');
  fc.hidden = !n; fc.textContent = n ? n + '件' : '';
}

/* ---------- 地図 ---------- */
const svg = $('#map');
const FULL = [-10,-10,800,800];
let vb = FULL.slice(), raf = 0;
function buildMap(){
  const NS = 'http://www.w3.org/2000/svg';
  let h = `<defs><filter id="glow" x="-30%" y="-30%" width="160%" height="160%"><feGaussianBlur stdDeviation="5" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>`;
  // 方位記号と沖縄の囲み
  h += `<g aria-hidden="true"><circle class="deco" cx="110" cy="110" r="56" stroke-width="1"/><circle class="deco" cx="110" cy="110" r="44" stroke-width=".6" stroke-dasharray="2 4"/>
    <path d="M110 48 L117 110 L110 172 L103 110Z" fill="var(--muted)" opacity=".55"/><path d="M48 110 L110 104 L172 110 L110 116Z" fill="var(--muted)" opacity=".3"/>
    <text class="deco-t" x="110" y="40" text-anchor="middle">N</text>
    <rect class="deco" x="-6" y="${12*C-6}" width="${C+12}" height="${C+12}" rx="10" stroke-dasharray="4 4"/></g>`;
  for (const r of REGIONS){
    const prefs = PREF_DEFS.filter(p => p.region === r.id);
    let tiles = '';
    for (const p of prefs){
      const x = p.x*C+G, y = p.y*C+G, w = p.w*C-2*G, hh = p.h*C-2*G;
      const big = p.w*p.h > 1 && p.w === p.h;
      if (SPECIAL.has(p.id)){
        tiles += `<g class="tile special" data-pref="${p.id}"><title>${p.full}</title><rect x="${x}" y="${y}" width="${w}" height="${hh}" rx="9"/>
          <text class="sname" x="${x+14}" y="${y+hh/2+5}">${p.short}</text>
          <text class="pcount" x="${x+w-14}" y="${y+hh/2+7}" text-anchor="end" font-size="19"></text></g>`;
        continue;
      }
      tiles += `<g class="tile${big ? ' big' : ''}" data-pref="${p.id}"><title>${p.full}</title><rect x="${x}" y="${y}" width="${w}" height="${hh}" rx="9"/>
        <text class="pname" x="${x+w/2}" y="${y+15}" text-anchor="middle">${p.short}</text>
        <text class="pcount" x="${x+w/2}" y="${y+hh/2+(big?12:9)}" text-anchor="middle" font-size="${big?30:19}"></text></g>`;
    }
    // エリア外周の線分（同エリアのマス同士が接する辺は除く）
    const cells = new Set();
    prefs.forEach(p => { for (let i=0;i<p.w;i++) for (let j=0;j<p.h;j++) cells.add(`${p.x+i},${p.y+j}`); });
    let d = '';
    cells.forEach(k => {
      const [cx,cy] = k.split(',').map(Number), X = cx*C, Y = cy*C;
      if (!cells.has(`${cx},${cy-1}`)) d += `M${X} ${Y}H${X+C}`;
      if (!cells.has(`${cx},${cy+1}`)) d += `M${X} ${Y+C}H${X+C}`;
      if (!cells.has(`${cx-1},${cy}`)) d += `M${X} ${Y}V${Y+C}`;
      if (!cells.has(`${cx+1},${cy}`)) d += `M${X+C} ${Y}V${Y+C}`;
    });
    h += `<g class="region" data-region="${r.id}" style="--rc:var(--r-${r.id})" tabindex="0" role="button" aria-label="${r.name}">${tiles}<path class="glow" d="${d}"/><path class="core" d="${d}"/></g>`;
  }
  svg.innerHTML = h;
}
function regionBox(id){
  const ps = PREF_DEFS.filter(p => p.region === id);
  const x0 = Math.min(...ps.map(p=>p.x))*C, y0 = Math.min(...ps.map(p=>p.y))*C;
  const x1 = Math.max(...ps.map(p=>p.x+p.w))*C, y1 = Math.max(...ps.map(p=>p.y+p.h))*C;
  const size = Math.max(x1-x0, y1-y0) + C*0.9;
  return [(x0+x1)/2 - size/2, (y0+y1)/2 - size/2, size, size];
}
function animateVB(target){
  cancelAnimationFrame(raf);
  const from = vb.slice(), t0 = performance.now();
  const dur = matchMedia('(prefers-reduced-motion: reduce)').matches ? 0 : 560;
  const step = now => {
    const t = dur ? Math.min(1,(now-t0)/dur) : 1, e = 1 - Math.pow(1-t,3);
    vb = from.map((v,i) => v + (target[i]-v)*e);
    svg.setAttribute('viewBox', vb.map(n=>n.toFixed(2)).join(' '));
    if (t < 1) raf = requestAnimationFrame(step);
  };
  raf = requestAnimationFrame(step);
}
function selectRegion(id){
  state.region = id; state.openPrefs.clear();
  const special = id && REGION.get(id).special;
  if (special) state.openPrefs.add(id);
  animateVB(id && !special ? regionBox(id) : FULL);
  render();
}

/* ---------- 描画 ---------- */
function render(opts = {}){
  syncChips();
  const shows = state.data.shows.filter(matches);
  state.filtered = shows;
  const byPref = new Map();
  shows.forEach(s => { const v = venueOf(s); if (v && v.pref) byPref.set(v.pref, (byPref.get(v.pref)||0)+s.weight); });
  $('#resultN').textContent = sumW(shows);

  svg.classList.toggle('zoomed', !!state.region && !REGION.get(state.region).special);
  svg.querySelectorAll('.region').forEach(g => {
    const id = g.dataset.region;
    g.classList.toggle('active', state.region === id);
    g.classList.toggle('dim', !!state.region && state.region !== id);
    g.setAttribute('aria-pressed', state.region === id);
  });
  svg.querySelectorAll('.tile').forEach(t => {
    const n = byPref.get(t.dataset.pref) || 0;
    t.classList.toggle('empty', n === 0);
    t.classList.toggle('sel', state.openPrefs.has(t.dataset.pref));
    t.querySelector('.pcount').textContent = n ? n : '';
  });

  const byRegion = new Map();
  shows.forEach(s => { const r = regionOfShow(s); if (r) byRegion.set(r,(byRegion.get(r)||0)+s.weight); });
  $('#legend').innerHTML = REGIONS.map(r => `<button type="button" class="lg" data-region="${r.id}" style="--rc:var(--r-${r.id})" aria-current="${state.region===r.id}"><i></i>${r.name}<span>${byRegion.get(r.id)||0}</span></button>`).join('');

  $('#crumb').innerHTML = state.region
    ? `<button type="button" class="linkbtn" data-home>全国</button><span class="sep">›</span><strong>${REGION.get(state.region).name}</strong><button type="button" class="btn ghost" data-home>全国に戻る</button>`
    : `<strong>全国</strong><span class="sep" style="font-size:.8rem"><span class="hint-mouse">エリアにカーソルを合わせて、クリックで拡大</span><span class="hint-touch">エリアをタップすると拡大します</span></span>`;

  const panel = $('#panel'), key = state.region || '';
  const prevSc = panel.querySelector('.panel-scroll'), prevTop = prevSc ? prevSc.scrollTop : 0;
  const prevYs = panel.querySelector('.years'), prevLeft = prevYs ? prevYs.scrollLeft : 0;
  panel.innerHTML = state.region ? regionPanel(shows, byPref) : overviewPanel(shows);
  if (opts.keepScroll && panel.dataset.key === key){ const sc = panel.querySelector('.panel-scroll'); if (sc) sc.scrollTop = prevTop; }
  panel.dataset.key = key;
  // 年ボタンの横スクロール位置は描き直しても保つ。選択中の年が見えないときだけ、見える位置までずらす
  const ys = panel.querySelector('.years');
  if (ys){
    ys.scrollLeft = prevLeft;
    const on = ys.querySelector('[aria-pressed="true"]');
    if (on){
      const r = on.getBoundingClientRect(), box = ys.getBoundingClientRect();
      if (r.left < box.left) ys.scrollLeft -= box.left - r.left + 8;
      else if (r.right > box.right) ys.scrollLeft += r.right - box.right + 8;
    }
  }
}
function yearTabs(){
  const years = [...new Set(state.data.shows.map(s => s.year))].sort().reverse();
  return `<div class="years" role="group" aria-label="開催年の切り替え"><button type="button" class="chip" data-year="" aria-pressed="${!state.f.year}">すべて</button>${years.map(y => `<button type="button" class="chip" data-year="${y}" aria-pressed="${state.f.year===y}">${y}</button>`).join('')}</div>`;
}
function memberStats(shows){
  const badges = MEMBERS_DEFAULT.map(m => {
    const [ini, k] = MEMBER_STYLE[m];
    const n = sumW(shows.filter(s => s.members.includes(m)));
    return `<span class="mb" title="${esc(m)}：${n}公演"><i class="${k==='senra'?'lt':''}" style="--mc:var(--m-${k})" aria-hidden="true">${ini}</i><span class="vh">${esc(m)}</span><b>${n}</b></span>`;
  }).join('');
  // 4人分をひとまとまりにして、入りきらないときは4人そろって次の行へ
  return `<span class="mgroup">${badges}</span>`;
}
function memberLabel(ms){
  if (MEMBERS_DEFAULT.every(m => ms.includes(m)) && ms.length === MEMBERS_DEFAULT.length) return '4人';
  return ms.join('・') || '—';
}
function showRow(s){
  const v = venueOf(s), r = regionOfShow(s), pf = v.pref ? PREF.get(v.pref) : null;
  return `<li><button type="button" class="show" data-venue="${esc(v.id)}">
    <time datetime="${s.date}">${fmtDate(s.date)}</time>
    <span class="t">${esc(s.title)}${s.note ? ` <small style="color:var(--muted);font-weight:400">${noteHtml(s.note)}</small>` : ''}</span>
    <span class="m"><span><span class="rdot" style="--rc:var(--r-${r||'chubu'})"></span>${esc(v.name)}${pf ? `（${pf.short}）` : ''}</span><span class="tag">${esc(s.type)}</span><span>${esc(memberLabel(s.members))}</span></span>
  </button></li>`;
}
function emptyMsg(){ return `<p class="empty-msg">条件に合う公演がありません。<button type="button" class="linkbtn" data-clear>条件をクリア</button></p>`; }
function overviewPanel(shows){
  const live = shows.filter(s => s.weight > 0);
  const vset = new Set(live.map(s => s.venueId));
  const pset = new Set(live.map(s => venueOf(s).pref).filter(p => p && !SPECIAL.has(p)));
  return `<div class="panel-head">
      <div><p class="eyebrow">公演記録</p><h2>${state.f.year ? state.f.year+'年の公演' : 'すべての年の公演'}</h2></div>
      ${yearTabs()}
      <div class="stats"><span><b>${sumW(shows)}</b>公演</span><span><b>${vset.size}</b>会場</span><span><b>${pset.size}</b>都道府県</span>${memberStats(shows)}</div>
    </div>
    <div class="panel-scroll" tabindex="0" aria-label="公演一覧">${shows.length ? `<ol class="list">${shows.map(showRow).join('')}</ol>` : emptyMsg()}</div>`;
}
function venueBlocks(shows, prefId){
  const groups = new Map();
  shows.forEach(s => { const v = venueOf(s); if (v.pref === prefId){ if (!groups.has(v.id)) groups.set(v.id,[]); groups.get(v.id).push(s); } });
  return [...groups.entries()].sort((a,b) => sumW(b[1]) - sumW(a[1])).map(([vid,list]) => {
    const v = state.data.venues.get(vid);
    return `<div class="venue"><button type="button" class="venue-btn" data-venue="${esc(vid)}"><b>${esc(v.name)}</b><small>${sumW(list)}公演${typeof v.capacity === 'number' ? `・最大${v.capacity.toLocaleString('ja-JP')}人` : ''}</small></button>
      <ul class="dates">${list.map(s => `<li><time datetime="${s.date}">${fmtDate(s.date)}</time><span>${esc(s.title)}${s.note ? `（${noteHtml(s.note)}）` : ''}</span></li>`).join('')}</ul></div>`;
  }).join('');
}
function regionPanel(shows, byPref){
  const r = REGION.get(state.region);
  const prefs = PREF_DEFS.filter(p => p.region === r.id);
  const total = prefs.reduce((a,p) => a + (byPref.get(p.id)||0), 0);
  const inRegion = shows.filter(s => regionOfShow(s) === r.id);
  let body;
  if (r.special){
    const blocks = venueBlocks(shows, r.id);
    body = blocks ? `<div class="pref-body" style="padding:12px 0 0">${blocks}</div>` : emptyMsg();
  } else {
    body = '<div>' + prefs.map(p => {
      const n = byPref.get(p.id) || 0, has = inRegion.some(s => venueOf(s).pref === p.id);
      const open = state.openPrefs.has(p.id) && has;
      return `<div class="pref" id="pref-${p.id}"><button type="button" class="pref-btn" data-pref="${p.id}" aria-expanded="${open}" ${has?'':'disabled'}>
        <span class="chev" aria-hidden="true">${has ? '›' : '·'}</span>${p.full}<span class="n">${has ? n+'公演' : '公演なし'}</span></button>${open ? `<div class="pref-body">${venueBlocks(shows, p.id)}</div>` : ''}</div>`;
    }).join('') + '</div>';
  }
  return `<div class="panel-head">
      <div><p class="eyebrow">${r.special ? '公演記録' : 'エリア'}</p><h2 style="--rc:var(--r-${r.id})"><i></i>${r.name}${state.f.year ? `<small style="font-size:.8rem;color:var(--muted);font-family:var(--font)">${state.f.year}年</small>` : ''}</h2></div>
      ${yearTabs()}
      <div class="stats"><span><b>${total}</b>公演</span>${memberStats(inRegion)}</div>
      ${r.special ? '' : '<p class="eyebrow" style="letter-spacing:.04em">都道府県名を押すと会場が開きます</p>'}
    </div>
    <div class="panel-scroll" tabindex="0" aria-label="${r.name}の公演">${body}</div>`;
}

/* ---------- 会場詳細 ---------- */
function openVenue(id){
  const v = state.data.venues.get(id); if (!v) return;
  const pf = v.pref ? PREF.get(v.pref) : null;
  const all = state.data.shows.filter(s => s.venueId === id);
  const hit = new Set(state.filtered.filter(s => s.venueId === id));
  const cap = typeof v.capacity === 'number' ? v.capacity.toLocaleString('ja-JP') + '人' : (v.capacity || '—');
  const seats = v.seats.length ? `<div class="seats">${v.seats.map((s,i) => `<figure><button type="button" data-zoom="${i}" aria-label="${esc(s.label||'座席表')}を拡大"><img src="${esc(seatSrc(s.src))}" alt="${esc(v.name)} 座席表 ${esc(s.label)}" loading="lazy"></button><figcaption>${esc(s.label || '座席表')}</figcaption></figure>`).join('')}</div>`
    : `<p class="empty-msg" style="padding:0">座席表の登録はありません。</p>`;
  $('#dlgBody').innerHTML = `
    <div class="dlg-head"><div><p class="eyebrow">${pf ? pf.full : '都道府県不明'}${pf ? '・'+REGION.get(pf.region).name : ''}</p><h3 id="dlgTitle">${esc(v.name)}</h3></div><button type="button" class="x" data-close aria-label="閉じる">×</button></div>
    <dl class="info">
      <dt>住所</dt><dd>${esc(v.address || '—')}</dd>
      <dt>最大キャパ</dt><dd>${esc(cap)}</dd>
      <dt>アクセス</dt><dd>${esc(v.access || '—')}</dd>
      ${v.url ? `<dt>公式サイト</dt><dd><a href="${esc(v.url)}" target="_blank" rel="noopener noreferrer">${esc(v.url)}</a></dd>` : ''}
      ${v.memo ? `<dt>メモ</dt><dd>${esc(v.memo)}</dd>` : ''}
    </dl>
    <div class="sec"><h4>過去の座席表</h4>${seats}</div>
    <div class="sec"><h4>この会場の開催記録（全${sumW(all)}公演${hit.size !== all.length ? `・うち現在の条件に一致 ${sumW([...hit])}` : ''}）</h4>
      <div class="tablewrap"><table><thead><tr><th>日付</th><th>公演名</th><th>種類</th><th>出演</th></tr></thead><tbody>
      ${all.map(s => `<tr class="${hit.has(s)?'':'off'}"><td class="d">${fmtDate(s.date)}${s.start?`<br><small>${s.start.split('\n').map(t => esc(t)+'開演').join('<br>')}</small>`:''}</td><td class="ttl">${esc(s.title)}${s.note?`<br><small>${noteHtml(s.note)}</small>`:''}</td><td>${esc(s.type)}</td><td>${esc(memberLabel(s.members))}</td></tr>`).join('')}
      </tbody></table></div></div>`;
  $('#dlgBody').dataset.venue = id;
  const dlg = $('#venueDlg');
  if (!dlg.open) dlg.showModal();
  dlg.scrollTop = 0;
}

/* ---------- イベント ---------- */
svg.addEventListener('click', e => {
  const g = e.target.closest('.region'); if (!g) return;
  const tile = e.target.closest('.tile');
  if (state.region === g.dataset.region && REGION.get(state.region).special) return;
  if (state.region === g.dataset.region && tile){
    const id = tile.dataset.pref;
    const opening = !state.openPrefs.has(id);
    state.openPrefs.has(id) ? state.openPrefs.delete(id) : state.openPrefs.add(id);
    render({keepScroll:true});
    const el = document.getElementById('pref-'+id), sc = document.querySelector('.panel-scroll');
    if (el && sc) sc.scrollTop += el.getBoundingClientRect().top - sc.getBoundingClientRect().top;
    // 縦に並ぶ画面（スマホ・タブレット縦）では、開いた会場リストが見える位置まで移動
    if (opening && matchMedia('(max-width:860px)').matches) $('.panel-wrap').scrollIntoView({behavior:'smooth', block:'start'});
  } else selectRegion(g.dataset.region);
});
svg.addEventListener('keydown', e => {
  const g = e.target.closest('.region');
  if (g && (e.key === 'Enter' || e.key === ' ')){ e.preventDefault(); selectRegion(g.dataset.region); }
});
$('#legend').addEventListener('click', e => { const b = e.target.closest('[data-region]'); if (b) selectRegion(state.region === b.dataset.region ? null : b.dataset.region); });
$('#legend').addEventListener('mouseover', e => { const b = e.target.closest('[data-region]'); svg.querySelectorAll('.region').forEach(g => g.classList.toggle('hot', !!b && g.dataset.region === b.dataset.region)); });
$('#legend').addEventListener('mouseleave', () => svg.querySelectorAll('.region.hot').forEach(g => g.classList.remove('hot')));
document.addEventListener('click', e => {
  if (e.target.closest('[data-home]')) selectRegion(null);
  if (e.target.closest('[data-clear]')) clearFilters();
  const t = e.target.closest('[data-type]');
  if (t){ const v = t.dataset.type; state.f.types.has(v) ? state.f.types.delete(v) : state.f.types.add(v); render(); }
  const m = e.target.closest('[data-member]');
  if (m){ const v = m.dataset.member; state.f.members.has(v) ? state.f.members.delete(v) : state.f.members.add(v); render(); }
  const y = e.target.closest('[data-year]');
  if (y){ state.f.year = y.dataset.year; render(); }
});
$('#panel').addEventListener('click', e => {
  const v = e.target.closest('[data-venue]'); if (v){ openVenue(v.dataset.venue); return; }
  const p = e.target.closest('[data-pref]');
  if (p){ const id = p.dataset.pref; state.openPrefs.has(id) ? state.openPrefs.delete(id) : state.openPrefs.add(id); render({keepScroll:true}); }
});
$('#venueDlg').addEventListener('click', e => {
  const dlg = $('#venueDlg');
  if (e.target === dlg || e.target.closest('[data-close]')){ dlg.close(); return; }
  const z = e.target.closest('[data-zoom]');
  if (z){
    const v = state.data.venues.get($('#dlgBody').dataset.venue), s = v.seats[+z.dataset.zoom];
    const ov = document.createElement('div'); ov.className = 'zoomimg';
    ov.innerHTML = `<img src="${esc(seatSrc(s.src))}" alt="${esc(v.name)} 座席表"><p>${esc(s.label || '座席表')}・クリックで閉じる</p>`;
    ov.addEventListener('click', () => ov.remove());
    $('#dlgBody').appendChild(ov);
  }
});
function clearFilters(){ state.f = {q:'', year:'', types:new Set(), members:new Set()}; $('#q').value=''; state.limit = 60; render(); }
$('#clearBtn').addEventListener('click', clearFilters);
$('#year').addEventListener('change', e => { state.f.year = e.target.value; state.limit = 60; render(); });
let qt; $('#q').addEventListener('input', e => { clearTimeout(qt); qt = setTimeout(() => { state.f.q = e.target.value.trim(); state.limit = 60; render(); }, 150); });


/* ---------- 起動 ---------- */
// 地図を画面の残りの高さに合わせて縮め、ページ全体がスクロールしないようにする
function fitMap(){
  if (matchMedia('(max-width:860px)').matches) svg.style.removeProperty('--map-max');
  else {
  const top = svg.getBoundingClientRect().top + scrollY;
  const below = $('#legend').offsetHeight + 12 + 12 + 32; // 凡例・カード余白・ページ下余白
  svg.style.setProperty('--map-max', Math.max(360, Math.floor(innerHeight - top - below)) + 'px');
  }
  svg.classList.toggle('compact', svg.getBoundingClientRect().width < 520);
}
// 種類・メンバーの絞り込み欄：縦に余裕のある画面では最初から開く。開閉は端末ごとに記憶
const FKEY = 'urasaka-map:filters-open';
function setFiltersOpen(open, save){
  $('#fmore').hidden = !open; $('#ftoggle').setAttribute('aria-expanded', open);
  layoutTypes();
  if (save){ try { localStorage.setItem(FKEY, open ? '1' : '0'); } catch(_){} }
  fitMap();
}
$('#ftoggle').addEventListener('click', () => setFiltersOpen($('#fmore').hidden, true));
{
  let saved = null; try { saved = localStorage.getItem(FKEY); } catch(_){}
  const open = saved !== null ? saved === '1' : (innerWidth > 600 && innerHeight >= 900);
  $('#fmore').hidden = !open; $('#ftoggle').setAttribute('aria-expanded', open);
}
// 4行のどれかが幅からはみ出すときだけ、折り返し表示に切り替える
function layoutTypes(){
  const box = $('#typeChips');
  if (!box || $('#fmore').hidden) return;
  box.classList.remove('flow');
  const over = [...box.querySelectorAll('.trow')].some(r => r.scrollWidth > box.clientWidth + 1);
  box.classList.toggle('flow', over);
}
addEventListener('resize', () => requestAnimationFrame(() => { layoutTypes(); fitMap(); }));
buildMap();
// 読み込みが終わるまでは空のデータで画面を組み立て、読み込み中と表示する
setData({venues:new Map(), shows:[], warnings:[]}, 'sample', '公演データを読み込み中…');
$('#panel').innerHTML = '<div class="panel-head"><p class="eyebrow">公演記録</p><h2>読み込み中…</h2></div>';
(async () => {
  const fail = msg => {
    $('#dataStatusText').textContent = '公演データを読み込めませんでした';
    $('#panel').innerHTML = `<div class="panel-head"><p class="eyebrow">公演記録</p><h2>データを読み込めませんでした</h2><p class="empty-msg" style="padding:0">${esc(msg)}</p></div>`;
  };
  if (!window.XLSX){ fail('Excel読み込み用のファイル（js/xlsx.core.min.js）が見つからないか、壊れています。'); return; }
  try {
    const res = await fetch(DATA_FILE, {cache:'no-cache'});
    if (!res.ok) throw new Error(`${DATA_FILE} が見つかりません（${res.status}）。`);
    const d = fromWorkbook(XLSX.read(await res.arrayBuffer(), {type:'array', cellDates:false}), XLSX);
    setData(d, 'site', `公演データ ${sumW(d.shows)}公演`);
  } catch(err){ fail(err.message); }
})();
/* ---------- 日付から検索 ---------- */
function md(s){ return s.date.slice(5); }   // "MM-DD"（起動直後から使うので関数宣言にする）
function fillDateSelects(){
  const years = [...new Set(state.data.shows.map(s => s.year))].sort().reverse();
  const yv = $('#dYear').value;
  $('#dYear').innerHTML = '<option value="">指定なし</option>' + years.map(y => `<option value="${y}">${y}年</option>`).join('');
  $('#dYear').value = years.includes(yv) ? yv : '';
  if ($('#dMonth').options.length === 1)
    $('#dMonth').insertAdjacentHTML('beforeend', Array.from({length:12}, (_,i) => `<option value="${i+1}">${i+1}月</option>`).join(''));
  fillDays();
}
// 月に合わせて日の選択肢を出す（2月は29日まで）
function fillDays(){
  const m = +$('#dMonth').value, dv = $('#dDay').value;
  const max = m ? [31,29,31,30,31,30,31,31,30,31,30,31][m-1] : 31;
  $('#dDay').innerHTML = '<option value="">選択</option>' + Array.from({length:max}, (_,i) => `<option value="${i+1}">${i+1}日</option>`).join('');
  if (dv && +dv <= max) $('#dDay').value = dv;
}
function renderDateResult(){
  const y = $('#dYear').value, m = $('#dMonth').value, d = $('#dDay').value, out = $('#dateResult');
  if (!m || !d){ out.innerHTML = '<p class="dmsg">月/日を選択してください</p>'; return; }
  const key = `${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  const hits = state.data.shows.filter(s => md(s) === key && (!y || s.year === y))
    .sort((a,b) => b.date.localeCompare(a.date) || a.id.localeCompare(b.id));   // 新しい順
  const label = `${y ? y + '年' : ''}${m}月${d}日`;
  if (!hits.length){ out.innerHTML = `<p class="dmsg">該当する公演はありませんでした</p>`; return; }
  out.innerHTML = `<div class="dhead"><h3>${label}の公演</h3><span class="stats"><span><b>${sumW(hits)}</b>公演</span></span></div>
    <ol class="list dlist">${hits.map(showRow).join('')}</ol>`;
}
// 同じ月日の開催回数ランキング（全年・開催中止は除く・昼夜2公演は2回）
function renderDateRank(){
  const cnt = new Map();
  state.data.shows.forEach(s => { if (s.weight) cnt.set(md(s), (cnt.get(md(s)) || 0) + s.weight); });
  // 回数ごとにまとめる → 回数の多い順に 1位,2位,3位…（同率がいくつあっても順位は詰めて数える）
  const groups = new Map();
  [...cnt.entries()].sort((a,b) => a[0].localeCompare(b[0])).forEach(([k,n]) => { if (!groups.has(n)) groups.set(n, []); groups.get(n).push(k); });
  const ranks = [...groups.entries()].sort((a,b) => b[0] - a[0]).slice(0, 5);
  $('#dateRank').innerHTML = ranks.length ? ranks.map(([n, keys], i) => {
    const dates = keys.map(k => { const [mm,dd] = k.split('-').map(Number); return `<button type="button" class="rd" data-md="${mm}-${dd}">${mm}月${dd}日</button>`; }).join('');
    return `<li class="${i===0?'first':''}"><span class="rn">${i+1}位</span><span class="rds">${dates}</span><span class="rc">${n}回</span></li>`;
  }).join('') : '<li class="dmsg">データがありません</li>';
}
function refreshDatePage(){ fillDateSelects(); renderDateRank(); renderDateResult(); }
$('#dateForm').addEventListener('change', e => { if (e.target.id === 'dMonth') fillDays(); renderDateResult(); });
$('#dateForm').addEventListener('submit', e => e.preventDefault());
// リセット：年・月・日をすべて未選択に戻す
$('#dateForm').addEventListener('reset', () => setTimeout(() => { fillDays(); renderDateResult(); }, 0));
$('#dateRank').addEventListener('click', e => {
  const b = e.target.closest('[data-md]'); if (!b) return;
  const [mm,dd] = b.dataset.md.split('-');
  $('#dYear').value = ''; $('#dMonth').value = mm; fillDays(); $('#dDay').value = dd; renderDateResult();
  $('#dateResult').scrollIntoView({behavior:'smooth', block:'nearest'});
});
$('#dateResult').addEventListener('click', e => { const v = e.target.closest('[data-venue]'); if (v) openVenue(v.dataset.venue); });
/* ---------- 画面の切り替え（トップ／お問い合わせ先／プライバシーポリシー） ---------- */
const VIEWS = {main:'view-main', date:'view-date', contact:'view-contact', privacy:'view-privacy'};
function showView(name){
  if (!VIEWS[name]) name = 'main';
  for (const [k,id] of Object.entries(VIEWS)) document.getElementById(id).hidden = k !== name;
  document.querySelectorAll('.gnav a[data-view]').forEach(a => {
    const v = a.dataset.view === 'top' ? 'main' : a.dataset.view;
    if (v === name) a.setAttribute('aria-current','page'); else a.removeAttribute('aria-current');
  });
  if (name === 'main'){ layoutTypes(); fitMap(); }
  scrollTo(0,0);
  try { history.replaceState(null, '', name === 'main' ? location.pathname + location.search : '#' + name); } catch(_){}
}
// 「トップ」：絞り込み・拡大・開いた県をすべて最初の状態に戻して地図の画面へ
function goTop(){
  state.f = {q:'', year:defaultYear(state.data), types:new Set(), members:new Set()}; $('#q').value = '';
  selectRegion(null);
  showView('main');
}
const menuBtn = $('#menuBtn'), gnav = $('#gnav'), navBackdrop = $('#navBackdrop');
function setMenu(open){
  menuBtn.setAttribute('aria-expanded', open);
  menuBtn.setAttribute('aria-label', open ? 'メニューを閉じる' : 'メニューを開く');
  gnav.classList.toggle('open', open); navBackdrop.classList.toggle('open', open);
  if (open){ gnav.removeAttribute('inert'); setTimeout(() => gnav.querySelector('a').focus({preventScroll:true}), 60); }
  else gnav.setAttribute('inert','');
}
menuBtn.addEventListener('click', () => setMenu(menuBtn.getAttribute('aria-expanded') !== 'true'));
navBackdrop.addEventListener('click', () => setMenu(false));
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && menuBtn.getAttribute('aria-expanded') === 'true'){ setMenu(false); menuBtn.focus(); }
});
document.addEventListener('click', e => {
  const a = e.target.closest('a[data-view]'); if (!a) return;
  e.preventDefault(); setMenu(false);
  a.dataset.view === 'top' ? goTop() : showView(a.dataset.view);
});
addEventListener('hashchange', () => showView(location.hash.slice(1)));
{ const h = location.hash.slice(1); if (h === 'date' || h === 'contact' || h === 'privacy') showView(h); }
