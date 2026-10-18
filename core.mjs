export const UNITS = ['smp', 'sma'];
export const COLORS = ['#087f67', '#e5aa45', '#6489d1', '#b277bf', '#e57869', '#479daa'];
export const STATUSES = {draft:'Persiapan', open:'Pemilihan dibuka', closed:'Pemilihan ditutup'};
export function escapeHTML(value='') {
  return String(value).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
export function normalizeRFID(value) {
  // Tetap string: 00001234 berbeda dari 1234.
  return String(value ?? '').trim().toUpperCase();
}
export function normalizeCode(value) { return String(value ?? '').trim().toUpperCase(); }
export function voterKey(code) {
  const clean = normalizeCode(code);
  if (!clean || clean.length > 80 || /[\u0000-\u001f]/.test(clean)) throw new Error('Kode pemilih wajib diisi, maksimal 80 karakter.');
  return Array.from(new TextEncoder().encode(clean), b=>b.toString(16).padStart(2,'0')).join('');
}
export function initials(name) { return String(name).trim().split(/\s+/).slice(0,2).map(s=>s[0]||'').join('').toUpperCase(); }
export function percentage(n, total) { return total ? Math.round(n/total*1000)/10 : 0; }
export function photoSrc(value) {
  return typeof value === 'string' && /^data:image\/jpeg;base64,[A-Za-z0-9+/=]+$/.test(value) && value.length <= 110000 ? value : '';
}
export function tally(candidates, ballots, voters) {
  const counts = Object.fromEntries(candidates.map(c=>[c.id,0]));
  for (const b of ballots) { if (Object.hasOwn(counts,b.candidateId)) counts[b.candidateId]++; }
  const valid = Object.values(counts).reduce((a,b)=>a+b,0);
  const voterIds = new Set(voters.map(v=>v.id));
  const voted = new Set(ballots.filter(b=>voterIds.has(b.id)).map(b=>b.id));
  return {counts, total:ballots.length, valid, unrecognized:ballots.length-valid,
    registered:voters.length, voted:voted.size, remaining:Math.max(0,voters.length-voted.size),
    participation:percentage(voted.size,voters.length)};
}
export function validateCandidates(candidates) {
  if (candidates.length < 2) throw new Error('Daftarkan minimal 2 calon terlebih dahulu.');
  const ids = new Set(), codes = new Set();
  for (const c of candidates) {
    const rfid = normalizeRFID(c.rfid);
    if (!c.name?.trim() || !c.code?.trim()) throw new Error('Nama dan kode calon belum lengkap.');
    if (!rfid || rfid.length > 80) throw new Error(`Kartu RFID calon ${c.code} belum terdaftar atau terlalu panjang.`);
    if (ids.has(rfid)) throw new Error('ID RFID setiap calon dalam satu pemilihan harus berbeda.');
    if (codes.has(c.code)) throw new Error('Kode calon harus berbeda.');
    ids.add(rfid); codes.add(c.code);
  }
  return true;
}
export function checkVote({election, candidates, voters, ballots}, studentId, candidateId) {
  if (election?.status !== 'open') throw new Error('Pemilihan belum dibuka atau sudah ditutup.');
  if (!voters.some(v=>v.id===studentId)) throw new Error('Pemilih tidak terdaftar pada periode ini.');
  if (!candidates.some(c=>c.id===candidateId)) throw new Error('Calon tidak ditemukan.');
  if (ballots.some(b=>b.id===studentId)) throw new Error('Siswa ini sudah menggunakan hak pilih.');
  return true;
}
export function candidateForRFID(candidates, raw) {
  const id = normalizeRFID(raw);
  const found = candidates.filter(c=>normalizeRFID(c.rfid)===id && id);
  if (found.length>1) throw new Error('ID kartu ganda. Minta panitia memeriksa pengaturan calon.');
  if (!found.length) throw new Error('Kartu tidak dikenal untuk pemilihan ini. Tempelkan kartu calon yang terdaftar.');
  return found[0];
}
export function parseCSV(text) {
  text = String(text).replace(/^\uFEFF/,'').replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  const first = text.split('\n')[0] || '';
  const delimiter = (first.match(/;/g)||[]).length>(first.match(/,/g)||[]).length ? ';' : ',';
  const rows = []; let row=[], field='', quoted=false;
  for(let i=0;i<text.length;i++) {
    const c=text[i];
    if(c==='"') {
      if(quoted && text[i+1]==='"'){field+='"';i++;}
      else if(quoted || field===''){quoted=!quoted;}
      else throw new Error('Format tanda petik CSV tidak valid.');
    } else if(c===delimiter && !quoted){row.push(field);field='';}
    else if(c==='\n' && !quoted){row.push(field);if(row.some(v=>v.trim()))rows.push(row);row=[];field='';}
    else field+=c;
  }
  if(quoted) throw new Error('CSV memiliki tanda petik yang belum ditutup.');
  row.push(field);if(row.some(v=>v.trim()))rows.push(row);
  return rows;
}
export function parseVotersCSV(text) {
  const rows=parseCSV(text);
  if(rows.length<2)throw new Error('CSV harus berisi judul kolom dan minimal satu siswa.');
  const headers=rows.shift().map(x=>x.trim().toLowerCase());
  const indexes=['kode','nama','kelas'].map(h=>headers.indexOf(h));
  if(indexes.some(i=>i<0))throw new Error('Kolom CSV wajib: kode,nama,kelas.');
  const seen=new Set();
  return rows.map((r,i)=>{
    const code=normalizeCode(r[indexes[0]]), name=String(r[indexes[1]]||'').trim(), classroom=String(r[indexes[2]]||'').trim();
    const id=voterKey(code);
    if(!name || name.length>120 || !classroom || classroom.length>40)throw new Error(`Periksa nama/kelas pada baris ${i+2}.`);
    if(seen.has(id))throw new Error(`Kode ${code} muncul dua kali dalam file. Impor dibatalkan.`);
    seen.add(id);return {id,code,name,classroom};
  });
}
export function csvCell(value) {
  let text=String(value ?? '');
  // Mencegah nilai siswa diperlakukan sebagai formula saat CSV dibuka di Excel.
  if(/^[=+\-@\t\r]/.test(text))text="'"+text;
  return '"'+text.replace(/"/g,'""')+'"';
}
export function makeCSV(rows) {return '\uFEFF'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n');}
export function uniqueId(prefix='') { return prefix + (globalThis.crypto?.randomUUID?.() || Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,10)); }

export function demoSeed() {
  const data={schools:{},elections:{}};
  for(const [idx,unit] of UNITS.entries()) {
    const id=`demo-2026-${unit}`;
    data.schools[unit]={unit,currentElectionId:id,history:[id]};
    const names=idx ? ['Naufal Ramadhan','Farhan Maulana'] : ['Ahmad Fauzan','Muhammad Rizki'];
    const candidates=names.map((name,n)=>({id:'calon-'+String.fromCharCode(65+n),code:String.fromCharCode(65+n),name,deputy:n?'Fikri Ramadhan':'Ilham Pratama',classroom:idx?'XI IPA':'VIII A',vision:n?'OSIS yang aktif, kreatif, dan dekat dengan siswa.':'Bersama membangun sekolah yang berprestasi dan berakhlak.',mission:'Menghidupkan kegiatan siswa\nMenjaga kebersihan lingkungan\nMenampung aspirasi seluruh siswa',rfid:n?'0000000002':'0000000001',photo:'',color:COLORS[n],order:n}));
    const total=idx?240:360, suara=idx?146:206, a=idx?67:118;
    const voters=Array.from({length:total},(_,i)=>{const code=(idx?'SMA':'SMP')+String(i+1).padStart(4,'0');return {id:voterKey(code),code,name:'Siswa Contoh '+String(i+1).padStart(3,'0'),queueNo:i+1,classroom:(idx?['X A','X B','XI A','XI B','XII A','XII B']:['VII A','VII B','VIII A','VIII B','IX A','IX B'])[i%6]};});
    const ballots=voters.slice(0,suara).map((v,i)=>({id:v.id,candidateId:candidates[i<a?0:1].id,castAt:new Date(Date.now()-(suara-i)*90000).toISOString(),operatorUid:'DEMO'}));
    data.elections[id]={election:{id,unit,title:`Pemilihan Ketua OSIS ${unit.toUpperCase()}`,period:'2026/2027 · Contoh',status:'open',createdAt:new Date().toISOString(),openedAt:new Date().toISOString(),closedAt:null},candidates,voters,ballots,calls:[],stations:{}};
  }
  return data;
}
