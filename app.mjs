import {UNITS,COLORS,STATUSES,escapeHTML as e,normalizeRFID,normalizeCode,voterKey,initials,percentage,photoSrc,tally,validateCandidates,candidateForRFID,parseVotersCSV,makeCSV,uniqueId} from './core.mjs';
import {DemoStore,createFirebaseStore,readableError} from './data.mjs';
import {studentFromNames,parseStudentRows,parseStudentCSV,newStudentsOnly,callStatus,queueRows,ownCall,CALL_LABELS,callText,queueStats} from './students.mjs';
import {readExcel} from './excel.mjs';

const config=window.APP_CONFIG||{};
const configured=['apiKey','authDomain','projectId','appId'].every(k=>typeof config.firebase?.[k]==='string' && config.firebase[k].trim());
const schoolName=config.schoolName||'Al-Miftah';
const paths={
 mic:'M9 4a3 3 0 0 1 6 0v8a3 3 0 0 1-6 0z M5 10v2a7 7 0 0 0 14 0v-2 M12 19v3 M8 22h8',
 grid:'M3 3h7v7H3z M14 3h7v7h-7z M3 14h7v7H3z M14 14h7v7h-7z',
 card:'M3 5h18v14H3z M3 10h18 M7 15h4',
 users:'M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2 M16 3a4 4 0 0 1 0 8 M22 21v-2a4 4 0 0 0-3-3.87 M13 7a4 4 0 1 1-8 0 4 4 0 0 1 8 0',
 list:'M9 5h11 M9 12h11 M9 19h11 M3 5h1 M3 12h1 M3 19h1',
 chart:'M4 3v17h17 M8 15v-4 M13 15V6 M18 15V9',
 settings:'M4 7h16 M4 17h16 M8 4v6 M16 14v6',
 arrow:'M4 12h16 M14 6l6 6-6 6',
 check:'M5 12l4 4L19 6',
 down:'M6 9l6 6 6-6',
 calendar:'M5 5h14v16H5z M8 3v4 M16 3v4 M5 10h14',
 help:'M9.1 9a3 3 0 0 1 5.8 1c0 2-3 3-3 3 M12 17h.01 M22 12a10 10 0 1 1-20 0 10 10 0 0 1 20 0',
 shield:'M12 3l8 3v6c0 5-8 9-8 9s-8-4-8-9V6z M8 12l3 3 5-6',
 rfid:'M4 8h7v11H4z M8 4a10 10 0 0 1 12 12 M10 7a6 6 0 0 1 7 7 M11 11a2 2 0 0 1 2 2',
 plus:'M12 5v14 M5 12h14',
 search:'M21 21l-5-5 M18 10a7 7 0 1 1-14 0 7 7 0 0 1 14 0',
 download:'M12 3v12 M7 10l5 5 5-5 M4 17v4h16v-4',
 upload:'M12 16V4 M7 9l5-5 5 5 M4 17v4h16v-4',
 trash:'M4 7h16 M9 7V4h6v3 M6 7l1 14h10l1-14 M10 11v6 M14 11v6',
 edit:'M4 20l5-1L20 8l-4-4L5 15z M13 7l4 4',
 close:'M6 6l12 12 M6 18 18 6',
 clock:'M12 7v5l3 2 M22 12a10 10 0 1 1-20 0 10 10 0 0 1 20 0',
 lock:'M6 10h12v11H6z M8 10V7a4 4 0 0 1 8 0v3',
 expand:'M8 3H3v5 M16 3h5v5 M3 16v5h5 M21 16v5h-5',
 sound:'M3 9h4l5-4v14l-5-4H3z M16 8a6 6 0 0 1 0 8 M19 5a10 10 0 0 1 0 14',
 logout:'M9 4H4v16h5 M12 12h9 M17 8l4 4-4 4',
 trophy:'M8 3h8v7a4 4 0 0 1-8 0z M8 5H4v3a4 4 0 0 0 4 4 M16 5h4v3a4 4 0 0 1-4 4 M12 14v6 M8 21h8',
 image:'M3 3h18v18H3z M3 16l6-6 4 4 3-3 5 5 M17 7h.01',
 print:'M6 8V3h12v5 M6 17H3V8h18v9h-3 M6 14h12v7H6z',
 school:'M3 10l9-7 9 7 M5 9v12h14V9 M10 21v-7h4v7 M8 10h.01 M16 10h.01'
};
const icon=(key,cls='')=>`<svg class="icon ${cls}" viewBox="0 0 24 24" aria-hidden="true"><path d="${paths[key]||paths.grid}"></path></svg>`;
const labels={dashboard:'Dashboard',calling:'Pemanggilan',voting:'Bilik suara',candidates:'Data calon',voters:'Daftar pemilih',results:'Hasil suara',settings:'Pengaturan'};
const navIcons={dashboard:'grid',calling:'mic',voting:'card',candidates:'users',voters:'list',results:'chart',settings:'settings'};
const emptyPacket=()=>({school:null,election:null,candidates:[],voters:[],ballots:[],calls:[],ready:false});
const safeRead=(k,fallback='')=>{try{return localStorage.getItem(k)||fallback;}catch{return fallback;}};
const safeSave=(k,v)=>{try{localStorage.setItem(k,v);}catch{/* Preferensi tampilan bersifat opsional. */}};
const stationId=(()=>{try{let id=sessionStorage.getItem('osis-bilik-v2');if(!id){id=uniqueId('bilik-');sessionStorage.setItem('osis-bilik-v2',id);}return id;}catch{return uniqueId('bilik-');}})();
function useAccountStation(uid){try{const key='osis-bilik-v2-'+uid;let id=sessionStorage.getItem(key);if(!id){id=uniqueId('bilik-');sessionStorage.setItem(key,id);}state.stationId=id;}catch{state.stationId=uniqueId('bilik-');}}
const state={view:'dashboard',unit:UNITS.includes(safeRead('osis-unit'))?safeRead('osis-unit'):'smp',selectedId:null,
 packet:emptyPacket(),backend:null,logged:false,loginError:'',booting:false,online:navigator.onLine,
 armed:null,sending:false,success:false,scanError:'',search:'',classFilter:'all',statusFilter:'all',page:1,
 voterSearch:'',kiosk:false,sound:safeRead('osis-sound','1')==='1',confirm:null,modalPhoto:'',photoLoading:false,
 importRows:[],importDuplicates:0,photoToken:null,csvToken:null,unsub:null,stationId,queueBusy:false,callClass:'all',callFilter:'all',callSearch:'',callPage:1,callSound:safeRead('osis-call-sound','1')==='1',voiceURI:safeRead('osis-voice','')};
const number=n=>new Intl.NumberFormat('id-ID').format(n);
const pct=n=>new Intl.NumberFormat('id-ID',{maximumFractionDigits:1}).format(n)+'%';
const dateValue=v=>v?.toDate?.()||new Date(v||0);
const formatTime=v=>{const d=dateValue(v);return Number.isNaN(d.getTime())?'—':d.toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit',timeZone:'Asia/Jakarta'});};
const isDemo=()=>state.backend?.mode==='demo';
const isAdmin=()=>state.backend?.role?.role==='admin';
const current=()=>state.packet.election?.id===state.packet.school?.currentElectionId;
const editable=()=>isAdmin() && current() && state.packet.election?.status==='draft' && state.packet.ready;
const totals=()=>tally(state.packet.candidates,state.packet.ballots,state.packet.voters);
const electionId=()=>{if(!state.packet.election)throw new Error('Buat periode pemilihan terlebih dahulu.');return state.packet.election.id;};
const canVote=()=>!!state.packet.election && state.packet.election.status==='open' && current() && state.packet.ready && (isDemo()||state.online);
const thisCall=()=>ownCall(state.packet,state.stationId,state.backend?.uid);
const statusBadge=()=>{const status=state.packet.election?.status;return `<span class="status dot ${status==='open'?'':status==='draft'?'amber':'gray'}">${e(STATUSES[status]||'Belum ada periode')}</span>`;};

function toast(message,error=false){
 const item=document.createElement('div');item.className='toast'+(error?' error':'');item.innerHTML=icon(error?'help':'check')+`<span>${e(message)}</span>`;
 document.getElementById('toast-root').append(item);setTimeout(()=>item.remove(),error?8000:4500);
}
function report(error){const message=readableError(error);const box=document.getElementById('form-error');if(box)box.textContent=message;toast(message,true);}
function focusScan(){if(state.view==='voting'&&state.armed&&!state.sending&&!document.querySelector('dialog[open]'))document.getElementById('rfid-input')?.focus({preventScroll:true});}
function requireNoSession(){if(state.armed || state.sending || state.queueBusy || (current()&&state.packet.election?.status==='open'&&thisCall()))throw new Error('Selesaikan sesi atau tekan Lanjut/lewati untuk menunda panggilan terlebih dahulu.');}
function navigate(view){if(!['calling','voting'].includes(view))requireNoSession();if(state.sending||state.queueBusy)throw new Error('Tunggu proses penyimpanan selesai.');state.view=view;state.success=false;state.scanError='';render();window.scrollTo({top:0,behavior:'instant'});}
function empty(title,text,button='') {return `<div class="panel empty">${icon('school')}<h2>${e(title)}</h2><p>${e(text)}</p>${button}</div>`;}
function avatar(c,portraitNote=true){const src=photoSrc(c.photo);return `<div class="portrait">${src?`<img src="${src}" alt="Foto ${e(c.name)}">`:`<span class="initials" aria-hidden="true">${e(initials(c.name))}</span>${portraitNote?'<span class="portrait-note">Foto calon belum diunggah</span>':''}`}</div>`;}
function barChart(candidates,counts){
 const w=640,h=247,left=38,right=15,top=20,bottom=191,width=w-left-right;
 const max=Math.max(10,...candidates.map(c=>counts[c.id]||0)), ceiling=Math.ceil(max/4)*4;
 const slot=width/Math.max(1,candidates.length),bw=Math.min(88,slot*.48);
 let grid='';for(let i=0;i<=4;i++){const y=bottom-(bottom-top)*i/4;grid+=`<path d="M${left} ${y}H${w-right}" stroke="#edf2ee" ${i?'stroke-dasharray="3 5"':''}/><text x="${left-10}" y="${y+3}" text-anchor="end" font-size="10" fill="#9aaba0">${number(ceiling*i/4)}</text>`;}
 let bars=candidates.map((c,i)=>{const v=counts[c.id]||0,bh=(bottom-top)*v/ceiling,x=left+slot*i+slot/2;return `<rect x="${x-bw/2}" y="${bottom-bh}" width="${bw}" height="${Math.max(0,bh)}" rx="6" fill="${COLORS.includes(c.color)?c.color:COLORS[i%COLORS.length]}"/><text x="${x}" y="${bottom-bh-10}" text-anchor="middle" font-size="12" font-weight="600" fill="#52765e">${number(v)}</text><text x="${x}" y="214" text-anchor="middle" font-size="11" fill="#687f6e">Calon ${e(c.code)}</text><text x="${x}" y="232" text-anchor="middle" font-size="9" fill="#9aaba0">${e(c.name.length>19?c.name.slice(0,18)+'…':c.name)}</text>`;}).join('');
 return `<svg class="chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="Grafik suara ${candidates.map(c=>e('calon '+c.code+': '+number(counts[c.id]||0))).join(', ')}">${grid}${bars}${candidates.length?'':`<text x="320" y="115" text-anchor="middle" fill="#9aaba0" font-size="13">Belum ada calon</text>`}</svg>`;
}
function stats(){const t=totals();return `<div class="stat-grid">${[
 ['Pemilih terdaftar',number(t.registered),'users','Daftar pemilih tetap'],['Suara masuk',number(t.total),'card','<span class="green">'+pct(t.participation)+'</span> dari total pemilih'],
 ['Belum memilih',number(t.remaining),'clock','Siswa dengan hak pilih tersedia'],['Partisipasi',pct(t.participation),'chart','Satu siswa, satu suara']
 ].map(([label,value,ic,sub])=>`<div class="stat-card"><div class="stat-icon">${icon(ic)}</div><div class="stat-label">${label}</div><div class="stat-number">${value}</div><div class="stat-sub">${sub}</div></div>`).join('')}</div>`;}
function ballotArt(){return `<svg viewBox="0 0 340 250" aria-hidden="true"><defs><linearGradient id="boxgrad" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#f5f0de"/><stop offset="1" stop-color="#c5dbcb"/></linearGradient></defs><ellipse cx="184" cy="220" rx="104" ry="13" fill="#034b3c" opacity=".5"/><path d="m91 130 111-17 66 37-111 19z" fill="#faf7eb"/><path d="m91 130 66 39v69l-66-39z" fill="#cfdecd"/><path d="m157 169 111-19v69l-111 19z" fill="url(#boxgrad)"/><path d="m119 135 83-12 25 13-83 13z" fill="#0d6450"/><rect x="151" y="44" width="68" height="90" rx="5" transform="rotate(12 185 89)" fill="#fcfaf1"/><rect x="164" y="60" width="39" height="5" rx="2" transform="rotate(12 185 89)" fill="#c8dacf"/><circle cx="183" cy="95" r="15" fill="#e8f2e7"/><path d="m175 94 6 7 13-13" fill="none" stroke="#48815d" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/><path d="M244 68v18 M235 77h18 M69 70v12 M63 76h12" stroke="#edcc84" stroke-width="3" stroke-linecap="round"/><circle cx="276" cy="109" r="3" fill="#e4c480"/><circle cx="87" cy="177" r="2" fill="#82b19c"/><path d="m187 184 13-2v18l-13 2z" fill="#82aa91"/><path d="m237 26 8 3-5 9-8-3z" fill="#c9a54d" opacity=".8"/><path d="m111 38 5 6-5 5-6-6z" fill="#7ead96"/></svg>`;}
function renderDashboard(){
 const p=state.packet,t=totals(), recent=[...p.ballots].sort((a,b)=>dateValue(b.castAt)-dateValue(a.castAt)).slice(0,5);
 return `${stats()}<section class="hero"><div class="hero-content"><div class="eyebrow">${icon('shield')} PESTA DEMOKRASI PELAJAR</div><h2>Suaramu hari ini.<br><em>Pemimpin kita esok hari.</em></h2><p>Wujudkan OSIS yang aktif dan berprestasi.<br>Pilih dengan hati, untuk ${e(state.unit.toUpperCase())} ${e(schoolName)}.</p><div class="btn-row"><button class="btn gold" data-view="calling">Mulai pemanggilan ${icon('arrow')}</button><button class="btn light" data-view="candidates">Lihat calon</button></div></div><div class="hero-art">${ballotArt()}</div></section>
 <div class="content-grid"><section class="panel"><div class="panel-head"><div><h2>Perolehan suara</h2><p>${p.election?.status==='closed'?'Hasil akhir':'Hasil sementara'} · ${e(state.unit.toUpperCase())} ${e(schoolName)}</p></div><button class="btn ghost tiny" data-view="results">Detail ${icon('arrow')}</button></div>${barChart(p.candidates,t.counts)}<div class="legend">${p.candidates.map((c,i)=>`<span><i style="--c:${COLORS[i%COLORS.length]}"></i>Calon ${e(c.code)}</span>`).join('')}</div></section>
 <section class="panel"><div class="panel-head"><div><h2>Aktivitas pemilihan</h2><p>Suara yang sudah dikonfirmasi</p></div>${icon('clock')}</div><div class="timeline">${recent.length?recent.map((b,i)=>`<div class="timeline-item"><span class="timeline-icon">${icon('check')}</span><div><b>Suara ke-${number(t.total-i)} tersimpan</b><small>${formatTime(b.castAt)} WIB · ${state.unit.toUpperCase()}</small></div></div>`).join(''):'<p class="muted tiny">Belum ada suara masuk.</p>'}</div><div class="panel-note">Pilihan siswa tidak ditampilkan di aktivitas.</div></section></div>`;
}
function renderCandidates(){
 const p=state.packet;
 return `<div class="toolbar"><span class="muted tiny">${number(p.candidates.length)} calon · ${e(state.unit.toUpperCase())} ${e(schoolName)}</span><button class="btn" data-action="add-candidate" ${editable()?'':'disabled'}>${icon('plus')} Tambah calon</button></div>${!editable()?'<div class="hint">Nama, foto, dan kartu calon dapat diatur oleh admin saat tahap <b>Persiapan</b>. Periode yang telah dibuka menyimpan data calon secara tetap.</div>':''}
 ${p.candidates.length?`<div class="candidate-grid">${p.candidates.map(c=>`<article class="candidate-card" style="--c:${c.color}"><span class="candidate-number">CALON ${e(c.code)}</span>${avatar(c)}<div class="candidate-body"><span class="tag">${e(c.classroom||state.unit.toUpperCase())}</span><h3 class="mt-20">${e(c.name)}</h3>${c.deputy?`<p class="deputy">Wakil: ${e(c.deputy)}</p>`:''}<p class="vision">“${e(c.vision||'Visi calon belum diisi.')}”</p><div class="candidate-footer"><span class="rfid-badge">${icon('rfid')} ${c.rfid?'Kartu RFID terdaftar':'Belum ada kartu'}</span><div class="btn-row"><button class="btn secondary small" data-action="candidate-detail" data-id="${e(c.id)}">Profil</button>${editable()?`<button class="icon-btn" data-action="edit-candidate" data-id="${e(c.id)}" title="Ubah calon ${e(c.code)}">${icon('edit')}</button><button class="icon-btn" data-action="delete-candidate" data-id="${e(c.id)}" title="Hapus calon ${e(c.code)}">${icon('trash')}</button>`:''}</div></div></div></article>`).join('')}</div>`:empty('Siapa pemimpin berikutnya?','Tambahkan minimal dua calon. Unggah foto dan daftarkan kartu RFID A dan B.',`<button class="btn" data-action="add-candidate" ${editable()?'':'disabled'}>${icon('plus')} Tambah calon pertama</button>`)}`;
}
function callBadge(status){return `<span class="status ${status==='skipped'?'amber':status==='waiting'?'gray':''}">${e(CALL_LABELS[status]||status)}</span>`;}
function voiceOptions(){const voices=window.speechSynthesis?.getVoices?.()||[];return '<option value="">Otomatis (Indonesia)</option>'+voices.map(v=>`<option value="${e(v.voiceURI)}" ${state.voiceURI===v.voiceURI?'selected':''}>${e(v.name)} · ${e(v.lang)}</option>`).join('');}
function renderCalling(){
 const p=state.packet,call=thisCall(),voter=p.voters.find(v=>v.id===call?.id),q=queueStats(p),classes=[...new Set(p.voters.map(v=>v.classroom))].sort((a,b)=>a.localeCompare(b,'id',{numeric:true}));
 const waiting=queueRows(p,{classroom:state.callClass,status:'waiting'}),skipped=queueRows(p,{classroom:state.callClass,status:'skipped'});
 const all=[...p.voters].sort((a,b)=>(a.queueNo??Number.MAX_SAFE_INTEGER)-(b.queueNo??Number.MAX_SAFE_INTEGER)||a.name.localeCompare(b.name,'id',{numeric:true}));
 const rows=all.filter(v=>(state.callClass==='all'||v.classroom===state.callClass)&&(state.callFilter==='all'||callStatus(p,v.id)===state.callFilter)&&`${v.name} ${v.classroom}`.toLowerCase().includes(state.callSearch.toLowerCase()));
 const pages=Math.max(1,Math.ceil(rows.length/30));state.callPage=Math.min(state.callPage,pages);const shown=rows.slice((state.callPage-1)*30,state.callPage*30);
 return `<div class="call-stat-grid">${[['Menunggu',q.waiting,'clock'],['Dipanggil / memilih',q.called+q.active,'mic'],['Dilewati',q.skipped,'arrow'],['Sudah memilih',q.voted,'check']].map(([label,n,ic])=>`<div class="call-stat">${icon(ic)}<div><b>${number(n)}</b><small>${label}</small></div></div>`).join('')}</div>
 <div class="calling-layout"><section class="panel call-main"><div class="panel-head"><div><h2>${icon('mic')} Panggilan siswa</h2><p>${e(state.unit.toUpperCase())} ${e(schoolName)} · Satu sesi per siswa</p></div>${voter?callBadge(call.status):'<span class="status gray">Siap memanggil</span>'}</div>
 <div class="called-student" aria-live="polite">${voter?`<span class="eyebrow">${call.status==='active'?'SISWA SEDANG MEMILIH':'SILAKAN MENUJU BILIK SUARA'}</span><h2>${e(voter.name)}</h2><span class="called-class">Kelas ${e(voter.classroom)}</span><p>${call.status==='active'?'RFID hanya aktif untuk siswa ini.':'Panitia memeriksa kehadiran sebelum membuka sesi RFID.'}</p>`:`<div class="call-symbol">${icon('mic')}</div><h2>${canVote()?'Siapa pemilih berikutnya?':'Pemanggilan belum dibuka'}</h2><p>${canVote()?'Tekan Panggil berikutnya untuk menyebut nama dan kelas.':'Buka pemilihan dan pastikan data terhubung terlebih dahulu.'}</p>`}</div>
 <div class="call-primary-actions"><button class="btn" data-action="call-next" ${canVote()&&!state.queueBusy&&!state.sending?'':'disabled'}>${icon('arrow')} ${voter?'Lanjut':'Panggil berikutnya'}</button><button class="btn secondary" data-action="repeat-call" ${voter&&!state.queueBusy?'':'disabled'}>${icon('sound')} Panggil ulang</button>${voter?`<button class="btn soft" data-action="activate-called" ${canVote()&&!state.queueBusy&&!state.sending?'':'disabled'}>${icon('rfid')} ${call.status==='active'?'Buka bilik suara':'Siswa hadir — buka sesi'}</button>`:''}</div>
 <p class="call-footnote">${voter?'Jika siswa belum hadir, tekan Lanjut. Ia masuk daftar Dilewati dan tetap memiliki hak pilih.':number(waiting.length)+' siswa menunggu'+(state.callClass==='all'?'':' di kelas '+e(state.callClass))+'. Urutan mengikuti data yang ditambahkan.'}</p>
 ${state.queueBusy?'<div class="loading-line" aria-label="Memperbarui antrean"></div>':''}
 ${voter?'<button class="text-btn" data-action="skip-only">Tunda panggilan ini tanpa memanggil siswa lain</button>':''}</section>
 <aside class="panel"><div class="panel-head"><div><h2>Siswa yang dilewati</h2><p>Masih dapat dipanggil dan memilih</p></div><span class="status amber">${number(skipped.length)}</span></div><p class="tiny muted">Panggil kembali siswa yang sudah hadir. Hak pilih tidak berkurang saat dilewati.</p><button class="btn secondary full-width" data-action="call-skipped" ${skipped.length&&canVote()&&!state.queueBusy&&!state.sending?'':'disabled'}>${icon('mic')} Panggil yang dilewati</button><div class="divider"></div><label class="checkline"><input id="call-sound-toggle" type="checkbox" ${state.callSound?'checked':''}> Suara pemanggilan</label><div class="call-voice"><label for="call-voice">Suara pemanggil</label><select id="call-voice">${voiceOptions()}</select></div><button class="btn ghost tiny" data-action="test-call-sound">${icon('sound')} Uji suara nama & kelas</button><div class="hint mb-0 mt-20">Kartu A/B baru aktif setelah panitia menekan <b>Siswa hadir — buka sesi</b>. Setelah satu suara tersimpan, siswa dikeluarkan dari antrean.</div></aside></div>
 <section class="panel mt-20"><div class="toolbar"><div class="search">${icon('search')}<input id="call-search" value="${e(state.callSearch)}" placeholder="Cari nama atau kelas…" aria-label="Cari antrean siswa"></div><div class="filters"><select id="call-class" aria-label="Filter kelas panggilan"><option value="all">Semua kelas</option>${classes.map(c=>`<option value="${e(c)}" ${state.callClass===c?'selected':''}>${e(c)}</option>`).join('')}</select><select id="call-filter" aria-label="Filter status panggilan"><option value="all">Semua status</option>${Object.entries(CALL_LABELS).map(([k,label])=>`<option value="${k}" ${state.callFilter===k?'selected':''}>${label}</option>`).join('')}</select></div></div><p class="tiny muted">Filter kelas juga berlaku untuk tombol Lanjut. Pencarian nama hanya menyaring tabel.</p>
 <div class="table-wrap"><table><thead><tr><th>No.</th><th>Nama siswa</th><th>Kelas</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${shown.map((v,i)=>{const status=callStatus(p,v.id),record=(p.calls||[]).find(c=>c.id===v.id),mine=record?.stationId===state.stationId&&record?.operatorUid===state.backend.uid;return `<tr><td>${(state.callPage-1)*30+i+1}</td><td class="name-cell">${e(v.name)}</td><td>${e(v.classroom)}</td><td>${callBadge(status)}</td><td>${status==='voted'?'<span class="tiny muted">Hak pilih terpakai</span>':['called','active'].includes(status)?mine?`<button class="btn secondary small" data-action="repeat-call">Panggil ulang</button>`:`<span class="tiny muted">Bilik lain</span>${isAdmin()?` <button class="btn secondary small" data-action="release-call" data-id="${e(v.id)}" ${state.queueBusy?'disabled':''}>Lepas panggilan</button>`:''}`:`<button class="btn secondary small" data-action="call-specific" data-id="${e(v.id)}" ${canVote()&&!state.queueBusy&&!state.sending?'':'disabled'}>${icon('mic')} Panggil</button>`}</td></tr>`;}).join('')||'<tr><td colspan="5" style="text-align:center;padding:28px">Tidak ada siswa pada filter ini.</td></tr>'}</tbody></table></div>
 <div class="table-footer"><span>${number(rows.length)} siswa · Halaman ${state.callPage}/${pages}</span><div class="btn-row"><button class="btn secondary small" data-action="call-previous-page" ${state.callPage===1?'disabled':''}>Sebelumnya</button><button class="btn secondary small" data-action="call-next-page" ${state.callPage>=pages?'disabled':''}>Berikutnya</button></div></div></section>`;
}
function filteredVoters(){
 const voted=new Set(state.packet.ballots.map(b=>b.id)),q=state.search.toLowerCase();
 return state.packet.voters.filter(v=>(state.classFilter==='all'||v.classroom===state.classFilter)&&(state.statusFilter==='all'||((state.statusFilter==='voted')===voted.has(v.id)))&&(!q||`${v.name} ${v.code} ${v.classroom}`.toLowerCase().includes(q))).sort((a,b)=>a.name.localeCompare(b.name,'id',{numeric:true}));
}
function renderVoters(){
 const p=state.packet,rows=filteredVoters(),voted=new Set(p.ballots.map(b=>b.id));const pages=Math.max(1,Math.ceil(rows.length/35));state.page=Math.min(state.page,pages);const shown=rows.slice((state.page-1)*35,state.page*35);
 const classes=[...new Set(p.voters.map(v=>v.classroom))].sort();
 return `<div class="toolbar"><div class="search">${icon('search')}<input id="voter-search" placeholder="Cari nama, kode, atau kelas…" aria-label="Cari pemilih" value="${e(state.search)}"></div><div class="btn-row"><button class="btn secondary" data-action="download-template">${icon('download')} Template Excel</button><button class="btn secondary" data-action="import-voters" ${editable()?'':'disabled'}>${icon('upload')} Impor Excel / CSV</button><button class="btn" data-action="add-voter" ${editable()?'':'disabled'}>${icon('plus')} Tambah siswa</button></div></div>
 <section class="panel"><div class="toolbar"><div><h2 class="mb-0" style="font-size:16px">Daftar pemilih tetap</h2><span class="muted tiny">${number(p.voters.length)} siswa terdaftar untuk periode ini</span></div><div class="filters"><select id="class-filter" aria-label="Filter kelas"><option value="all">Semua kelas</option>${classes.map(c=>`<option ${state.classFilter===c?'selected':''} value="${e(c)}">${e(c)}</option>`).join('')}</select><select id="status-filter" aria-label="Filter status"><option value="all" ${state.statusFilter==='all'?'selected':''}>Semua status</option><option value="pending" ${state.statusFilter==='pending'?'selected':''}>Belum memilih</option><option value="voted" ${state.statusFilter==='voted'?'selected':''}>Sudah memilih</option></select></div></div>
 <div class="table-wrap"><table><thead><tr><th>No.</th><th>Kode / NIS</th><th>Nama siswa</th><th>Kelas</th><th>Hak pilih</th>${editable()?'<th>Aksi</th>':''}</tr></thead><tbody>${shown.length?shown.map((v,i)=>`<tr><td>${(state.page-1)*35+i+1}</td><td class="mono">${e(v.code.startsWith('AUTO-')?'Otomatis':v.code)}</td><td class="name-cell">${e(v.name)}</td><td>${e(v.classroom)}</td><td><span class="status ${voted.has(v.id)?'':'gray'}">${voted.has(v.id)?'Sudah memilih':'Belum memilih'}</span></td>${editable()?`<td><button class="icon-btn" data-action="delete-voter" data-id="${e(v.id)}" title="Hapus ${e(v.name)}">${icon('trash')}</button></td>`:''}</tr>`).join(''):`<tr><td colspan="6" style="text-align:center;padding:35px">${p.voters.length?'Tidak ada siswa yang cocok.':'Belum ada pemilih. Tambahkan siswa atau impor CSV.'}</td></tr>`}</tbody></table></div>
 <div class="table-footer"><span>${number(rows.length)} hasil · Halaman ${state.page}/${pages}</span><div class="btn-row"><button class="btn secondary small" data-action="previous-page" ${state.page===1?'disabled':''}>Sebelumnya</button><button class="btn secondary small" data-action="next-page" ${state.page>=pages?'disabled':''}>Berikutnya</button></div></div></section>`;
}
function renderVoting(){
 const p=state.packet,call=thisCall(),called=p.voters.find(v=>v.id===call?.id),active=p.voters.find(v=>v.id===state.armed);
 return `<div class="kiosk-head"><div class="brand"><img src="./assets/logo.svg" alt=""><div><b>${e(state.unit.toUpperCase())} ${e(schoolName)}</b><small>PEMILIHAN KETUA OSIS ${isDemo()?'· DEMO':''}</small></div></div><button class="btn secondary small" data-action="exit-kiosk">${icon('lock')} Panitia</button></div>
 <div class="voting-grid"><aside class="voter-panel"><h2>${icon('users')} Sesi pemilih</h2>
 ${called?`<div class="active-voter"><span class="status dot">${call.status==='active'?'Satu sesi aktif':'Sudah dipanggil'}</span><b>${e(called.name)}</b><small>Kelas ${e(called.classroom)}</small></div><p class="tiny muted">${call.status==='active'?'Satu suara untuk siswa ini. Tunggu konfirmasi sebelum pemilih berikutnya.':'Cocokkan identitas siswa yang hadir sebelum mengaktifkan kartu pilihan.'}</p><button class="btn full-width" data-action="activate-called" ${state.sending||state.queueBusy||!canVote()?'disabled':''}>${icon('rfid')} ${call.status==='active'?'Lanjutkan sesi':'Siswa hadir — buka sesi'}</button><button class="btn secondary full-width mt-20" data-action="repeat-call" ${state.sending||state.queueBusy?'disabled':''}>${icon('sound')} Panggil ulang</button><button class="btn secondary full-width mt-20" data-action="call-next" ${state.sending||state.queueBusy||!canVote()?'disabled':''}>Lanjut / siswa tidak hadir ${icon('arrow')}</button>`:`<p class="tiny muted">Panggil nama dan kelas siswa dari antrean sebelum membuka hak pilih.</p><button class="btn full-width" data-view="calling">${icon('mic')} Buka pemanggilan</button>`}
 <ol class="steps"><li><span><b>Panggil nama dan kelas</b>Gunakan antrean pemanggilan siswa.</span></li><li><span><b>Verifikasi hadir</b>Panitia membuka sesi untuk siswa yang tepat.</span></li><li><span><b>Tempel kartu pilihan</b>Kartu A untuk calon A, kartu B untuk calon B.</span></li><li><span><b>Suara terkunci</b>Siswa yang sudah memilih tidak dapat memilih ulang.</span></li></ol><div class="divider"></div><button class="btn ghost tiny" data-action="test-sound">${icon('sound')} Uji suara notifikasi</button></aside>
 <section class="vote-area"><div class="vote-title"><div><h2>Pilih pemimpinmu.</h2><p>${e(state.unit.toUpperCase())} ${e(schoolName)} · ${e(p.election?.period||'')}</p></div><button class="btn secondary small normal-only" data-action="enter-kiosk">${icon('expand')} Layar bilik</button></div>
 ${state.success?`<div class="success-box" role="status"><div class="success-mark">${icon('check')}</div><h2>Suaramu telah tersimpan!</h2><p>Terima kasih sudah menggunakan hak pilih.<br>Silakan berikan kesempatan kepada pemilih berikutnya.</p><span class="status">1 hak pilih telah digunakan${isDemo()?' · SIMULASI':''}</span></div><div class="normal-only"><button class="btn full-width" data-action="next-voter">Panggil siswa berikutnya ${icon('arrow')}</button></div>`:
 `${!canVote()?`<div class="hint warn">${!current()?'Anda sedang melihat arsip.':p.election?.status!=='open'?'Pemilihan belum dibuka atau sudah ditutup.':!state.online&&!isDemo()?'Internet terputus. Sesi dikunci sampai tersambung kembali.':'Menunggu konfirmasi data dari Firebase.'}</div>`:''}
 <div class="choice-grid">${p.candidates.map(c=>`<article class="choice-card" style="--c:${c.color}"><span class="candidate-number">${e(c.code)}</span>${avatar(c)}<h3>${e(c.name)}</h3><p>${c.deputy?'Wakil: '+e(c.deputy):e(c.classroom||'Calon Ketua OSIS')}</p><div class="rfid-label">${icon('rfid')} Tempel kartu ${e(c.code)}</div></article>`).join('')}</div>
 <div class="scan-box ${active?'':'waiting'}" role="region" aria-label="Pemindai kartu RFID"><div class="scan-icon">${icon(state.sending?'clock':'rfid')}</div><h3>${state.sending?'Menyimpan suara…':active?'Tempelkan satu kartu pilihan':called?'Menunggu verifikasi kehadiran':'Menunggu siswa dipanggil'}</h3><p>${active?'Pastikan kolom di bawah aktif. Reader mengirim ID lalu Enter atau Tab.':'Panitia memanggil dan memverifikasi siswa sebelum RFID diaktifkan.'}</p><form id="scan-form"><label class="sr-only" for="rfid-input">ID kartu RFID pilihan</label><input id="rfid-input" name="rfid" type="text" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="${active?'TEMPELKAN KARTU DI READER':'SESI TERKUNCI'}" maxlength="80" ${active&&!state.sending&&canVote()?'':'disabled'}><button type="submit" class="sr-only" tabindex="-1">Kirim kartu</button></form>${state.scanError?`<p class="form-error" style="margin:12px 0 0">${e(state.scanError)}</p>`:''}
 ${isDemo()&&active?`<div class="demo-simulate">${p.candidates.map(c=>`<button class="btn secondary small" data-action="simulate-card" data-id="${e(c.id)}" ${state.sending?'disabled':''}>Simulasi kartu ${e(c.code)}</button>`).join('')}</div>`:''}</div>`}</section></div>`;
}
function renderResults(){
 const p=state.packet,t=totals(), ranks=[...p.candidates].sort((a,b)=>(t.counts[b.id]||0)-(t.counts[a.id]||0));
 const top=ranks[0],tie=ranks.length>1 && t.counts[ranks[0].id]===t.counts[ranks[1].id];
 const classes=[...new Set(p.voters.map(v=>v.classroom))].sort(), ids=new Set(p.ballots.map(b=>b.id));
 return `<div class="print-title"><h1>Pemilihan OSIS ${e(state.unit.toUpperCase())} ${e(schoolName)}</h1><p>Periode ${e(p.election?.period||'')} · ${e(STATUSES[p.election?.status]||'')} ${isDemo()?'· DATA DEMO':''}</p></div>
 <div class="toolbar"><p class="muted tiny mb-0">${p.election?.status==='closed'?'Hasil akhir pemilihan':'Perolehan sementara; hasil dapat berubah selama pemilihan dibuka.'}</p><div class="btn-row"><button class="btn secondary" data-action="export-results">${icon('download')} Unduh rekap CSV</button><button class="btn" data-action="print-results">${icon('print')} Cetak / PDF</button></div></div>${stats()}
 ${t.unrecognized?`<div class="hint error">Ada ${number(t.unrecognized)} suara dengan calon yang tidak ditemukan. Periksa data sebelum menetapkan hasil.</div>`:''}
 <div class="results-grid"><div><section class="panel"><div class="panel-head"><div><h2>Grafik perolehan suara</h2><p>${number(t.total)} suara terkonfirmasi · ${e(state.unit.toUpperCase())}</p></div>${statusBadge()}</div>${barChart(p.candidates,t.counts)}<div class="sr-only"><table><caption>Jumlah suara setiap calon</caption><thead><tr><th>Calon</th><th>Suara</th></tr></thead><tbody>${p.candidates.map(c=>`<tr><td>${e(c.name)}</td><td>${t.counts[c.id]||0}</td></tr>`).join('')}</tbody></table></div></section>
 <section class="panel"><div class="panel-head"><div><h2>Partisipasi per kelas</h2><p>Sudah memilih dibandingkan jumlah pemilih terdaftar</p></div>${icon('school')}</div>${classes.map(c=>{const all=p.voters.filter(v=>v.classroom===c),done=all.filter(v=>ids.has(v.id)).length,percent=percentage(done,all.length);return `<div class="class-row"><b>${e(c)}</b><div class="progress" aria-label="${e(c)} ${pct(percent)}"><span style="--pct:${percent}%"></span></div><span class="right">${done}/${all.length} · ${pct(percent)}</span></div>`;}).join('')||'<p class="tiny muted">Belum ada kelas terdaftar.</p>'}</section></div>
 <div><section class="panel"><div class="panel-head"><div><h2>Partisipasi pemilih</h2><p>Setiap suara berarti</p></div></div><div class="donut" style="--pct:${t.participation}" role="img" aria-label="Partisipasi ${pct(t.participation)}"><div><b>${pct(t.participation)}</b><small>sudah memilih</small></div></div><div class="legend"><span><i style="--c:var(--green)"></i>Memilih (${number(t.voted)})</span><span><i style="--c:#dfe9e2"></i>Belum (${number(t.remaining)})</span></div></section>
 <section class="panel"><div class="panel-head"><div><h2>${p.election?.status==='closed'?'Hasil akhir':'Urutan sementara'}</h2><p>${!t.total?'Belum ada suara':tie?'Perolehan teratas sama':p.election?.status==='closed'?'Perolehan tertinggi: Calon '+e(top?.code):'Hasil belum ditetapkan'}</p></div>${icon('trophy')}</div>${ranks.map(c=>`<div class="rank-item"><span class="rank-num" style="--c:${c.color}">${e(c.code)}</span><div class="rank-info"><b>${e(c.name)}</b><small>Calon ${e(c.code)}</small></div><div class="rank-value"><strong>${number(t.counts[c.id]||0)}</strong><small>${pct(percentage(t.counts[c.id]||0,t.total))}</small></div></div>`).join('')}</section></div></div>`;
}
function renderSettings(){
 const p=state.packet,st=p.election?.status;
 return `<div class="settings-grid"><div><section class="panel"><div class="panel-head"><div><h2>Kontrol pemilihan</h2><p>Khusus ${e(state.unit.toUpperCase())} ${e(schoolName)}</p></div>${statusBadge()}</div>
 ${p.election?`<div class="setting-item"><div><b>Periode ${e(p.election.period)}</b><p>${current()?'Periode aktif sekolah ini':'Arsip periode; sesi suara dinonaktifkan'}</p></div>${icon('calendar')}</div><div class="setting-item"><div><b>${st==='open'?'Tutup pemilihan':'Buka pemilihan'}</b><p>${st==='draft'?'Lengkapi calon, kartu RFID, dan daftar pemilih.':st==='closed'?'Buka kembali jika masih ada siswa yang belum memilih.':'Hentikan penerimaan suara dan simpan hasil akhir.'}</p></div><button class="btn ${st==='open'?'secondary':''} small" data-action="toggle-election" ${isAdmin()&&current()&&p.ready?'':'disabled'}>${st==='open'?'Tutup':st==='closed'?'Buka kembali':'Buka'}</button></div>`:'<p class="muted tiny">Belum ada periode. Buat periode terlebih dahulu.</p>'}
 <div class="setting-item"><div><b>Periode pemilihan baru</b><p>Periode lama tetap tersimpan sebagai arsip.</p></div><button class="btn secondary small" data-action="new-period" ${isAdmin()?'':'disabled'}>${icon('plus')} Buat</button></div>
 <div class="hint mb-0 mt-20">Saat pemilihan sudah dibuka, data calon dan DPT dikunci. Untuk pemilihan tahun berikutnya, buat periode baru.</div></section>
 <section class="panel"><div class="panel-head"><div><h2>Kartu RFID USB</h2><p>Reader dengan mode keyboard / HID</p></div>${icon('rfid')}</div><p class="tiny muted">Hubungkan reader ke USB komputer. Tempelkan kartu pada kolom uji berikut untuk melihat ID yang dikirim reader. Uji ini tidak mencatat suara.</p><form id="rfid-test-form"><div class="field"><label for="rfid-test">Uji baca kartu</label><input id="rfid-test" name="rfid" type="text" autocomplete="off" maxlength="80" placeholder="Klik, lalu tempelkan kartu"><small id="rfid-test-output">ID disimpan sebagai teks, termasuk angka nol di depan.</small></div></form><label class="checkline"><input id="sound-toggle" type="checkbox" ${state.sound?'checked':''}> Suara notifikasi setelah suara tersimpan</label><button class="btn secondary small" data-action="test-sound">${icon('sound')} Uji suara</button></section></div>
 <div><section class="panel"><div class="panel-head"><div><h2>Koneksi & akun panitia</h2><p>${isDemo()?'Anda sedang mencoba data contoh':'Penyimpanan Firebase sekolah'}</p></div><span class="status ${isDemo()?'amber':''}">${isDemo()?'DEMO':'Firebase'}</span></div>
 <dl class="connection-data"><dt>Penyimpanan aktif</dt><dd>${isDemo()?'Browser ini · data simulasi':'Cloud Firestore'}</dd><dt>Proyek Firebase</dt><dd class="mono">${e(config.firebase?.projectId||'Belum diisi')}</dd><dt>Akses panitia</dt><dd>${isDemo()?'Admin demo':e(state.backend.role?.role||'—')} · ${e(state.backend?.role?.units.map(u=>u.toUpperCase()).join(', ')||'—')}</dd>${!isDemo()?`<dt>Email akun</dt><dd>${e(state.backend.auth?.currentUser?.email||'—')}</dd>`:''}</dl>
 ${isDemo()?`<div class="hint warn">Data demo tidak dikirim ke Firebase. Isi <b>app-config.js</b> dan ikuti panduan pemasangan untuk mulai menggunakan data sekolah.</div><button class="btn full-width" data-action="connect-firebase">${icon('shield')} ${configured?'Masuk ke Firebase':'Panduan sambungkan Firebase'}</button>`:`<button class="btn secondary full-width" data-action="signout">${icon('logout')} Keluar akun panitia</button>`}</section>
 <section class="panel"><div class="panel-head"><div><h2>Rekap & panduan</h2><p>Siapkan pemilihan dengan mudah</p></div>${icon('list')}</div><div class="setting-item"><div><b>Rekap hasil pemilihan</b><p>Jumlah suara dan persentase setiap calon.</p></div><button class="btn secondary small" data-action="export-results" ${p.election?'':'disabled'}>${icon('download')} CSV</button></div><div class="setting-item"><div><b>Status hak pilih siswa</b><p>Daftar siswa, kelas, dan status sudah/belum memilih.</p></div><button class="btn secondary small" data-action="export-voters" ${p.election?'':'disabled'}>${icon('download')} CSV</button></div><div class="setting-item"><div><b>Panduan Firebase & GitHub</b><p>Langkah pemasangan dan petunjuk pemakaian.</p></div><a class="btn secondary small" href="./panduan/PANDUAN.html" target="_blank" rel="noopener">Baca ${icon('arrow')}</a></div></section></div></div>`;
}

function render(){
 document.body.classList.toggle('kiosk',state.kiosk);
 if(!state.logged){renderLogin();return;}
 const focus=document.activeElement;const retain=focus?.id && document.getElementById('app').contains(focus)?{id:focus.id,value:focus.value,start:focus.selectionStart,end:focus.selectionEnd}:null;
 const now=new Date().toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long',year:'numeric',timeZone:'Asia/Jakarta'});
 const subtitle={dashboard:'Pantau perjalanan suara dan partisipasi siswa hari ini.',calling:'Panggil nama dan kelas, lalu verifikasi siswa yang hadir.',voting:'Satu siswa, satu kesempatan untuk menentukan pilihan.',candidates:'Kenali calon pemimpin OSIS dan daftarkan kartu pilihannya.',voters:'Kelola hak pilih siswa secara terpisah untuk setiap jenjang.',results:'Setiap suara dihitung. Setiap pilihan berarti.',settings:'Atur periode, koneksi, dan kebutuhan panitia.'};
 const connection=isDemo()?'<span class="status dot amber">Mode demo</span>':!state.online?'<span class="status dot red">Internet terputus</span>':!state.packet.ready?'<span class="status dot amber">Menghubungkan…</span>':'<span class="status dot">Firebase terhubung</span>';
 let body;
 if(!state.packet.election&&state.view!=='settings')body=empty(state.packet.ready?'Mulai pemilihan OSIS':'Memuat data pemilihan…',state.packet.ready?'Buat periode baru untuk jenjang ini, lalu tambahkan calon dan daftar pemilih.':'Menunggu data dari penyimpanan. Jika terus menunggu, periksa koneksi dan hak akses staff.',`<button class="btn" data-view="settings">${icon('settings')} Buka pengaturan</button>`);
 else body=({dashboard:renderDashboard,calling:renderCalling,voting:renderVoting,candidates:renderCandidates,voters:renderVoters,results:renderResults,settings:renderSettings}[state.view]||renderDashboard)();
 const archiveOptions=state.packet.school?.history||[];
 document.getElementById('app').innerHTML=`<div class="shell"><aside class="sidebar"><div class="brand"><img src="./assets/logo.svg" alt="Logo Al-Miftah"><div><b>AL-MIFTAH</b><small>PEMILIHAN OSIS</small></div></div><div class="nav-caption">Ruang panitia</div><nav class="nav" aria-label="Menu utama">${Object.keys(labels).map(key=>`<button data-view="${key}" class="${state.view===key?'active':''}" ${state.view===key?'aria-current="page"':''} title="${labels[key]}">${icon(navIcons[key])}<span class="nav-label">${labels[key]}</span></button>`).join('')}</nav><div class="nav-bottom"><div class="side-note">${icon('shield')}<b>Suara kecil, perubahan besar.</b><p>Pemilihan yang tertib dimulai dari panitia yang siap.</p><button data-action="help">Panduan penggunaan ${icon('arrow')}</button></div><div class="side-footer"><span>SMP & SMA ${e(schoolName)}</span><span>v1.1</span></div></div></aside>
 <div class="main-wrap"><header class="topbar"><div class="breadcrumb">Ruang panitia <span>/</span><span>${labels[state.view]}</span></div><div class="top-actions">${connection}<span class="separator"></span><button class="icon-btn" data-action="help" title="Panduan">${icon('help')}</button><button class="profile" data-action="profile" title="Akun panitia"><span class="avatar">${isDemo()?'DM':'PN'}</span><span><b>${isDemo()?'Panitia Demo':'Panitia OSIS'}</b><small>${isAdmin()?'Administrator':'Operator bilik'}</small></span></button></div></header>
 <main id="main">${isDemo()?'<div class="demo-banner"><span><b>MODE DEMO</b> · Nama, pemilih, dan suara di layar adalah data contoh di browser ini.</span><button data-action="connect-firebase">Sambungkan Firebase →</button></div>':!state.online?'<div class="demo-banner" role="alert"><span><b>Internet terputus.</b> Sesi suara dikunci sampai koneksi kembali.</span></div>':''}
 <div class="page-head"><div><div class="eyebrow">PEMILIHAN KETUA OSIS</div><h1>${state.view==='dashboard'?'Dashboard pemilihan':labels[state.view]}</h1><p>${subtitle[state.view]}</p></div><div class="school-switch" role="group" aria-label="Pilih jenjang">${UNITS.map(u=>`<button data-unit="${u}" class="${state.unit===u?'active':''}" aria-pressed="${state.unit===u}" ${state.backend.role?.units.includes(u)?'':'disabled'}>${u.toUpperCase()}</button>`).join('')}</div></div>
 <div class="context-row"><div>${statusBadge()}${!current()&&state.packet.election?'<span class="status amber">Arsip periode</span>':''}<span class="period">${icon('calendar')}${archiveOptions.length>1?`<select id="period-select" aria-label="Pilih periode"><option value="">Periode aktif · ${current()?e(state.packet.election?.period||''):'kembali'}</option>${archiveOptions.filter(id=>id!==state.packet.school?.currentElectionId).map((id,i)=>`<option value="${e(id)}" ${state.selectedId===id?'selected':''}>Arsip ${i+1}${state.packet.election?.id===id?' · '+e(state.packet.election.period):''}</option>`).join('')}</select>`:e(state.packet.election?.period||'Belum ada periode')}</span></div><span class="muted tiny">${e(now)}</span></div>
 ${!state.packet.ready&&!isDemo()?'<div class="loading-line" aria-label="Menghubungkan Firebase"></div>':''}${body}
 <footer class="bottom-note"><span>${icon('shield')} Satu siswa, satu suara. Untuk ${e(schoolName)} yang lebih baik.</span><span>${e(state.unit.toUpperCase())} ${e(schoolName)} · ${isDemo()?'Simulasi lokal':'Data Firebase'}</span></footer></main></div></div>`;
 if(retain){const next=document.getElementById(retain.id);if(next&&!next.disabled){if('value'in next)next.value=retain.value;next.focus({preventScroll:true});try{next.setSelectionRange(retain.start,retain.end);}catch{}}}
 if(state.view==='voting'&&state.armed&&!retain)setTimeout(focusScan,0);
}

function renderLogin(){
 document.getElementById('app').innerHTML=`<div class="login-page"><section class="login-visual"><div class="brand"><img src="./assets/logo.svg" alt=""><div><b>AL-MIFTAH</b><small>PEMILIHAN OSIS · SMP & SMA</small></div></div><div><div class="eyebrow" style="color:#9fc6ac">RUANG DEMOKRASI PELAJAR</div><h1>Satu suara.<br><em>Banyak harapan.</em></h1><p>Pemilihan pemimpin OSIS yang lebih tertib, mudah, dan dekat dengan siswa.</p>${ballotArt()}</div><span class="login-credits">SMP & SMA ${e(schoolName)} · Satu siswa, satu suara.</span></section><section class="login-form-area"><div class="login-card"><div class="eyebrow">SELAMAT DATANG, PANITIA</div><h2>Siap memilih masa depan?</h2><p>Masuk untuk mengelola pemilihan OSIS sekolah.</p>${state.booting?'<div class="loading-line"></div>':''}${state.loginError?`<div class="hint error">${e(state.loginError)}</div>`:''}<form id="login-form"><div id="form-error" class="form-error"></div><div class="field"><label for="email">Email panitia</label><input id="email" type="email" name="email" autocomplete="username" placeholder="Email dari Firebase Authentication" required></div><div class="field"><label for="password">Password</label><input id="password" type="password" name="password" autocomplete="current-password" minlength="6" required placeholder="Masukkan password akun panitia"></div><button class="btn" type="submit" ${state.booting||!configured?'disabled':''}>${icon('lock')} Masuk ke ruang panitia</button></form><p class="tiny"><button class="text-btn" data-action="launch-demo">Coba aplikasi dengan data demo</button></p><div class="divider"></div><a class="tiny" href="./panduan/PANDUAN.html" target="_blank" rel="noopener">Panduan pemasangan Firebase dan GitHub →</a></div></section></div>`;
}

function subscribeUnit(){
 state.unsub?.();state.packet=emptyPacket();state.armed=null;state.success=false;
 const unit=state.unit;render();
 state.unsub=state.backend.subscribe(unit,state.selectedId,packet=>{
   const oldId=state.packet.election?.id;
   if(oldId && packet.election?.id && oldId!==packet.election.id){state.armed=null;state.success=false;state.selectedId=null;}
   packet.candidates.sort((a,b)=>a.order-b.order||a.code.localeCompare(b.code));state.packet=packet;
   if(!state.sending){const call=thisCall();state.armed=canVote()&&call?.status==='active'?call.id:null;}
   render();
 },err=>{state.armed=null;report(err);});
}
async function launchDemo(){state.unsub?.();state.backend?.destroy?.();state.backend=new DemoStore();useAccountStation('DEMO');state.logged=true;state.selectedId=null;state.view='dashboard';subscribeUnit();}
async function boot(){
 if(!configured){await launchDemo();return;}
 state.booting=true;render();
 try{state.backend=await createFirebaseStore(config.firebase);state.logged=await state.backend.authorize();
 if(state.logged){useAccountStation(state.backend.uid);if(!state.backend.role.units.includes(state.unit))state.unit=state.backend.role.units[0];subscribeUnit();}}
 catch(error){state.loginError=readableError(error);}
 finally{state.booting=false;render();}
}

function modal(title,body,large=false){
 closeModal();
 document.getElementById('modal-root').innerHTML=`<dialog class="modal ${large?'large':''}" id="app-modal" aria-labelledby="modal-title"><div class="modal-head"><h2 id="modal-title">${e(title)}</h2><button class="icon-btn" type="button" data-action="close-modal" aria-label="Tutup">${icon('close')}</button></div><div class="modal-body">${body}</div></dialog>`;
 const dialog=document.getElementById('app-modal');dialog.showModal();
 dialog.addEventListener('close',()=>{if(document.getElementById('app-modal')===dialog){document.getElementById('modal-root').innerHTML='';state.confirm=null;setTimeout(focusScan,0);}},{once:true});
}
function closeModal(){const d=document.getElementById('app-modal');if(d){d.close();document.getElementById('modal-root').innerHTML='';}state.confirm=null;state.photoToken=null;state.csvToken=null;state.photoLoading=false;}
function formActions(label='Simpan'){return `<div class="modal-actions"><button class="btn secondary" type="button" data-action="close-modal">Batal</button><button class="btn" type="submit">${label}</button></div>`;}
function confirmation(title,text,label,fn,danger=false){
 modal(title,`<p>${e(text)}</p><div class="modal-actions"><button class="btn secondary" data-action="close-modal">Batal</button><button class="btn ${danger?'danger':''}" data-action="confirm-action">${e(label)}</button></div>`);
 state.confirm=fn;
}
function guide(){modal('Cara menggunakan aplikasi',`<div class="hint">${isDemo()?'<b>Mode demo aktif.</b> Data contoh disimpan di browser.':'Data sekolah disimpan di Firebase.'}</div><ol class="guide-list"><li>Pilih jenjang <b>SMP</b> atau <b>SMA</b>. Masing-masing memiliki data dan antrean sendiri.</li><li>Buat periode pada Pengaturan. Tambahkan calon dan daftarkan kartu pilihan A/B.</li><li>Masukkan siswa secara manual atau impor <b>Excel .xlsx / CSV</b>. Cukup kolom <b>nama</b> dan <b>kelas</b>. Data yang sama dilewati.</li><li>Buka pemilihan, lalu masuk ke <b>Pemanggilan</b>. Tekan <b>Panggil berikutnya</b> untuk membacakan nama dan kelas.</li><li>Jika siswa tidak hadir, tekan <b>Lanjut</b>. Ia masuk daftar <b>Dilewati</b> dan tetap dapat dipanggil kembali.</li><li>Jika siswa hadir, cocokkan identitasnya dan tekan <b>Siswa hadir — buka sesi</b>. Hanya siswa ini yang boleh menempelkan kartu.</li><li>Tunggu <b>Suaramu telah tersimpan</b>. Hak pilih terkunci dan siswa tidak dipanggil kembali.</li><li>Tutup pemilihan setelah selesai. Unduh rekap atau cetak PDF.</li></ol><div class="hint warn">Reader harus mengetik ID melalui USB mode keyboard/HID, diakhiri Enter atau Tab. Panitia tetap mengawasi identitas siswa. Uji suara nama/kelas pada komputer sebelum pemilihan.</div><a class="btn full-width" href="./panduan/PANDUAN.html" target="_blank" rel="noopener">Panduan pemasangan dan pembaruan ${icon('arrow')}</a>`,true);}
function candidateForm(id){
 if(!editable())throw new Error('Calon hanya dapat diatur admin pada tahap Persiapan.');
 const existing=state.packet.candidates.find(c=>c.id===id);
 if(!existing && state.packet.candidates.length>=20)throw new Error('Maksimal 20 calon pada satu periode.');
 const nextCode='ABCDEFGHIJKLMNOPQRST'.split('').find(x=>!state.packet.candidates.some(c=>c.code===x))||'';
 const c=existing||{id:uniqueId('calon-'),code:nextCode,name:'',deputy:'',classroom:'',vision:'',mission:'',rfid:'',photo:'',color:COLORS[state.packet.candidates.length%COLORS.length],order:state.packet.candidates.length};
 state.modalPhoto=photoSrc(c.photo);state.photoLoading=false;
 modal(existing?'Ubah calon OSIS':'Tambah calon OSIS',`<form id="candidate-form" data-id="${e(c.id)}" data-order="${c.order}"><div id="form-error" class="form-error"></div><div class="photo-upload"><div class="preview" id="photo-preview">${state.modalPhoto?`<img src="${state.modalPhoto}" alt="Pratinjau foto calon">`:icon('image')}</div><div><label for="candidate-photo" class="upload-label">Foto calon / foto pasangan</label><input id="candidate-photo" type="file" accept="image/jpeg,image/png,image/webp"><div class="upload-help" id="photo-help">JPG, PNG, atau WebP. Foto akan dikompres otomatis.</div><button class="text-btn" type="button" data-action="remove-photo">Hapus foto</button></div></div>
 <div class="field-row"><div class="field"><label for="candidate-code">Kode calon</label><input id="candidate-code" name="code" value="${e(c.code)}" maxlength="4" pattern="[A-Za-z0-9]{1,4}" required placeholder="A"></div><div class="field"><label for="candidate-class">Kelas</label><input id="candidate-class" name="classroom" maxlength="40" value="${e(c.classroom)}" placeholder="VIII A / XI IPA"></div></div>
 <div class="field"><label for="candidate-name">Nama calon ketua</label><input id="candidate-name" name="name" value="${e(c.name)}" required maxlength="120"></div><div class="field"><label for="candidate-deputy">Nama calon wakil (opsional)</label><input id="candidate-deputy" name="deputy" value="${e(c.deputy)}" maxlength="120"></div>
 <div class="field"><label for="candidate-rfid">ID kartu RFID untuk calon ini</label><input id="candidate-rfid" name="rfid" type="text" value="${e(c.rfid)}" required maxlength="80" autocomplete="off" class="mono" placeholder="Klik di sini, tempelkan kartu pilihan"><small>Tekan Enter dari reader hanya mengisi ID pada kolom ini. Angka nol di depan tetap disimpan.</small></div>
 <div class="field"><label for="candidate-vision">Visi</label><textarea id="candidate-vision" name="vision" maxlength="500" rows="2">${e(c.vision)}</textarea></div><div class="field"><label for="candidate-mission">Misi (satu per baris)</label><textarea id="candidate-mission" name="mission" maxlength="1500">${e(c.mission)}</textarea></div><div class="field"><label for="candidate-color">Warna calon</label><select id="candidate-color" name="color">${COLORS.map((color,i)=>`<option value="${color}" ${color===c.color?'selected':''}>${['Hijau','Emas','Biru','Ungu','Koral','Toska'][i]}</option>`).join('')}</select></div>${formActions('Simpan calon')}</form>`,true);
}
function detailCandidate(id){
 const c=state.packet.candidates.find(c=>c.id===id);if(!c)return;
 modal('Profil calon '+c.code,`<div style="border-radius:12px;overflow:hidden;margin-bottom:22px">${avatar(c)}</div><h2>${e(c.name)}</h2>${c.deputy?`<p>Wakil: <b>${e(c.deputy)}</b></p>`:''}<span class="status">${e(c.classroom||state.unit.toUpperCase())}</span><h3 class="mt-20">Visi</h3><p>${e(c.vision||'Belum diisi.')}</p><h3>Misi</h3><ul class="guide-list">${(c.mission||'Belum diisi.').split('\n').filter(Boolean).map(m=>`<li>${e(m)}</li>`).join('')}</ul>${isAdmin()?`<div class="hint"><b>ID kartu ${e(c.code)}:</b> <span class="mono">${e(c.rfid)}</span></div>`:''}`);
}
function voterForm(){
 if(!editable())throw new Error('DPT hanya dapat diubah admin pada tahap Persiapan.');
 modal('Tambah pemilih',`<form id="voter-form"><div id="form-error" class="form-error"></div><div class="field"><label for="student-name">Nama lengkap siswa</label><input id="student-name" name="name" maxlength="120" required></div><div class="field"><label for="student-class">Kelas</label><input id="student-class" name="classroom" maxlength="40" required placeholder="VII A / X A"></div><div class="field"><label for="student-code">Kode / NIS (opsional)</label><input id="student-code" name="code" type="text" maxlength="80" placeholder="Boleh dikosongkan"></div><div class="hint">Cukup nama dan kelas. Data yang sama tidak ditambahkan dua kali.</div>${formActions('Simpan pemilih')}</form>`);
}
function importForm(){
 if(!editable())throw new Error('DPT hanya dapat diimpor saat tahap Persiapan.');
 state.importRows=[];state.importDuplicates=0;
 modal('Impor Excel / CSV',`<form id="import-form"><div id="form-error" class="form-error"></div><p>Gunakan dua kolom: <b>nama</b> dan <b>kelas</b>, pada lembar pertama. Kode/NIS boleh ditambahkan tetapi tidak wajib. Nama dan kelas yang sama akan dilewati.</p><div class="btn-row"><button type="button" class="btn secondary small" data-action="download-template">${icon('download')} Template Excel</button><button type="button" class="btn secondary small" data-action="download-csv-template">Template CSV</button></div><div class="field mt-20"><label for="csv-file">Pilih file Excel (.xlsx) atau CSV</label><input id="csv-file" type="file" accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" required><small>Maksimal 5 MB atau 5.000 siswa. File .xls lama: Save As .xlsx terlebih dahulu.</small></div><div id="import-preview" class="hint">Pilih file untuk memeriksa nama dan kelas sebelum disimpan.</div><div id="import-progress" class="tiny muted"></div>${formActions('Impor ke '+state.unit.toUpperCase())}</form>`,true);
}
async function compressPhoto(file){
 if(!['image/jpeg','image/png','image/webp'].includes(file.type))throw new Error('Gunakan foto JPG, PNG, atau WebP.');
 if(file.size>10*1024*1024)throw new Error('Foto terlalu besar. Pilih foto di bawah 10 MB.');
 const url=URL.createObjectURL(file), image=new Image();
 try{image.src=url;await image.decode();const scale=Math.min(1,600/image.naturalWidth,800/image.naturalHeight);const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(image.naturalWidth*scale));canvas.height=Math.max(1,Math.round(image.naturalHeight*scale));
 const ctx=canvas.getContext('2d');ctx.fillStyle='#f0f5f0';ctx.fillRect(0,0,canvas.width,canvas.height);ctx.drawImage(image,0,0,canvas.width,canvas.height);
 for(let quality=.82;quality>=.22;quality-=.1){const data=canvas.toDataURL('image/jpeg',quality);if(data.length<=110000)return data;}
 throw new Error('Foto masih terlalu besar setelah kompresi. Potong atau kecilkan foto, lalu coba lagi.');
 }finally{URL.revokeObjectURL(url);}
}
function download(filename,text,type='text/csv;charset=utf-8'){
 const blob=new Blob([text],{type}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function numberedRows(rows){let order=Math.max(state.packet.voters.length,...state.packet.voters.map(v=>v.queueNo||0));return rows.map(v=>({...v,queueNo:++order}));}
function downloadExcelTemplate(){const a=document.createElement('a');a.href='./panduan/Template_Pemilih_'+state.unit.toUpperCase()+'.xlsx';a.download='Template_Pemilih_'+state.unit.toUpperCase()+'.xlsx';document.body.append(a);a.click();a.remove();}
function exportResults(){
 const p=state.packet,t=totals();electionId();
 const rows=[['REKAP PEMILIHAN OSIS',state.unit.toUpperCase(),schoolName],['Periode',p.election.period],['Status',STATUSES[p.election.status]],['Jenis data',isDemo()?'DEMO / SIMULASI':'FIREBASE'],['Pemilih terdaftar',t.registered],['Suara masuk',t.total],['Belum memilih',t.remaining],['Partisipasi (%)',t.participation],[],['Kode calon','Nama ketua','Nama wakil','Suara','Persentase (%)']];
 for(const c of p.candidates)rows.push([c.code,c.name,c.deputy,t.counts[c.id]||0,percentage(t.counts[c.id]||0,t.total)]);
 download(`Rekap_OSIS_${state.unit.toUpperCase()}${isDemo()?'_DEMO':''}.csv`,makeCSV(rows));
}
function exportVoters(){const ids=new Set(state.packet.ballots.map(b=>b.id));electionId();download(`Status_Pemilih_${state.unit.toUpperCase()}${isDemo()?'_DEMO':''}.csv`,makeCSV([['kode','nama','kelas','status'],...state.packet.voters.map(v=>[v.code,v.name,v.classroom,ids.has(v.id)?'Sudah memilih':'Belum memilih'])]));}
function say(text){
 if(!window.speechSynthesis){toast('Suara tidak tersedia di browser ini. Nama dan kelas tetap ditampilkan.',true);return;}
 const synth=window.speechSynthesis,voices=synth.getVoices?.()||[];
 synth.cancel();synth.resume?.();const utterance=new SpeechSynthesisUtterance(text);utterance.lang='id-ID';utterance.rate=.9;
 const voice=voices.find(v=>v.voiceURI===state.voiceURI)||voices.find(v=>/^id([_-]|$)/i.test(v.lang));if(voice)utterance.voice=voice;
 utterance.onerror=event=>{if(!['canceled','interrupted'].includes(event.error))toast('Suara belum terdengar. Klik Panggil ulang dan periksa volume/voice Indonesia.',true);};
 synth.speak(utterance);
}
function speak(){if(state.sound)say('Suara berhasil disimpan. Terima kasih.');}
function announce(voter){if(state.callSound)say(callText(voter,state.unit,schoolName));}
async function callOne(id){
 if(state.queueBusy||state.sending)return false;if(!canVote())throw new Error('Pemilihan belum siap menerima panggilan.');
 const voter=state.packet.voters.find(v=>v.id===id);if(!voter)throw new Error('Siswa tidak ditemukan.');
 if(thisCall()?.id===id){announce(voter);return true;}
 state.queueBusy=true;state.scanError='';state.success=false;render();
 try{await state.backend.callVoter(electionId(),id,state.stationId);state.armed=null;state.view='calling';announce(voter);return true;}
 finally{state.queueBusy=false;render();}
}
async function callNext(status='waiting'){
 if(state.queueBusy||state.sending)return;
 const rows=queueRows(state.packet,{classroom:state.callClass,status});
 for(const v of rows.slice(0,12)){
   try{if(await callOne(v.id))return;}
   catch(error){if(error.code!=='queue/busy')throw error;}
 }
 if(rows.length)throw new Error('Antrean sedang berubah di bilik lain. Coba tombol Lanjut lagi.');
 const call=thisCall();if(call){state.queueBusy=true;render();try{await state.backend.skipVoter(electionId(),call.id,state.stationId);state.armed=null;window.speechSynthesis?.cancel();}finally{state.queueBusy=false;}}
 state.view='calling';state.success=false;render();toast(status==='skipped'?'Tidak ada siswa yang dilewati pada kelas ini.':'Antrean menunggu sudah selesai. Siswa Dilewati masih dapat dipanggil kembali.');
}
async function activateCalled(){
 const call=thisCall();if(!call||!canVote())throw new Error('Panggil siswa terlebih dahulu dan periksa koneksi.');
 if(state.queueBusy||state.sending)return;state.queueBusy=true;render();
 try{await state.backend.activateVoter(electionId(),call.id,state.stationId);state.armed=call.id;state.success=false;state.scanError='';state.view='voting';}
 finally{state.queueBusy=false;render();focusScan();}
}
async function skipOnly(){const call=thisCall();if(!call||state.sending||state.queueBusy)return;state.queueBusy=true;render();try{await state.backend.skipVoter(electionId(),call.id,state.stationId);state.armed=null;state.view='calling';window.speechSynthesis?.cancel();toast('Siswa dilewati. Hak pilih tetap tersedia.');}finally{state.queueBusy=false;render();}}
async function castRFID(raw){
 if(state.sending)return;
 if(!state.armed)throw new Error('Panitia harus membuka satu sesi pemilih terlebih dahulu.');
 if(!canVote())throw new Error('Pemilihan belum siap menerima suara. Periksa status dan koneksi.');
 let candidate;
 try{candidate=candidateForRFID(state.packet.candidates,raw);}
 catch(error){state.scanError=error.message;render();focusScan();return;}
 const studentId=state.armed,eid=electionId();state.sending=true;state.scanError='';render();
 try{
   await state.backend.castVote(eid,studentId,candidate.id,state.stationId);
   state.armed=null;state.success=true;state.scanError='';speak();
 }catch(error){state.armed=null;state.success=false;state.scanError=readableError(error)+' Periksa status siswa di Daftar Pemilih sebelum membuka sesi kembali.';report(error);}
 finally{state.sending=false;render();}
}
async function enterKiosk(){state.kiosk=true;render();try{await document.documentElement.requestFullscreen?.();}catch{/* Tampilan bilik tetap tersedia tanpa fullscreen browser. */}focusScan();}
async function leaveKiosk(){state.kiosk=false;closeModal();if(document.fullscreenElement)await document.exitFullscreen?.();render();}
function requestExitKiosk(){
 if(state.sending)throw new Error('Tunggu sampai penyimpanan suara selesai.');
 if(isDemo()){leaveKiosk().catch(report);return;}
 modal('Kembali ke ruang panitia',`<form id="reauth-form"><div id="form-error" class="form-error"></div><p>Masukkan password akun panitia untuk keluar dari tampilan bilik.</p><div class="field"><label for="reauth-password">Password panitia</label><input type="password" id="reauth-password" name="password" autocomplete="current-password" required></div>${formActions('Buka ruang panitia')}</form>`);
}
function openPeriodForm(){
 requireNoSession();if(!isAdmin())throw new Error('Hanya admin yang dapat membuat periode.');
 modal('Buat periode pemilihan',`<form id="period-form"><div id="form-error" class="form-error"></div><p>Buat pemilihan baru untuk <b>${e(state.unit.toUpperCase())} ${e(schoolName)}</b>. Calon, daftar pemilih, dan suara dimulai dari kosong. Periode lama tetap dapat dilihat melalui menu arsip.</p><div class="field"><label for="new-period">Nama periode</label><input id="new-period" name="period" required maxlength="60" placeholder="Contoh: 2026/2027"></div><div class="hint">Tutup periode aktif sebelum membuat periode berikutnya.</div>${formActions('Buat periode')}</form>`);
}
function toggleElection(){
 requireNoSession();if(!isAdmin()||!current()||!state.packet.ready)throw new Error('Kontrol pemilihan hanya tersedia bagi admin pada periode aktif.');
 const p=state.packet,open=p.election.status!=='open';
 if(open){validateCandidates(p.candidates);if(!p.voters.length)throw new Error('Tambahkan daftar pemilih terlebih dahulu.');}
 confirmation(open?'Buka pemilihan?':'Tutup pemilihan?',open?'Pastikan calon, ID RFID, dan daftar siswa sudah benar. Setelah dibuka, data calon dan DPT dikunci untuk periode ini.':'Penerimaan suara akan dihentikan. Suara yang sudah tersimpan tetap dipertahankan dan dapat direkap.',open?'Buka pemilihan':'Tutup pemilihan',async()=>{await state.backend.setStatus(electionId(),open?'open':'closed');toast(open?'Pemilihan dibuka. Bilik suara siap digunakan.':'Pemilihan ditutup. Hasil tersimpan.');},!open);
}
async function connectFirebase(){
 requireNoSession();
 if(!configured){guide();return;}
 state.unsub?.();state.backend?.destroy?.();state.logged=false;state.packet=emptyPacket();state.backend=null;state.loginError='';state.kiosk=false;await boot();
}
async function action(key,button){
 const id=button?.dataset.id;
 if(state.kiosk && !['exit-kiosk','close-modal','confirm-action',...(isDemo()?['simulate-card']:[])].includes(key))return;
 switch(key){
 case 'help':guide();break;
 case 'close-modal':closeModal();break;
 case 'profile':requireNoSession();navigate('settings');break;
 case 'launch-demo':await launchDemo();break;
 case 'connect-firebase':await connectFirebase();break;
 case 'add-candidate':candidateForm();break;
 case 'edit-candidate':candidateForm(id);break;
 case 'candidate-detail':detailCandidate(id);break;
 case 'delete-candidate':{if(!editable())throw new Error('Data calon sedang dikunci.');const c=state.packet.candidates.find(c=>c.id===id);if(c)confirmation('Hapus calon '+c.code+'?',c.name+' akan dihapus dari periode ini.','Hapus calon',async()=>{await state.backend.deleteCandidate(electionId(),id);toast('Calon dihapus.');},true);break;}
 case 'delete-voter':{if(!editable())throw new Error('DPT sedang dikunci.');const v=state.packet.voters.find(v=>v.id===id);if(v)confirmation('Hapus pemilih?',v.name+' akan dihapus dari daftar pemilih periode ini.','Hapus pemilih',async()=>{await state.backend.deleteVoter(electionId(),id);toast('Pemilih dihapus.');},true);break;}
 case 'remove-photo':state.photoToken=null;state.photoLoading=false;state.modalPhoto='';document.getElementById('photo-preview').innerHTML=icon('image');document.getElementById('candidate-photo').value='';document.getElementById('photo-help').textContent='Foto dihapus. Simpan calon untuk menerapkan.';break;
 case 'add-voter':voterForm();break;
 case 'import-voters':importForm();break;
 case 'download-template':downloadExcelTemplate();break;
 case 'download-csv-template':download(`Template_Pemilih_${state.unit.toUpperCase()}.csv`,makeCSV([['nama','kelas'],['Nama Siswa Pertama',state.unit==='smp'?'VII A':'X A'],['Nama Siswa Kedua',state.unit==='smp'?'VII A':'X A']]));break;
 case 'export-results':exportResults();break;
 case 'export-voters':exportVoters();break;
 case 'print-results':window.print();break;
 case 'previous-page':state.page=Math.max(1,state.page-1);render();break;
 case 'next-page':state.page++;render();break;
 case 'call-next':await callNext();break;
 case 'call-skipped':await callNext('skipped');break;
 case 'call-specific':await callOne(id);break;
 case 'repeat-call':{const call=thisCall(),voter=state.packet.voters.find(v=>v.id===call?.id);if(voter)announce(voter);break;}
 case 'activate-called':await activateCalled();break;
 case 'skip-only':case 'cancel-session':await skipOnly();break;
 case 'next-voter':state.success=false;state.view='calling';await callNext();break;
 case 'call-previous-page':state.callPage=Math.max(1,state.callPage-1);render();break;
 case 'call-next-page':state.callPage++;render();break;
 case 'test-call-sound':say(callText({name:'Ahmad',classroom:state.unit==='smp'?'VII A':'X A'},state.unit,schoolName));break;
 case 'release-call':if(isAdmin())confirmation('Lepas panggilan bilik lain?','Pastikan siswa belum sedang menempelkan kartu. Panggilan akan ditunda dan dapat dipanggil kembali; suara yang sudah tersimpan tetap terkunci.','Lepas panggilan',async()=>{await state.backend.releaseCall(electionId(),id);toast('Panggilan dilepas ke daftar Dilewati.');});break;
 case 'simulate-card':if(isDemo()){const c=state.packet.candidates.find(c=>c.id===id);if(c)await castRFID(c.rfid);}break;
 case 'test-sound':speak();if(!state.sound)toast('Aktifkan suara notifikasi di Pengaturan terlebih dahulu.');break;
 case 'enter-kiosk':await enterKiosk();break;
 case 'exit-kiosk':requestExitKiosk();break;
 case 'new-period':openPeriodForm();break;
 case 'toggle-election':toggleElection();break;
 case 'confirm-action':{const fn=state.confirm;if(fn){closeModal();button.disabled=true;await fn();}break;}
 case 'signout':requireNoSession();state.unsub?.();await state.backend.signOut();state.logged=false;state.packet=emptyPacket();state.kiosk=false;render();break;
 }
}

document.addEventListener('click',event=>{
 const button=event.target.closest('button');if(!button||button.disabled)return;
 if(button.dataset.view){try{if(!state.kiosk)navigate(button.dataset.view);}catch(error){report(error);}return;}
 if(button.dataset.unit){try{if(state.kiosk)return;requireNoSession();if(state.backend.role.units.includes(button.dataset.unit)){state.unit=button.dataset.unit;state.selectedId=null;state.search='';state.voterSearch='';state.classFilter='all';state.statusFilter='all';state.page=1;state.callClass='all';state.callFilter='all';state.callSearch='';state.callPage=1;safeSave('osis-unit',state.unit);subscribeUnit();}}catch(error){report(error);}return;}
 if(button.dataset.action)action(button.dataset.action,button).catch(report);
});

document.addEventListener('submit',async event=>{
 const form=event.target;
 if(!['login-form','scan-form','rfid-test-form','candidate-form','voter-form','period-form','import-form','reauth-form'].includes(form.id))return;
 event.preventDefault();if(form.dataset.busy==='1')return;
 const fd=new FormData(form),submit=form.querySelector('[type=submit]');
 form.dataset.busy='1';if(submit)submit.disabled=true;
 const errorBox=document.getElementById('form-error');if(errorBox)errorBox.textContent='';
 try{
  switch(form.id){
  case 'login-form':{
    if(!configured)throw new Error('Isi app-config.js dengan konfigurasi Firebase terlebih dahulu.');
    if(state.backend?.mode!=='firebase')state.backend=await createFirebaseStore(config.firebase);
    await state.backend.signIn(String(fd.get('email')).trim(),String(fd.get('password')));
    state.logged=true;state.loginError='';useAccountStation(state.backend.uid);if(!state.backend.role.units.includes(state.unit))state.unit=state.backend.role.units[0];state.selectedId=null;subscribeUnit();break;
  }
  case 'scan-form':await castRFID(String(fd.get('rfid')||''));break;
  case 'rfid-test-form':{const raw=normalizeRFID(fd.get('rfid'));document.getElementById('rfid-test-output').textContent=raw?'ID terbaca: '+raw+' — '+raw.length+' karakter (teks)':'Belum ada ID terbaca.';break;}
  case 'candidate-form':{
    if(!editable())throw new Error('Data calon dikunci.');if(state.photoLoading)throw new Error('Tunggu kompresi foto selesai.');
    const candidate={id:form.dataset.id,order:Number(form.dataset.order),code:String(fd.get('code')).trim().toUpperCase(),name:String(fd.get('name')).trim(),deputy:String(fd.get('deputy')).trim(),classroom:String(fd.get('classroom')).trim(),vision:String(fd.get('vision')).trim(),mission:String(fd.get('mission')).trim(),rfid:normalizeRFID(fd.get('rfid')),photo:state.modalPhoto,color:String(fd.get('color'))};
    if(!candidate.name||!candidate.code||!candidate.rfid)throw new Error('Kode, nama, dan kartu RFID calon wajib diisi.');
    if(state.packet.candidates.some(c=>c.id!==candidate.id && (c.code===candidate.code||normalizeRFID(c.rfid)===candidate.rfid)))throw new Error('Kode calon atau ID RFID sudah dipakai calon lain di jenjang/periode ini.');
    await state.backend.saveCandidate(electionId(),candidate);closeModal();toast('Calon dan kartu RFID tersimpan.');break;
  }
  case 'voter-form':{
    if(!editable())throw new Error('DPT sedang dikunci.');const row=await studentFromNames(fd.get('name'),fd.get('classroom'),fd.get('code'));
    if(!newStudentsOnly([row],state.packet.voters).length)throw new Error('Nama dan kelas atau kode siswa ini sudah terdaftar. Data tidak ditambahkan ulang.');
    await state.backend.addVoters(electionId(),numberedRows([row]));closeModal();toast('Siswa ditambahkan ke daftar pemilih.');break;
  }
  case 'period-form':{
    const period=String(fd.get('period')).trim();if(!period)throw new Error('Isi nama periode.');
    const id=await state.backend.createElection(state.unit,period);state.selectedId=id;closeModal();state.view='candidates';subscribeUnit();toast('Periode baru siap. Tambahkan calon dan daftar pemilih.');break;
  }
  case 'import-form':{
    if(!editable())throw new Error('DPT sedang dikunci.');if(!state.importRows.length)throw new Error('Pilih file Excel/CSV yang valid terlebih dahulu.');
    const fresh=newStudentsOnly(state.importRows,state.packet.voters);
    if(!fresh.length)throw new Error('Semua siswa sudah terdaftar. Tidak ada hak pilih baru yang ditambahkan.');
    const count=await state.backend.addVoters(electionId(),numberedRows(fresh),(done,total)=>{const el=document.getElementById('import-progress');if(el)el.textContent=`Menyimpan ${done}/${total} siswa… Jangan tutup halaman.`;});
    closeModal();state.importRows=[];toast(number(count)+' siswa berhasil ditambahkan.');break;
  }
  case 'reauth-form':await state.backend.reauthenticate(String(fd.get('password')));await leaveKiosk();break;
  }
 }catch(error){report(error);}
 finally{form.dataset.busy='0';if(submit)submit.disabled=false;}
});

document.addEventListener('change',async event=>{
 const el=event.target;
 try{
  if(el.id==='period-select'){requireNoSession();state.selectedId=el.value||null;state.search='';state.voterSearch='';state.classFilter='all';state.statusFilter='all';subscribeUnit();}
  if(el.id==='class-filter'){state.classFilter=el.value;state.page=1;render();}
  if(el.id==='status-filter'){state.statusFilter=el.value;state.page=1;render();}
  if(el.id==='call-class'){state.callClass=el.value;state.callPage=1;render();}
  if(el.id==='call-filter'){state.callFilter=el.value;state.callPage=1;render();}
  if(el.id==='call-sound-toggle'){state.callSound=el.checked;safeSave('osis-call-sound',state.callSound?'1':'0');if(!state.callSound)window.speechSynthesis?.cancel();}
  if(el.id==='call-voice'){state.voiceURI=el.value;safeSave('osis-voice',el.value);}
  if(el.id==='sound-toggle'){state.sound=el.checked;safeSave('osis-sound',state.sound?'1':'0');}
  if(el.id==='candidate-photo'){
    const file=el.files[0];if(!file)return;const token=uniqueId('foto-');state.photoToken=token;state.photoLoading=true;el.disabled=true;document.getElementById('photo-help').textContent='Mengompres foto…';
    try{const photo=await compressPhoto(file);if(state.photoToken===token&&document.getElementById('candidate-form')){state.modalPhoto=photo;document.getElementById('photo-preview').innerHTML=`<img src="${photo}" alt="Pratinjau foto calon">`;document.getElementById('photo-help').textContent='Foto siap disimpan · ±'+Math.round(photo.length*.75/1024)+' KB';}}
    finally{if(state.photoToken===token)state.photoLoading=false;el.disabled=false;}
  }
  if(el.id==='csv-file'){
    state.importRows=[];const file=el.files[0];if(!file)return;const token=uniqueId('csv-');state.csvToken=token;
    if(file.size>5*1024*1024)throw new Error('File terlalu besar. Maksimal 5 MB.');
    let parsed,sheetName='CSV';if(/\.xlsx$/i.test(file.name)){const excel=await readExcel(await file.arrayBuffer());parsed=await parseStudentRows(excel.rows);sheetName=excel.sheetName;}else if(/\.csv$/i.test(file.name))parsed=await parseStudentCSV(await file.text());else throw new Error('Gunakan .xlsx atau CSV. File .xls lama perlu Save As .xlsx.');
    if(state.csvToken!==token)return;
    const {rows,sourceCount}=parsed;state.importRows=rows;const fresh=newStudentsOnly(rows,state.packet.voters);state.importDuplicates=sourceCount-fresh.length;
    const errorBox=document.getElementById('form-error');if(errorBox)errorBox.textContent='';
    document.getElementById('import-preview').innerHTML=`<b>${sourceCount} baris siswa</b> · ${fresh.length} data baru · ${state.importDuplicates} duplikat/sudah terdaftar dilewati.<br>Lembar: ${e(sheetName)} · Tujuan: <b>${e(state.unit.toUpperCase())}</b> · Periode ${e(state.packet.election.period)}.<div class="table-wrap mt-20"><table><thead><tr><th>Nama</th><th>Kelas</th></tr></thead><tbody>${fresh.slice(0,5).map(v=>`<tr><td>${e(v.name)}</td><td>${e(v.classroom)}</td></tr>`).join('')}</tbody></table></div>`;
  }
 }catch(error){report(error);}
});

let searchTimer;
document.addEventListener('input',event=>{
 const el=event.target;
 if(el.id==='voter-search'){state.search=el.value;state.page=1;clearTimeout(searchTimer);searchTimer=setTimeout(render,140);}
 if(el.id==='call-search'){state.callSearch=el.value;state.callPage=1;clearTimeout(searchTimer);searchTimer=setTimeout(render,140);}
});
document.addEventListener('keydown',event=>{
 const el=event.target;
 if(el.id==='candidate-rfid'&&(event.key==='Enter'||event.key==='Tab')){if(event.key==='Enter')event.preventDefault();el.value=normalizeRFID(el.value);if(event.key==='Enter')document.getElementById('candidate-vision')?.focus();}
 if(['rfid-input','rfid-test'].includes(el.id)&&event.key==='Tab'&&el.value.trim()){event.preventDefault();el.form?.requestSubmit();}
});
window.addEventListener('offline',()=>{state.online=false;if(!isDemo()&&state.armed&&!state.sending){state.armed=null;state.scanError='Internet terputus. Buka sesi kembali setelah tersambung.';}render();});
window.addEventListener('online',()=>{state.online=true;render();});
window.addEventListener('beforeunload',event=>{if(state.armed||state.sending){event.preventDefault();event.returnValue='';}});
window.speechSynthesis?.addEventListener?.('voiceschanged',()=>{const select=document.getElementById('call-voice');if(select)select.innerHTML=voiceOptions();});
boot().catch(error=>{state.booting=false;state.logged=false;state.loginError=readableError(error);render();});
