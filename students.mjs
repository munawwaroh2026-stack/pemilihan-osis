import {normalizeCode,parseCSV,percentage} from './core.mjs';
export const CALL_LABELS={waiting:'Menunggu',called:'Dipanggil',active:'Sedang memilih',skipped:'Dilewati',voted:'Sudah memilih'};
const clean=v=>String(v??'').normalize('NFKC').replace(/\s+/g,' ').trim();
export function studentIdentity(name,classroom){
 const nameKey=clean(name).toLocaleUpperCase('id-ID').replace(/[^\p{L}\p{N}]+/gu,' ').trim().replace(/\s+/g,' ');
 const classKey=clean(classroom).toLocaleUpperCase('id-ID').replace(/[^\p{L}\p{N}]/gu,'');
 return nameKey && classKey ? nameKey+'|'+classKey : '';
}
export async function studentFromNames(name,classroom,code=''){
 name=clean(name);classroom=clean(classroom).toLocaleUpperCase('id-ID');
 if(!name || name.length>120 || !classroom || classroom.length>40)throw new Error('Nama dan kelas wajib diisi (nama maksimal 120, kelas 40 karakter).');
 const identity=studentIdentity(name,classroom);
 if(!identity)throw new Error('Nama dan kelas harus berisi huruf atau angka.');
 if(!globalThis.crypto?.subtle)throw new Error('Buka aplikasi melalui HTTPS GitHub Pages atau localhost untuk membuat identitas pemilih.');
 const hash=new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(identity)));
 const digest=Array.from(hash,b=>b.toString(16).padStart(2,'0')).join('');
 const normalizedCode=normalizeCode(code);
 if(normalizedCode.length>80)throw new Error('Kode/NIS maksimal 80 karakter.');
 return {id:'nama-'+digest,code:normalizedCode||'AUTO-'+digest.slice(0,12).toUpperCase(),name,classroom};
}
export async function parseStudentRows(matrix){
 const data=matrix.filter(row=>row.some(c=>String(c??'').trim()));
 if(data.length<2)throw new Error('File harus berisi judul kolom nama, kelas dan minimal satu siswa.');
 const headers=data[0].map(v=>clean(v).toLowerCase());
 const nameIndex=headers.findIndex(h=>['nama','nama siswa','nama lengkap'].includes(h));
 const classIndex=headers.findIndex(h=>['kelas','rombel','kelas siswa'].includes(h));
 const codeIndex=headers.findIndex(h=>['kode','nis','nisn','kode pemilih'].includes(h));
 if(nameIndex<0||classIndex<0)throw new Error('Kolom wajib adalah nama dan kelas. Kode/NIS boleh dikosongkan.');
 if(data.length>5001)throw new Error('Maksimal 5.000 siswa per file.');
 const rows=[],identities=new Set(),codes=new Map();let duplicateCount=0;
 for(let i=1;i<data.length;i++){
   let student;try{student=await studentFromNames(data[i][nameIndex],data[i][classIndex],codeIndex<0?'':data[i][codeIndex]);}
   catch(error){throw new Error(`Baris ${i+1}: ${error.message}`);}
   if(identities.has(student.id)){duplicateCount++;continue;}
   if(!student.code.startsWith('AUTO-')&&codes.has(student.code)&&codes.get(student.code)!==student.id)throw new Error(`Kode/NIS ${student.code} dipakai oleh dua nama/kelas berbeda. Periksa DPT.`);
   identities.add(student.id);codes.set(student.code,student.id);rows.push(student);
 }
 return {rows,duplicateCount,sourceCount:data.length-1};
}
export async function parseStudentCSV(text){return parseStudentRows(parseCSV(text));}
export function newStudentsOnly(rows,existing){
 const ids=new Set(existing.map(v=>v.id)), identities=new Set(existing.map(v=>studentIdentity(v.name,v.classroom))),codes=new Set(existing.filter(v=>v.code&&!v.code.startsWith('AUTO-')).map(v=>normalizeCode(v.code)));
 return rows.filter(v=>!ids.has(v.id)&&!identities.has(studentIdentity(v.name,v.classroom))&&(!v.code||v.code.startsWith('AUTO-')||!codes.has(normalizeCode(v.code))));
}
export function callStatus(packet,vid){
 if(packet.ballots.some(b=>b.id===vid))return 'voted';
 return (packet.calls||[]).find(c=>c.id===vid)?.status||'waiting';
}
export function queueRows(packet,{classroom='all',status='waiting'}={}){
 return packet.voters.filter(v=>(classroom==='all'||v.classroom===classroom)&&callStatus(packet,v.id)===status).sort((a,b)=>{
   if(status==='skipped'){
     const time=v=>{const raw=(packet.calls||[]).find(c=>c.id===v.id)?.updatedAt;return raw?.toMillis?.()||new Date(raw||0).getTime()||0;};
     const byTime=time(a)-time(b);if(byTime)return byTime;
   }
   return (a.queueNo??Number.MAX_SAFE_INTEGER)-(b.queueNo??Number.MAX_SAFE_INTEGER)||a.classroom.localeCompare(b.classroom,'id',{numeric:true})||a.name.localeCompare(b.name,'id',{numeric:true})||a.id.localeCompare(b.id);
 });
}
export function ownCall(packet,stationId,uid){return (packet.calls||[]).find(c=>['called','active'].includes(c.status)&&c.stationId===stationId&&c.operatorUid===uid&&!packet.ballots.some(b=>b.id===c.id));}
export function spokenClass(classroom){
 const text=clean(classroom).toUpperCase().replace(/^KELAS\s+/,'');
 const parts=text.match(/^(XII|XI|X|IX|VIII|VII|VI|V|IV|III|II|I|12|11|10|[1-9])[\s.\-]*(.*)$/);
 if(!parts)return text;
 const words={I:'satu',II:'dua',III:'tiga',IV:'empat',V:'lima',VI:'enam',VII:'tujuh',VIII:'delapan',IX:'sembilan',X:'sepuluh',XI:'sebelas',XII:'dua belas',1:'satu',2:'dua',3:'tiga',4:'empat',5:'lima',6:'enam',7:'tujuh',8:'delapan',9:'sembilan',10:'sepuluh',11:'sebelas',12:'dua belas'};
 return [words[parts[1]],parts[2]].filter(Boolean).join(' ');
}
export function callText(voter,unit,schoolName){return `Perhatian kepada ananda ${voter.name}, kelas ${spokenClass(voter.classroom)}. Silakan menuju bilik pemilihan OSIS ${unit.toUpperCase()} ${schoolName}.` ;}
export function queueStats(packet){const states={waiting:0,called:0,active:0,skipped:0,voted:0};for(const v of packet.voters)states[callStatus(packet,v.id)]++;return {...states,participation:percentage(states.voted,packet.voters.length)};}
export function assertOwnedCall(call,stationId,uid,statuses=['called','active']){
 if(!call||!statuses.includes(call.status)||call.stationId!==stationId||call.operatorUid!==uid)throw new Error('Panggilan/sesi tidak aktif di bilik ini. Panggil siswa dan verifikasi kehadirannya terlebih dahulu.');
}
