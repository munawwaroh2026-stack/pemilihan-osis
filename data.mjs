import {demoSeed, checkVote, uniqueId, UNITS} from './core.mjs';
import {newStudentsOnly,assertOwnedCall} from './students.mjs';

export const DEMO_STORAGE_KEY='almiftah-osis-demo-v1';
const clone=v=>structuredClone(v);
const stripId=({id,...rest})=>rest;

export class DemoStore {
  constructor(storage=globalThis.localStorage) {
    this.mode='demo';this.uid='DEMO';this.role={role:'admin',units:UNITS};this.storage=storage;this.listeners=new Set();
    this.memory=null;
    this.storageListener=e=>{if(e.key===DEMO_STORAGE_KEY)this.emit();};
    globalThis.addEventListener?.('storage',this.storageListener);
    if(!this.read(false))this.write(demoSeed());
  }
  read(fallback=true){
    try {const data=JSON.parse(this.storage?.getItem(DEMO_STORAGE_KEY)||'null');return data?.schools && data?.elections ? data : fallback?this.memory:null;}
    catch{return fallback?this.memory:null;}
  }
  write(data){
    const text=JSON.stringify(data);
    if(this.storage)this.storage.setItem(DEMO_STORAGE_KEY,text);
    this.memory=clone(data);this.emit();
  }
  emit(){for(const fn of this.listeners)fn();}
  subscribe(unit, selectedId, callback){
    const fn=()=>{
      const data=this.read();const school=data?.schools[unit]||null;
      const id=selectedId && school?.history.includes(selectedId)?selectedId:school?.currentElectionId;
      callback({calls:[],...clone(data?.elections[id]||{election:null,candidates:[],voters:[],ballots:[]}),school:clone(school),ready:true});
    };
    this.listeners.add(fn);fn();return ()=>this.listeners.delete(fn);
  }
  async createElection(unit,period){
    const data=this.read(), id=uniqueId(unit+'-');
    const previous=data.elections[data.schools[unit]?.currentElectionId];
    if(previous?.election.status==='open')throw new Error('Tutup periode aktif terlebih dahulu sebelum membuat periode baru.');
    data.elections[id]={election:{id,unit,title:`Pemilihan Ketua OSIS ${unit.toUpperCase()}`,period,status:'draft',createdAt:new Date().toISOString(),openedAt:null,closedAt:null},candidates:[],voters:[],ballots:[],calls:[],stations:{}};
    const history=[...(data.schools[unit]?.history||[]),id];
    data.schools[unit]={unit,currentElectionId:id,history};this.write(data);return id;
  }
  assertDraft(record){if(record?.election.status!=='draft')throw new Error('Data hanya dapat diubah pada tahap Persiapan. Buat periode baru untuk pemilihan berikutnya.');}
  async saveCandidate(eid,candidate){const data=this.read(),r=data.elections[eid];this.assertDraft(r);const i=r.candidates.findIndex(c=>c.id===candidate.id);if(i<0)r.candidates.push(candidate);else r.candidates[i]=candidate;this.write(data);}
  async deleteCandidate(eid,id){const data=this.read(),r=data.elections[eid];this.assertDraft(r);r.candidates=r.candidates.filter(c=>c.id!==id);this.write(data);}
  async addVoters(eid,rows){const data=this.read(),r=data.elections[eid];this.assertDraft(r);const fresh=newStudentsOnly(rows,r.voters);let order=Math.max(r.voters.length,...r.voters.map(v=>v.queueNo||0));r.voters.push(...fresh.map(v=>({...v,queueNo:++order})));this.write(data);return fresh.length;}
  async deleteVoter(eid,id){const data=this.read(),r=data.elections[eid];this.assertDraft(r);r.voters=r.voters.filter(v=>v.id!==id);this.write(data);}
  async setStatus(eid,status){const data=this.read(),r=data.elections[eid];if(status==='draft')throw new Error('Periode yang berjalan tidak dapat dikembalikan ke persiapan.');r.election.status=status;r.election[status==='open'?'openedAt':'closedAt']=new Date().toISOString();this.write(data);}
  async changeQueue(eid,fn){
    const commit=()=>{const data=this.read(),r=data.elections[eid];if(data.schools[r?.election.unit]?.currentElectionId!==eid)throw new Error('Periode ini sudah menjadi arsip.');if(r?.election.status!=='open')throw new Error('Pemilihan belum dibuka atau sudah ditutup.');r.calls??=[];r.stations??={};const result=fn(r);this.write(data);return result;};
    if(globalThis.navigator?.locks) return navigator.locks.request(DEMO_STORAGE_KEY,async()=>commit());
    return commit();
  }
  async callVoter(eid,studentId,stationId){return this.changeQueue(eid,r=>{
    const v=r.voters.find(v=>v.id===studentId);if(!v||r.ballots.some(b=>b.id===studentId))throw new Error('Hak pilih siswa tidak tersedia.');
    const found=r.calls.find(c=>c.id===studentId);
    if(found&&['called','active'].includes(found.status)){const error=new Error('Siswa sedang dipanggil atau memilih di bilik lain.');error.code='queue/busy';throw error;}
    const previousId=r.stations[stationId]?.voterId,previous=r.calls.find(c=>c.id===previousId);
    const now=new Date().toISOString();
    if(previous&&['called','active'].includes(previous.status)&&previous.stationId===stationId&&previous.operatorUid==='DEMO'){previous.status='skipped';previous.updatedAt=now;}
    const call={id:studentId,status:'called',stationId,operatorUid:'DEMO',calledAt:now,updatedAt:now};
    if(found)Object.assign(found,call);else r.calls.push(call);
    r.stations[stationId]={operatorUid:'DEMO',voterId:studentId,updatedAt:now};return v;
  });}
  async activateVoter(eid,studentId,stationId){return this.changeQueue(eid,r=>{
    const call=r.calls.find(c=>c.id===studentId);assertOwnedCall(call,stationId,'DEMO',['called','active']);
    if(r.stations[stationId]?.voterId!==studentId||r.ballots.some(b=>b.id===studentId))throw new Error('Sesi pemilih sudah berubah.');
    call.status='active';call.updatedAt=new Date().toISOString();
  });}
  async skipVoter(eid,studentId,stationId){return this.changeQueue(eid,r=>{
    const call=r.calls.find(c=>c.id===studentId);assertOwnedCall(call,stationId,'DEMO');
    if(r.ballots.some(b=>b.id===studentId))throw new Error('Siswa sudah memilih; hak pilih tetap terkunci.');
    call.status='skipped';call.updatedAt=new Date().toISOString();r.stations[stationId]={operatorUid:'DEMO',voterId:'',updatedAt:call.updatedAt};
  });}
  async releaseCall(eid,studentId){return this.changeQueue(eid,r=>{
    const call=r.calls.find(c=>c.id===studentId);if(!call||!['called','active'].includes(call.status)||r.ballots.some(b=>b.id===studentId))throw new Error('Panggilan ini sudah tidak aktif.');call.status='skipped';call.updatedAt=new Date().toISOString();
  });}
  async castVote(eid,studentId,candidateId,stationId){return this.changeQueue(eid,r=>{
    checkVote(r,studentId,candidateId);const call=r.calls.find(c=>c.id===studentId);assertOwnedCall(call,stationId,'DEMO',['active']);
    if(r.stations[stationId]?.voterId!==studentId)throw new Error('Sesi pemilih sudah berubah.');
    const now=new Date().toISOString();r.ballots.push({id:studentId,candidateId,castAt:now,operatorUid:'DEMO',stationId});call.status='voted';call.updatedAt=now;r.stations[stationId]={operatorUid:'DEMO',voterId:'',updatedAt:now};
  });}
  async signOut(){this.destroy();}
  destroy(){globalThis.removeEventListener?.('storage',this.storageListener);this.listeners.clear();}
}

export async function createFirebaseStore(config) {
  // Versi dipatok agar pemasangan GitHub dapat diulang. Hanya produk yang dipakai dimuat.
  const base='https://www.gstatic.com/firebasejs/13.0.0/';
  const [appSDK,authSDK,dbSDK]=await Promise.all([
    import(base+'firebase-app.js'), import(base+'firebase-auth.js'), import(base+'firebase-firestore.js')
  ]);
  const app=appSDK.getApps().length ? appSDK.getApp() : appSDK.initializeApp(config);
  const auth=authSDK.getAuth(app), db=dbSDK.getFirestore(app);
  await authSDK.setPersistence(auth,authSDK.browserSessionPersistence);
  await auth.authStateReady();
  const {doc,collection,getDoc,setDoc,deleteDoc,onSnapshot,runTransaction,writeBatch,serverTimestamp}=dbSDK;
  const ref=(...path)=>doc(db,...path);
  const backend={mode:'firebase',auth,db,uid:null,role:null,
    async signIn(email,password){
      const result=await authSDK.signInWithEmailAndPassword(auth,email,password);
      return this.authorize(result.user);
    },
    async authorize(user=auth.currentUser){
      if(!user)return false;
      try{
        const snap=await getDoc(ref('staff',user.uid));const role=snap.data();
        if(!snap.exists() || !['admin','operator'].includes(role.role) || !Array.isArray(role.units) || !role.units.some(u=>UNITS.includes(u))) {
          throw new Error('Akun belum terdaftar sebagai panitia. Tambahkan dokumen staff sesuai panduan.');
        }
        this.uid=user.uid;this.role={...role,units:role.units.filter(u=>UNITS.includes(u))};return true;
      }catch(error){await authSDK.signOut(auth);throw error;}
    },
    async signOut(){await authSDK.signOut(auth);this.uid=null;this.role=null;},
    async reauthenticate(password){
      if(!auth.currentUser?.email)throw new Error('Silakan masuk ulang.');
      const credential=authSDK.EmailAuthProvider.credential(auth.currentUser.email,password);
      await authSDK.reauthenticateWithCredential(auth.currentUser,credential);
    },
    subscribe(unit,selectedId,callback,onError){
      let inner=[], activeId, stopped=false;
      let packet={school:null,election:null,candidates:[],voters:[],ballots:[],calls:[],ready:false};
      const received=new Map();
      const emit=()=>{if(!stopped)callback({...packet,ready:received.size===5 && [...received.values()].every(Boolean)});};
      const error=e=>{received.clear();emit();onError?.(e);};
      const outer=onSnapshot(ref('schools',unit),{includeMetadataChanges:true},snap=>{
        packet.school=snap.exists()?snap.data():null;
        const id=selectedId && packet.school?.history?.includes(selectedId)?selectedId:packet.school?.currentElectionId;
        if(!id){inner.forEach(fn=>fn());inner=[];activeId=null;packet={school:packet.school,election:null,candidates:[],voters:[],ballots:[],calls:[],ready:!snap.metadata.fromCache};callback({...packet});return;}
        if(id===activeId){emit();return;}
        activeId=id;inner.forEach(fn=>fn());received.clear();
        packet={school:packet.school,election:null,candidates:[],voters:[],ballots:[],calls:[],ready:false};emit();
        inner.push(onSnapshot(ref('elections',id),{includeMetadataChanges:true},s=>{packet.election=s.exists()?{id:s.id,...s.data()}:null;received.set('election',!s.metadata.fromCache);emit();},error));
        for(const key of ['candidates','voters','ballots','calls'])inner.push(onSnapshot(collection(db,'elections',id,key),{includeMetadataChanges:true},s=>{
          packet[key]=s.docs.map(d=>({id:d.id,...d.data()}));received.set(key,!s.metadata.fromCache);emit();
        },error));
      },error);
      return ()=>{stopped=true;outer();inner.forEach(fn=>fn());};
    },
    async createElection(unit,period){
      const id=uniqueId(unit+'-');
      await runTransaction(db,async tx=>{
        const schoolRef=ref('schools',unit), school=await tx.get(schoolRef);
        const history=school.data()?.history||[];
        const previousId=school.data()?.currentElectionId;
        if(previousId){const previous=await tx.get(ref('elections',previousId));if(previous.data()?.status==='open')throw new Error('Tutup periode aktif terlebih dahulu sebelum membuat periode baru.');}
        if(history.length>=100)throw new Error('Batas 100 periode tercapai. Hubungi pengelola data.');
        tx.set(ref('elections',id),{unit,title:`Pemilihan Ketua OSIS ${unit.toUpperCase()}`,period,status:'draft',createdAt:serverTimestamp(),openedAt:null,closedAt:null});
        tx.set(schoolRef,{unit,currentElectionId:id,history:[...history,id]});
      });return id;
    },
    async saveCandidate(eid,candidate){await setDoc(ref('elections',eid,'candidates',candidate.id),stripId(candidate));},
    async deleteCandidate(eid,id){await deleteDoc(ref('elections',eid,'candidates',id));},
    async addVoters(eid,rows,onProgress){
      let done=0;
      for(let i=0;i<rows.length;i+=100){
        const chunk=rows.slice(i,i+100),batch=writeBatch(db);
        for(const v of chunk)batch.set(ref('elections',eid,'voters',v.id),stripId(v));
        await batch.commit();done+=chunk.length;onProgress?.(done,rows.length);
      }return done;
    },
    async deleteVoter(eid,id){await deleteDoc(ref('elections',eid,'voters',id));},
    async setStatus(eid,status){
      if(!['open','closed'].includes(status))throw new Error('Status tidak valid.');
      await runTransaction(db,async tx=>{
        const er=ref('elections',eid),snap=await tx.get(er);
        if(!snap.exists())throw new Error('Periode tidak ditemukan.');
        const prev=snap.data().status;
        if((status==='open' && !['draft','closed'].includes(prev)) || (status==='closed' && prev!=='open'))throw new Error('Status berubah di perangkat lain. Muat ulang data.');
        tx.update(er,{status,[status==='open'?'openedAt':'closedAt']:serverTimestamp()});
      });
    },
    async callVoter(eid,studentId,stationId){
      await runTransaction(db,async tx=>{
        const er=ref('elections',eid),vr=ref('elections',eid,'voters',studentId),br=ref('elections',eid,'ballots',studentId),cr=ref('elections',eid,'calls',studentId),sr=ref('elections',eid,'stations',stationId);
        const election=await tx.get(er),voter=await tx.get(vr),ballot=await tx.get(br),call=await tx.get(cr),station=await tx.get(sr);
        if(election.data()?.status!=='open'||!voter.exists()||ballot.exists())throw new Error('Pemilihan atau hak pilih siswa sudah berubah.');
        if(station.exists()&&station.data().operatorUid!==auth.currentUser.uid)throw new Error('Bilik terikat pada akun lain. Keluar lalu masuk kembali.');
        if(call.exists()&&['called','active'].includes(call.data().status)){const error=new Error('Siswa sedang dipanggil/menentukan pilihan di bilik lain.');error.code='queue/busy';throw error;}
        const prevId=station.data()?.voterId;let previous,previousBallot,previousRef;
        if(prevId&&prevId!==studentId){previousRef=ref('elections',eid,'calls',prevId);previous=await tx.get(previousRef);previousBallot=await tx.get(ref('elections',eid,'ballots',prevId));}
        if(previous?.exists()&&['called','active'].includes(previous.data().status)&&!previousBallot.exists()&&previous.data().stationId===stationId&&previous.data().operatorUid===auth.currentUser.uid){
          assertOwnedCall(previous.data(),stationId,auth.currentUser.uid);tx.update(previousRef,{status:'skipped',updatedAt:serverTimestamp()});
        }
        tx.set(cr,{status:'called',stationId,operatorUid:auth.currentUser.uid,calledAt:serverTimestamp(),updatedAt:serverTimestamp()});
        tx.set(sr,{operatorUid:auth.currentUser.uid,voterId:studentId,updatedAt:serverTimestamp()});
      });
    },
    async activateVoter(eid,studentId,stationId){
      await runTransaction(db,async tx=>{
        const er=ref('elections',eid),cr=ref('elections',eid,'calls',studentId),br=ref('elections',eid,'ballots',studentId),sr=ref('elections',eid,'stations',stationId);
        const election=await tx.get(er),call=await tx.get(cr),ballot=await tx.get(br),station=await tx.get(sr);
        if(election.data()?.status!=='open'||ballot.exists()||station.data()?.voterId!==studentId)throw new Error('Pemilihan atau hak pilih siswa sudah berubah.');
        assertOwnedCall(call.data(),stationId,auth.currentUser.uid,['called','active']);if(call.data().status!=='active')tx.update(cr,{status:'active',updatedAt:serverTimestamp()});
      });
    },
    async skipVoter(eid,studentId,stationId){
      await runTransaction(db,async tx=>{
        const cr=ref('elections',eid,'calls',studentId),br=ref('elections',eid,'ballots',studentId),sr=ref('elections',eid,'stations',stationId);
        const call=await tx.get(cr),ballot=await tx.get(br),station=await tx.get(sr);assertOwnedCall(call.data(),stationId,auth.currentUser.uid);
        if(ballot.exists()||station.data()?.voterId!==studentId)throw new Error('Hak pilih atau sesi sudah berubah.');
        tx.update(cr,{status:'skipped',updatedAt:serverTimestamp()});tx.update(sr,{voterId:'',updatedAt:serverTimestamp()});
      });
    },
    async releaseCall(eid,studentId){await dbSDK.updateDoc(ref('elections',eid,'calls',studentId),{status:'skipped',updatedAt:serverTimestamp()});},
    async castVote(eid,studentId,candidateId,stationId){
      // Transaksi server: tidak mengantre suara offline. Dokumen suara satu per DPT.
      await runTransaction(db,async tx=>{
        const electionRef=ref('elections',eid), voterRef=ref('elections',eid,'voters',studentId),
          candidateRef=ref('elections',eid,'candidates',candidateId), ballotRef=ref('elections',eid,'ballots',studentId),callRef=ref('elections',eid,'calls',studentId),stationRef=ref('elections',eid,'stations',stationId);
        const election=await tx.get(electionRef), voter=await tx.get(voterRef), candidate=await tx.get(candidateRef),ballot=await tx.get(ballotRef),call=await tx.get(callRef),station=await tx.get(stationRef);
        checkVote({election:election.data(),voters:voter.exists()?[{id:voter.id}]:[],candidates:candidate.exists()?[{id:candidate.id}]:[],ballots:ballot.exists()?[{id:ballot.id}]:[]},studentId,candidateId);
        assertOwnedCall(call.data(),stationId,auth.currentUser.uid,['active']);if(station.data()?.voterId!==studentId)throw new Error('Sesi pemilih sudah berubah.');
        tx.set(ballotRef,{candidateId,castAt:serverTimestamp(),operatorUid:auth.currentUser.uid,stationId});
        tx.update(callRef,{status:'voted',updatedAt:serverTimestamp()});tx.update(stationRef,{voterId:'',updatedAt:serverTimestamp()});
      });
    }
  };
  return backend;
}

export function readableError(error) {
  const messages={
    'auth/invalid-credential':'Email atau password tidak sesuai.',
    'auth/wrong-password':'Password tidak sesuai.',
    'auth/user-not-found':'Email atau password tidak sesuai.',
    'auth/invalid-email':'Periksa kembali format email.',
    'auth/too-many-requests':'Terlalu banyak percobaan masuk. Coba lagi beberapa saat.',
    'auth/network-request-failed':'Tidak dapat menghubungi Firebase. Periksa internet.',
    'auth/operation-not-allowed':'Aktifkan Email/Password di Firebase Authentication.',
    'auth/unauthorized-domain':'Tambahkan domain GitHub Pages Anda di Authorized domains Firebase Auth.',
    'auth/invalid-api-key':'API key Firebase belum benar. Periksa app-config.js.',
    'permission-denied':'Akses ditolak. Periksa Firestore Rules, dokumen staff, dan tahap pemilihan.',
    'unavailable':'Koneksi Firebase terputus. Suara belum dikonfirmasi; periksa status pemilih sebelum mencoba ulang.',
    'resource-exhausted':'Kuota Firebase terpakai. Hentikan pemilihan dan periksa Usage di Firebase Console.'
  };
  return messages[error?.code] || error?.message || 'Terjadi kesalahan. Coba kembali.';
}
