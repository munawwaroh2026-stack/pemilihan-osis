// Pembaca tabel .xlsx untuk kolom nama/kelas, menggunakan JSZip dan SAX lokal.
// Tidak mengeksekusi rumus atau macro. Berkas .xls lama perlu Save As .xlsx/CSV.
const MAX_EXPANDED=20*1024*1024;
function preflight(buffer){
 const b=new DataView(buffer);if(b.byteLength<22)throw new Error('File Excel tidak valid.');
 let end=-1;for(let i=b.byteLength-22;i>=Math.max(0,b.byteLength-65557);i--)if(b.getUint32(i,true)===0x06054b50&&i+22+b.getUint16(i+20,true)===b.byteLength){end=i;break;}
 if(end<0)throw new Error('Gunakan file .xlsx yang tidak diproteksi password.');
 const count=b.getUint16(end+10,true),offset=b.getUint32(end+16,true);let p=offset,total=0;
 if(count>1000||offset>=end)throw new Error('Struktur Excel terlalu besar atau tidak didukung.');
 for(let i=0;i<count;i++){
   if(p+46>end||b.getUint32(p,true)!==0x02014b50)throw new Error('Struktur file Excel rusak.');
   if(b.getUint16(p+8,true)&1)throw new Error('Hapus password proteksi file Excel sebelum mengunggah.');
   const size=b.getUint32(p+24,true);total+=size;
   if(size>10*1024*1024||total>MAX_EXPANDED)throw new Error('Isi file Excel terlalu besar. Gunakan tabel nama/kelas yang sederhana.');
   p+=46+b.getUint16(p+28,true)+b.getUint16(p+30,true)+b.getUint16(p+32,true);
 }
}
function parseXML(text,handlers){
 if(/<!DOCTYPE|<!ENTITY/i.test(text))throw new Error('File Excel memiliki XML yang tidak didukung.');
 const parser=globalThis.sax?.parser(true,{trim:false,normalize:false});if(!parser)throw new Error('Komponen pembaca Excel belum termuat. Muat ulang halaman.');
 parser.onopentag=node=>handlers.open?.(node.name.split(':').pop().toLowerCase(),Object.fromEntries(Object.entries(node.attributes).map(([k,v])=>[k.toLowerCase(),String(v)])));
 parser.ontext=text=>handlers.text?.(text);parser.oncdata=parser.ontext;
 parser.onclosetag=name=>handlers.close?.(name.split(':').pop().toLowerCase());
 try{parser.write(text).close();}catch(error){throw new Error('Data XML Excel tidak valid: '+error.message.split('\n')[0]);}
}
function columnIndex(ref){const letters=/^([A-Z]+)/i.exec(ref)?.[1];if(!letters)return -1;let n=0;for(const c of letters.toUpperCase())n=n*26+c.charCodeAt(0)-64;return n-1;}
export function sheetMatrix(xml,shared=[]){
 const rows=[];let row=null,cell=null,tag='',tDepth=0;
 parseXML(xml,{
   open(name,a){tag=name;if(name==='row')row=[];if(name==='c')cell={ref:a.r||'',type:a.t||'n',value:'',inline:'',formula:false};if(name==='f'&&cell)cell.formula=true;if(name==='t')tDepth++;},
   text(text){if(!cell)return;if(tag==='v')cell.value+=text;else if(tDepth&&cell.type==='inlineStr')cell.inline+=text;},
   close(name){
     if(name==='t')tDepth--;
     if(name==='c'&&cell&&row){
       const col=cell.ref?columnIndex(cell.ref):row.length;
       if(col>63)throw new Error('Gunakan maksimal 64 kolom pada lembar impor.');
       if(cell.formula)throw new Error(`Sel ${cell.ref} berisi rumus. Salin dan tempel sebagai nilai terlebih dahulu.`);
       if(cell.type==='e')throw new Error(`Sel ${cell.ref} memiliki kesalahan Excel.`);
       let value=cell.type==='s'?shared[Number(cell.value)]:cell.type==='inlineStr'?cell.inline:cell.value;
       if(cell.type==='s'&&(value===undefined||!/^\d+$/.test(cell.value)))throw new Error('Referensi teks Excel tidak valid.');
       if(col>=0)row[col]=String(value??'');cell=null;
     }
     if(name==='row'&&row){if(row.some(v=>String(v??'').trim()))rows.push(Array.from({length:row.length},(_,i)=>row[i]??''));row=null;if(rows.length>5001)throw new Error('Maksimal 5.000 siswa per file.');}
     tag='';
   }
 });return rows;
}
export async function readExcel(buffer){
 if(buffer.byteLength>5*1024*1024)throw new Error('Maksimal ukuran Excel 5 MB.');preflight(buffer);
 if(!globalThis.JSZip)throw new Error('Komponen Excel belum termuat. Pastikan folder assets/vendor ikut diunggah.');
 const zip=await globalThis.JSZip.loadAsync(buffer,{checkCRC32:true});
 const read=async path=>{const entry=zip.file(path);if(!entry)throw new Error('Struktur Excel tidak lengkap: '+path);return entry.async('string');};
 let firstSheet;parseXML(await read('xl/workbook.xml'),{open(name,a){if(name==='sheet'&&!firstSheet){const relation=Object.keys(a).find(k=>k.endsWith(':id'));firstSheet={name:a.name||'Lembar pertama',rid:a[relation]};}}});
 if(!firstSheet)throw new Error('Excel belum memiliki lembar kerja.');
 let target;parseXML(await read('xl/_rels/workbook.xml.rels'),{open(name,a){if(name==='relationship'&&a.id===firstSheet.rid&&a.targetmode!=='External')target=a.target;}});
 if(!target)throw new Error('Lembar pertama bukan tabel yang didukung. Letakkan data nama/kelas di lembar pertama.');
 const segments=(target.startsWith('/')?target.slice(1):'xl/'+target).split('/'),parts=[];
 for(const part of segments){if(!part||part==='.')continue;if(part==='..')parts.pop();else parts.push(part);}
 const path=parts.join('/');if(!path.startsWith('xl/worksheets/'))throw new Error('Lembar pertama bukan lembar data biasa.');
 const shared=[];
 if(zip.file('xl/sharedStrings.xml')){
   let current='',inside=false,textDepth=0,phonetic=0;
   parseXML(await read('xl/sharedStrings.xml'),{open(name){if(name==='si'){current='';inside=true;}if(name==='rph')phonetic++;if(name==='t')textDepth++;},text(text){if(inside&&textDepth&&!phonetic)current+=text;},close(name){if(name==='t')textDepth--;if(name==='rph')phonetic--;if(name==='si'){shared.push(current);inside=false;}}});
 }
 return {rows:sheetMatrix(await read(path),shared),sheetName:firstSheet.name};
}
