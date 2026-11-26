'use strict';
const $=id=>document.getElementById(id);
const NOT0=["Etiket yok","Etiket silik","Hasarlı","Başka odadan geldi","Sistemde kaydı yok","Kullanılmıyor","Arızalı","Zimmet Yap","Bakılacak","Diğer"]; // Excel G sütunu listesi
const TUM='TÜM YERLER';
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const p2=n=>String(n).padStart(2,'0');
const norm=s=>String(s||'').replace(/[\u200b-\u200d\ufeff]/g,'').replace(/\u00a0/g,' ').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLocaleUpperCase('tr').replace(/İ/g,'I').replace(/\s+/g,' '); // aksan ve İ/I farkları yok sayılır (kaynakta İ/I gibi çift yazımlar olabilir)
const TUMN='TUM YERLER';
const kurum=s=>norm(s).startsWith('YÜKSEKOKUL SEKRETERLİĞİ');
const LS={get(k,d){try{const v=localStorage.getItem(k);return v===null?d:JSON.parse(v)}catch(e){return d}},set(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}}};
const tarihTR=d=>p2(d.getDate())+'.'+p2(d.getMonth()+1)+'.'+d.getFullYear();
const saatTR=d=>p2(d.getHours())+':'+p2(d.getMinutes())+':'+p2(d.getSeconds());
const damga=d=>d.getFullYear()+p2(d.getMonth()+1)+p2(d.getDate())+'-'+p2(d.getHours())+p2(d.getMinutes())+p2(d.getSeconds());
const isoTar=k=>{const a=k.tarih.split('.');return a[2]+'-'+a[1]+'-'+a[0]};
const temiz=s=>String(s||'').replace(/[\t\r\n]+/g,' ').trim();
function rast(n){const a='abcdefghijklmnopqrstuvwxyz0123456789';let s='';const b=new Uint8Array(n);crypto.getRandomValues(b);for(const x of b)s+=a[x%a.length];return s}
// Excel BarkodDuzelt/Anahtar: '-' içeriyorsa Bar_Kod (olduğu gibi), yoksa yalnız rakamlar
// Excel BarkodDuzelt ile aynı: . / boşluk atılınca yalnız rakam kalıyorsa rakam dizisi, değilse olduğu gibi (Bar_Kod: 255.2.1.1.1.1-215)
const duzelt=k=>{const t=String(k||'').replace(/[\r\n\t\u00a0]/g,'').trim();if(!t)return '';const s=t.replace(/[.\/ ]/g,'');return /^[0-9]+$/.test(s)?s:t};
function normAd(t){return String(t||'').replace(/İ/g,'i').replace(/I/g,'i').toLowerCase().replace(/ı/g,'i').replace(/ş/g,'s').replace(/ğ/g,'g').replace(/ü/g,'u').replace(/ö/g,'o').replace(/ç/g,'c').replace(/[^a-z0-9]/g,'')}
// ---------- IndexedDB ----------
let db;
const dbAc=()=>new Promise((res,rej)=>{const r=indexedDB.open('sayimDB',2);r.onupgradeneeded=()=>{const d=r.result;if(!d.objectStoreNames.contains('kuyruk'))d.createObjectStore('kuyruk',{keyPath:'kimlik'});if(!d.objectStoreNames.contains('kv'))d.createObjectStore('kv')};r.onsuccess=()=>{db=r.result;res()};r.onerror=()=>rej(r.error)});
const st=(n,m)=>db.transaction(n,m).objectStore(n);
const dbHepsi=()=>new Promise(r=>{const q=st('kuyruk','readonly').getAll();q.onsuccess=()=>r(q.result||[]);q.onerror=()=>r([])});
const dbYaz=k=>new Promise((res,rej)=>{const q=st('kuyruk','readwrite').put(k);q.onsuccess=()=>res();q.onerror=()=>rej(q.error)});
const dbSil=id=>new Promise(r=>{const q=st('kuyruk','readwrite').delete(id);q.onsuccess=()=>r();q.onerror=()=>r()});
const kvG=k=>new Promise(r=>{const q=st('kv','readonly').get(k);q.onsuccess=()=>r(q.result);q.onerror=()=>r(undefined)});
const kvS=(k,v)=>new Promise(r=>{const q=st('kv','readwrite').put(v,k);q.onsuccess=()=>r();q.onerror=()=>r()});
// ---------- veri ----------
let kayit=[];            // tüm kayıtlar (zaman sıralı)
let yerler=[],malzemeler=[],notlar=NOT0.slice(),bek=[],anah=new Map(),barm=new Map(),yerSet=new Map();
function indeksle(){anah=new Map();barm=new Map();yerSet=new Map();
 for(const r of bek){anah.set(r[0],r);if(r[2]){(barm.get(r[2])||barm.set(r[2],[]).get(r[2])).push(r)}
  const ns=new Set();for(const y of [r[9],kurum(r[11])?'':r[11]]){if(y)ns.add(norm(y))}for(const n of ns){if(!n)continue;(yerSet.get(n)||yerSet.set(n,[]).get(n)).push(r)}}}
function ayr(l){return String(l||'').replace(/^﻿/,'').replace(/\r/g,'').split('\n').filter(x=>x.trim()!=='')}
function listeKur(y,n,b,m){if(y!=null){yerler=ayr(y);if(yerler.length){const g=new Set(yerler.map(norm));LS.set('yerGecmis',LS.get('yerGecmis',[]).filter(x=>g.has(norm(x))))}}if(m!=null)malzemeler=ayr(m);if(n!=null){const x=ayr(n).flatMap(s=>s.split(';')).map(s=>s.trim()).filter(Boolean);if(x.length)notlar=x}
 if(b!=null){bek=ayr(b).map(l=>l.split('\t')).filter(r=>r.length>=10&&r[0]!=='Anahtar');indeksle()}
 
 $('lb').textContent='Yer: '+yerler.length+' · Not: '+notlar.length+' · Malzeme adı: '+malzemeler.length+' · Sistem kaydı: '+bek.length+(bek.length?'':'  (liste yok: Excel\'de Ayarlar > Mobil Listeleri Yaz, sonra buradan alın)')}
// ---------- durum (Excel Durum formülüyle aynı sıra) ----------
function durum(sicil,ad,yer,gor){
 const s=duzelt(sicil);
 if(!s&&!ad)return{t:'—',c:'no',r:null};
 if(!yer)return{t:'YER SEÇİN',c:'yer',r:null};
 if(!s)return{t:'FAZLA – sicilsiz (elle)',c:'faz',r:null};
 if(gor&&gor.has(s))return{t:'MÜKERRER',c:'muk',r:anah.get(s)||(barm.get(s)||[])[0]||null};
 const r=anah.get(s);
 if(!r){if(barm.has(s))return{t:'BARKOD (sicil değil)',c:'yer',r:barm.get(s)[0],n:barm.get(s).length};
  return bek.length?{t:'FAZLA – raporda yok',c:'faz',r:null}:{t:'(sistem listesi yok)',c:'no',r:null}}
 if(norm(yer)===TUMN)return{t:'YERİNDE',c:'ok',r};
 const y=norm(yer);
 if(norm(r[9])===y||(!kurum(r[11])&&norm(r[11])===y))return{t:'YERİNDE',c:'ok',r};
 return{t:'FAZLA – başka yerde kayıtlı',c:'faz',r};
}
function hesaplaTum(){const gor=new Set(),o=new Map();for(const k of kayit){const d=durum(k.sicil,k.malzeme,k.yer,gor);const s=duzelt(k.sicil);if(s)gor.add(s);o.set(k.kimlik,d)}return o}
// ---------- bilgi kartı ----------
function kartHtml(r,extra){if(!r)return'';
 const a=[['Açıklama',r[3]],['Ek Özellik',r[5]],['Kayıtlı Yer / Kişi',r[9]],['Kime verildi / kimde',r[11]],['Zimmet Tarihi',r[13]],['Fiş No',r[12]],['Fiş Son Durum',r[8]],['Sicil No',r[1]],['Bar_Kod',r[2]]].concat(extra||[]);
 return '<div class="bas">'+esc((r[4]||'')+' '+(r[6]||'')+' '+(r[7]||'')).trim()+'</div>'+a.filter(x=>x[1]&&String(x[1]).trim()).map(x=>'<b>'+x[0]+'</b><span'+((x[0]==='Sicil No'||x[0]==='Bar_Kod')?' data-cp="'+esc(x[1])+'" style="cursor:pointer;text-decoration:underline dotted"':'')+'>'+(x[2]?x[1]:esc(x[1]))+'</span>').join('')}
function zamanPil(z,d){return '<span class="zp"><span>'+esc(z)+'</span><em class="pl '+d.c+'">'+esc(d.t)+(d.n?' ('+d.n+'×)':'')+'</em></span>'}
function yerAdi(){return $('yer').value.trim()}
function yerU(){return $('uY').value.trim()||yerAdi()}
function yerK(){return $('kyY').value.trim()||yerAdi()}
let dzId=null,notZorla=false,sonMS='',mt=null;
function medyaTetik(){const s=duzelt($('sicil').value).replace(/[^0-9]/g,'');if(!s||s===sonMS)return;sonMS=s;clearTimeout(mt);mt=setTimeout(()=>{if(duzelt($('sicil').value).replace(/[^0-9]/g,'')===s)medyaGoster(s,yerAdi())},500)}
function onizle(){const s0=duzelt($('sicil').value);const gor=new Set(kayit.filter(k=>k.kimlik!==dzId).map(k=>duzelt(k.sicil)).filter(Boolean));
 const d=durum($('sicil').value,$('ad').value,yerAdi(),gor);const n=new Date();
 $('ban').className='ban '+d.c;$('ban').innerHTML=esc(d.t)+(d.n?'<small>'+d.n+' kayıtta aynı barkod</small>':'');$('ban').classList.toggle('hide',!!d.r);
 const gizle=!notZorla&&!$('not').value&&!$('ad').value;
 $('notBlok').classList.toggle('hide',gizle);$('adBlok').classList.toggle('hide',gizle);$('notAcW').classList.toggle('hide',!gizle);
 $('dzKutu').innerHTML='';$('ekle').textContent='Kaydet (Enter)';
 $('kart').innerHTML=kartHtml(d.r,d.r?[['Tarih / Saat',zamanPil(tarihTR(n)+' '+saatTR(n),d),1]]:[]);medyaTetik();return d}
function duzeltBaslat(id){const k=kayit.find(x=>x.kimlik===id);if(!k)return;dzId=id;$('yer').value=k.yer;LS.set('sonYer',k.yer);$('sicil').value=k.sicil;$('not').value=k.not||'';$('ad').value=k.malzeme||'';notZorla=true;
 $('dzBan').classList.remove('hide');$('dzBan').innerHTML='✎ <b>Düzeltme modu:</b> '+esc(k.sicil||k.malzeme)+(k.durum==='gitti'?' kaydı Excel\'de de güncellenecek (düzeltme dosyası gönderilir). ':' kaydı güncellenecek. ')+'Değiştirip Kaydet\'e basın. <a id="dzVaz" class="lnk" style="float:none;display:inline-block;margin-left:8px">Vazgeç</a> <a id="dzIpt" class="lnk" style="float:none;display:inline-block;margin-left:8px;color:var(--er)">Kaydı iptal et</a>';hepsi();$('sicil').focus();window.scrollTo(0,0)}
function duzeltBitir(){dzId=null;notZorla=false;$('dzBan').classList.add('hide');$('sicil').value='';$('not').value='';$('ad').value='';hepsi()}
document.addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-dz]');if(b){duzeltBaslat(b.dataset.dz);return}if(e.target.closest&&e.target.closest('[data-kay]')){ekle();return}if(e.target.id==='dzVaz')duzeltBitir();if(e.target.id==='dzIpt')kayitIptal()});
const gecSn=()=>0;
// ---------- yerleri sırayla gezme (telefon) ----------
function yerListe(){return yerler.filter(x=>norm(x)!==TUMN)}
function yerAt(a){$('yer').value=a;LS.set('sonYer',a);hepsi()}
function yerGit(dir){const l=yerListe();if(!l.length)return;let i=l.findIndex(x=>norm(x)===norm(yerAdi()));i=i<0?(dir>0?0:l.length-1):(i+dir+l.length)%l.length;yerAt(l[i])}
function yerSayilmamis(){const l=yerListe();if(!l.length)return;let i=l.findIndex(x=>norm(x)===norm(yerAdi()));const say=x=>kayit.some(k=>norm(k.yer)===norm(x));
 for(let j=1;j<=l.length;j++){const a=l[(Math.max(i,-1)+j)%l.length];if(!say(a)&&(yerSet.get(norm(a))||[]).length){yerAt(a);return}}
 $('yPos').textContent='Sistemde kaydı olan tüm yerlerde okutma yapılmış.'}
function yPosGuncelle(){const l=yerListe();const i=l.findIndex(x=>norm(x)===norm(yerAdi()));$('yPos').textContent=i>=0?('Listede '+(i+1)+' / '+l.length):''}
// ---------- özetler ----------
function yerKume(y){const n=norm(y);if(!n)return null;if(n===TUMN)return bek;return yerSet.get(n)||[]}
function yerOzet(y){const ds=hesaplaTum(),n=norm(y),tum=n===TUMN;
 const kay=yerKume(y)||[];const ok=new Set();let okut=0,faz=0,muk=0;const fz=[];
 for(const k of kayit){if(!tum&&norm(k.yer)!==n)continue;okut++;const d=ds.get(k.kimlik);
  if(d.t==='YERİNDE')ok.add(duzelt(k.sicil));else if(d.t==='MÜKERRER')muk++;else if(d.c==='faz'||d.c==='yer'){faz++;fz.push([k,d])}}
 const nok=kay.filter(r=>!ok.has(r[0]));
 return{kay:kay.length,okut,yerinde:ok.size,nok,fz,faz,muk,ds}}
function cntHtml(o){return [['Kayıtlı',o.kay,'c1','kayitli'],['Okutulan',o.okut,'','okutulan'],['Yerinde',o.yerinde,'c2','yerinde'],['Noksan',o.nok.length,'c3','noksan'],['Fazla',o.faz,'c4','fazla'],['Mükerrer',o.muk,'c5','mukerrer']].map(x=>'<div class="'+x[2]+'" data-g="'+x[3]+'">'+x[0]+'<b>'+x[1]+'</b></div>').join('')}
function ozetCiz(){const y=yerAdi();
 if(!y){$('oz').innerHTML='';$('cub').style.width='0';return}
 const o=yerOzet(y);$('oz').innerHTML=cntHtml(o);$('cub').style.width=(o.kay?Math.round(100*o.yerinde/o.kay):0)+'%';return o}
function ozetSayfa(){const y=yerAdi();$('ozYer').textContent=y?('Yer / kişi: '+y):'Önce Sayım sekmesinden yer / kişi seçin (veya TÜM YERLER yazın).';
 if(!y){for(const i of['oz2','noksan','fazla','mukl','sayl'])$(i).innerHTML='';return}
 const o=yerOzet(y);const tum=norm(y)===TUMN,n=norm(y);$('oz2').innerHTML=cntHtml(o);$('cub2').style.width=(o.kay?Math.round(100*o.yerinde/o.kay):0)+'%';
 $('nokN').textContent='('+o.nok.length+')';$('fazN').textContent='('+o.faz+')';
 $('noksan').innerHTML=o.nok.slice(0,500).map(r=>'<div class="it nok" data-cp="'+esc(r[1])+'"><b>'+esc(r[1])+'</b> '+esc(r[4])+'<small>'+esc((r[6]||'')+' '+(r[7]||''))+' · '+esc(r[9])+(r[11]?' · '+esc(r[11]):'')+'</small></div>').join('')||'<div class="msg">Noksan yok.</div>';
 $('fazla').innerHTML=o.fz.slice(0,500).map(x=>'<div class="it kfaz" data-cp="'+esc(x[0].sicil)+'"><b>'+esc(x[0].sicil||x[0].malzeme)+'</b> '+esc(x[1].t)+'<small>'+esc(x[1].r?(x[1].r[4]+' · kayıtlı: '+x[1].r[9]):'')+(x[0].not?' · '+esc(x[0].not):'')+'</small></div>').join('')||'<div class="msg">Fazla yok.</div>';
 const mk=[],sy=[];for(const k of kayit){if(!tum&&norm(k.yer)!==n)continue;const d=o.ds.get(k.kimlik);if(d.t==='MÜKERRER')mk.push([k,d]);else if(d.t==='YERİNDE')sy.push([k,d])}
 $('mukN').textContent='('+mk.length+')';$('sayN').textContent='('+sy.length+')';
 $('mukl').innerHTML=mk.map(x=>'<div class="it kmuk"><b>'+esc(x[0].sicil)+'</b> MÜKERRER<small>'+esc(x[0].yer)+' · '+x[0].tarih+' '+x[0].saat+'</small></div>').join('')||'<div class="msg">Mükerrer yok.</div>';
 $('sayl').innerHTML=sy.map(x=>'<div class="it say"><b>'+esc(x[0].sicil)+'</b> '+esc(x[1].r?x[1].r[4]:'')+'<small>'+x[0].tarih+' '+x[0].saat+'</small></div>').join('')}
async function kopya(s){try{await navigator.clipboard.writeText(s)}catch(e){const a=document.createElement('textarea');a.value=s;document.body.appendChild(a);a.select();try{document.execCommand('copy')}catch(_){}a.remove()}}
$('nokKop').onclick=async()=>{const y=yerAdi();if(!y)return;const o=yerOzet(y);await kopya(o.nok.map(r=>r[1]).join('\n'));$('nokKop').textContent='Kopyalandı ('+o.nok.length+' sicil)';setTimeout(()=>$('nokKop').textContent='Noksan sicilleri kopyala',2000)};
document.addEventListener('click',async e=>{if(e.target.closest&&e.target.closest('button'))return;const d=e.target.closest&&e.target.closest('[data-cp]');if(d){await kopya(d.dataset.cp);Application_toast('Kopyalandı: '+d.dataset.cp)}});
let tt=null;function Application_toast(s){const m=$('toast');m.textContent=s;m.classList.remove('hide');clearTimeout(tt);tt=setTimeout(()=>m.classList.add('hide'),1600)}
// ---- Yerler / Kişiler (Excel: YerlerKişiler + KişiYeriSayım)
function yerlerSayfa(){const ds=hesaplaTum();const ad=new Map();for(const a of yerler)ad.set(norm(a),a);for(const [n,l] of yerSet){if(!ad.has(n)){const r=l.find(x=>norm(x[9])===n)||l[0];ad.set(n,norm(r[9])===n?r[9]:r[11])}}
 const st=new Map();for(const k of kayit){const n=norm(k.yer);const s=st.get(n)||st.set(n,{o:0,y:0,f:0,m:0}).get(n);s.o++;const d=ds.get(k.kimlik);if(d.t==='YERİNDE')s.y++;else if(d.t==='MÜKERRER')s.m++;else if(d.c==='faz'||d.c==='yer')s.f++}
 const q=norm($('yara').value);const satirlar=[...ad.entries()].filter(x=>!q||x[0].includes(q)).sort((a,b)=>a[1].localeCompare(b[1],'tr'));
 $('yT').innerHTML='<table class="t y"><tr><th>Yer / Kişi</th><th>Kay.</th><th>Oku.</th><th>Yer.</th><th>Nok.</th><th>Faz.</th><th>Mük.</th><th>Durum</th></tr>'+satirlar.map(x=>{const k=(yerSet.get(x[0])||[]).length,s=st.get(x[0])||{o:0,y:0,f:0,m:0};const sy=s.o>0;
  return '<tr data-yy="'+esc(x[1])+'"><td>'+esc(x[1])+'</td><td>'+k+'</td><td>'+s.o+'</td><td>'+s.y+'</td><td>'+Math.max(0,k-s.y)+'</td><td>'+s.f+'</td><td>'+s.m+'</td><td><span class="st" style="background:var(--'+(sy?'k-say':'k-faz')+');color:var(--'+(sy?'k-sayt':'k-fazt')+')">'+(sy?'SAYILDI':'SAYILMADI')+'</span></td></tr>'}).join('')+'</table>'}
const KYD={sis:['TÜMÜ','SAYILDI','SAYILMADI','BAŞKA YERDE'],fiz:['TÜMÜ','YERİNDE','FAZLA (hepsi)','FAZLA – sicilsiz (elle)','FAZLA – raporda yok','FAZLA – başka yerde kayıtlı','MÜKERRER','BARKOD (sicil değil)','YER SEÇİN']};
function kyDoldur(){const k=$('kyK').value,s=$('kyD').value;$('kyD').innerHTML=KYD[k].map(x=>'<option>'+x+'</option>').join('');if(KYD[k].includes(s))$('kyD').value=s}
function kyListe(){$('kyYT').classList.toggle('hide',!$('kyY').value);const y=yerK(),n=norm(y),tum=!y||n===TUMN,fiz=$('kyK').value==='fiz',df=$('kyD').value;const ds=hesaplaTum();let h='',say=0;
 $('yDB').textContent='Kişi / Yer listesi'+(y&&!tum?': '+y:' (tüm yerler)');
 const satir=(r,d,cls,x)=>{say++;return say>500?'':'<div class="it '+cls+'" data-cp="'+esc(r?r[1]:x.sicil)+'"><b>'+say+'. '+esc(r?r[1]:(x.sicil||x.malzeme))+'</b> '+esc(r?r[4]:'')+' <span class="st">'+esc(d)+'</span><small>'+(r?esc('Bar_Kod: '+r[2]+' · '+r[3]+(r[5]?' · '+r[5]:'')+' · '+(r[6]||'')+' '+(r[7]||'')+' · Zimmet: '+(r[11]||'')+' · Yer: '+(r[9]||'')):esc(x.yer+' · '+x.tarih))+'</small></div>'};
 if(!fiz){const kay=tum?bek:(yerKume(y)||[]);const tar=new Set(),gor=new Map();for(const k of kayit){const s=duzelt(k.sicil);if(!s)continue;tar.add(norm(k.yer)+'|'+s);if(!gor.has(s))gor.set(s,k.yer)}
  for(const r of kay){let d,c;if(tar.has(norm(r[9])+'|'+r[0])){d='SAYILDI';c='say'}else if(gor.has(r[0])){d='BAŞKA YERDE: '+gor.get(r[0]);c='kfaz'}else{d='SAYILMADI';c='plain'}
   if(df!=='TÜMÜ'&&!(df==='BAŞKA YERDE'?d.startsWith('BAŞKA'):d===df))continue;h+=satir(r,d,c)}}
 else{for(const k of kayit){if(!tum&&norm(k.yer)!==n)continue;const d=ds.get(k.kimlik);if(!d.t||d.t==='—')continue;if(df!=='TÜMÜ'&&!(df==='FAZLA (hepsi)'?d.t.startsWith('FAZLA'):d.t===df))continue;
   h+=satir(d.r,d.t,d.t==='YERİNDE'?'say':d.t==='MÜKERRER'?'kmuk':d.t.startsWith('FAZLA')?'kfaz':'yer',k)}}
 $('kyS').textContent='Kayıt: '+say+(say>500?' (ilk 500 gösteriliyor)':'');$('yD').innerHTML=h||'<div class="msg">Kayıt yok.</div>'}
function yerDetay(y){if(y)$('kyY').value=y;kyListe();$('yDK').scrollIntoView({behavior:'smooth'})}
$('kyK').onchange=()=>{kyDoldur();kyListe()};$('kyD').onchange=kyListe;kyDoldur();
$('yara').addEventListener('input',yerlerSayfa);
$('yT').addEventListener('click',e=>{const r=e.target.closest('tr[data-yy]');if(!r)return;const y=r.dataset.yy;$('kyY').value=y;yerDetay(y)});

// ---- Sayaç linkleri
let kfil={yer:'',durum:''};
function git(g){const y=yerAdi();if(!y)return;const sek=s=>document.querySelector('nav button[data-s='+s+']').click();
 if(g==='noksan'||g==='fazla'){sek('s-oz');setTimeout(()=>$(g==='noksan'?'cNok':'cFaz').scrollIntoView({behavior:'smooth'}),50)}
 else if(g==='kayitli'){sek('s-yer');yerDetay(y)}
 else{kfil={yer:y,durum:g==='yerinde'?'YERİNDE':g==='mukerrer'?'MÜKERRER':''};sek('s-kay')}}
document.addEventListener('click',e=>{const d=e.target.closest&&e.target.closest('.cnt div[data-g]');if(d)git(d.dataset.g)});
function urunSayfa(){$('uYT').classList.toggle('hide',!$('uY').value);const y=yerU(),fiz=$('uK').value==='fiz',n=norm(y),tum=!y||n===TUMN;let kay=[];
 $('urYer').textContent=y&&!tum?('Yer / kişi: '+y):'Yer seçilmedi: tümü';
 if(!fiz){const l=tum?bek:(yerKume(y)||[]);kay=l.map(r=>({b:r[2],a:r[3],c:r[4],e:r[5],m:r[6],o:r[7]}))}
 else{const ds=hesaplaTum();for(const k of kayit){if(!tum&&norm(k.yer)!==n)continue;const d=ds.get(k.kimlik);if(!d.t||d.t==='—'||d.t==='MÜKERRER'||d.t==='YER SEÇİN')continue;const r=d.r;kay.push(r?{b:r[2],a:r[3],c:r[4],e:r[5],m:r[6],o:r[7]}:{b:'',a:k.malzeme,c:'',e:'',m:'',o:''})}}
 const gb=new Map(),ga=new Map();for(const x of kay){const g=gb.get(x.b)||gb.set(x.b,{x,n:0}).get(x.b);g.n++;ga.set(x.a,(ga.get(x.a)||0)+1)}
 const la=[...ga.entries()].sort((a,b)=>b[1]-a[1]);
 $('urSt').textContent='İSTATİSTİK ►  Toplam: '+kay.length+' adet | Çeşit: '+gb.size+' | Tür: '+ga.size+(la.length?' | En çok: '+la[0][0]+' ('+la[0][1]+')':'');
 $('urT').innerHTML='<table class="t"><tr><th>Bar_Kod</th><th>Açıklama</th><th>Cinsi</th><th>Ek Özellik</th><th>Marka</th><th>Model</th><th>Adet</th></tr>'+[...gb.values()].sort((a,b)=>b.n-a.n).slice(0,500).map(g=>'<tr><td data-cp="'+esc(g.x.b)+'" style="cursor:pointer">'+esc(g.x.b)+'</td><td>'+esc(g.x.a)+'</td><td>'+esc(g.x.c)+'</td><td>'+esc(g.x.e)+'</td><td>'+esc(g.x.m)+'</td><td>'+esc(g.x.o)+'</td><td>'+g.n+'</td></tr>').join('')+'</table>';
 $('urA').innerHTML='<table class="t"><tr><th>Açıklama (toplam)</th><th>Adet</th></tr>'+la.map(x=>'<tr><td>'+esc(x[0])+'</td><td>'+x[1]+'</td></tr>').join('')+'</table>'}
$('uK').onchange=urunSayfa;$('notAc').onclick=()=>{notZorla=true;onizle();$('not').focus()};$('yOnce').onclick=()=>yerGit(-1);$('ySonra').onclick=()=>yerGit(1);$('ySay').onclick=yerSayilmamis;$('kalanGit').onclick=()=>document.querySelector('nav button[data-s=s-oz]').click();$('nokOkut').onclick=()=>{document.querySelector('nav button[data-s=s-say]').click();$('sicil').value='';$('sicil').focus()};$('uY').addEventListener('input',urunSayfa);$('uYT').onclick=()=>{$('uY').value=yerAdi();urunSayfa()};$('kyY').addEventListener('input',kyListe);$('kyYT').onclick=()=>{$('kyY').value=yerAdi();kyListe()};
function sorgu(){const s=duzelt($('sq').value);if(!s){$('sqK').innerHTML='';$('sqS').innerHTML='';return}
 const r=anah.get(s)||(barm.get(s)||[])[0];$('sqK').innerHTML=r?kartHtml(r):'<span class="msg">Sistemde kayıt bulunamadı.</span>';
 $('sqS').innerHTML=kayit.filter(k=>duzelt(k.sicil)===s).map(k=>'<div class="it plain" data-cp="'+esc(k.sicil)+'">Sayımda okutuldu: <b>'+esc(k.yer)+'</b><small>'+k.tarih+' '+k.saat+(k.not?' · '+esc(k.not):'')+'</small></div>').join('')}
// ---------- kayıt satırları ----------
function satir(k,d,sil){return '<div class="it '+(d.c==='ok'?'':d.c)+'"'+(k.sicil?' data-cp="'+esc(k.sicil)+'"':'')+'>'+(sil&&k.durum!=='gitti'?'<button class="x" data-sil="'+k.kimlik+'">×</button>':'')+(sil?'<button class="x" data-dz="'+k.kimlik+'" style="font-size:16px;margin-right:10px;color:var(--ac)">✎</button>':'')+'<b>'+esc(k.sicil||k.malzeme)+'</b> '+esc(d.t)+' <span class="msg">'+(k.durum==='gitti'?'✓':k.durum==='hata'?'⚠ hata':'⏳')+'</span><small>'+esc(k.yer)+' · '+k.tarih+' '+k.saat+(k.not?' · '+esc(k.not):'')+(d.r?' · '+esc(d.r[4]):'')+'</small></div>'}
function liste(){const ds=hesaplaTum();const sr=[...kayit].sort((a,b)=>a.kimlik<b.kimlik?1:-1);
 $('son').innerHTML=sr.slice(0,6).map(k=>satir(k,ds.get(k.kimlik),true)).join('')||'<div class="msg">Henüz kayıt yok.</div>';
 const qs=norm($('ara').value).split(' ').filter(Boolean);const fy=norm(kfil.yer),fd=kfil.durum;const t=sr.filter(k=>(!fy||fy===TUMN||norm(k.yer)===fy)&&(!fd||ds.get(k.kimlik).t===fd)&&(!qs.length||(()=>{const d=ds.get(k.kimlik),m=norm([k.sicil,k.malzeme,k.yer,k.not,k.tarih,k.saat,d.t,d.r?d.r.join(' '):''].join(' '));return qs.every(w=>m.includes(w))})()));
 $('kf').innerHTML=(kfil.yer||kfil.durum)?('Süzgeç: '+esc(kfil.yer)+' '+esc(kfil.durum)+' <a href="#" id="kfT">temizle</a>'):'';
 $('kn').textContent='('+kayit.length+')';$('tum').innerHTML=t.slice(0,400).map(k=>satir(k,ds.get(k.kimlik),true)).join('');
 const b=kayit.filter(k=>k.durum==='bekliyor'||k.durum==='hata').length;$('bek').textContent=b+' bekliyor';}
let sonYerS='';
function hepsi(){yPosGuncelle();if(yerAdi()!==sonYerS){sonYerS=yerAdi();$('sicil').value='';$('not').value='';$('ad').value='';notZorla=false;sonMS='';sonS='';$('uY').value=sonYerS;$('kyY').value=sonYerS;if(sonYerS){md.h='yer';medyaGoster('',sonYerS)}else $('medK').classList.add('hide')}$('yTem').classList.toggle('hide',!$('yer').value);onizle();ozetCiz();liste();if($('s-pano').classList.contains('on'))panoCiz();if($('s-oz').classList.contains('on'))ozetSayfa();if($('s-ur').classList.contains('on'))urunSayfa();if($('s-yer').classList.contains('on')){yerlerSayfa();kyListe()}}
// ---------- ekleme ----------
let sonS='',sonZ=0;
function sicilGoster(){const v=duzelt($('sicil').value);if(v&&v!==$('sicil').value)$('sicil').value=v}
async function ekle(){sicilGoster();const yer=yerAdi(),sicil=duzelt($('sicil').value),ad=temiz($('ad').value);
 if(!sicil&&!ad){$('sicil').focus();return}
 if(!yer){$('mesaj').className='msg er';$('mesaj').textContent='Önce sayım yerini / kişiyi seçin.';$('yer').focus();return}
 if(dzId){const k0=kayit.find(x=>x.kimlik===dzId);if(k0){const g0=new Set(kayit.filter(k=>k.kimlik!==dzId).map(k=>duzelt(k.sicil)).filter(Boolean));const d0=durum(sicil,ad,yer,g0);
  const xl0=k0.xl||{yer:k0.yer,sicil:k0.sicil,malzeme:k0.malzeme},gitti0=k0.durum==='gitti';
  k0.sicil=sicil;k0.malzeme=ad;k0.not=$('not').value||k0.not||'';k0.yer=yer;if(gitti0){k0.xl={yer,sicil,malzeme:ad};dzK.push({kimlik:k0.kimlik,ad:k0.kimlik+'_d'+Date.now()+'.txt',metin:dzMetin('duzelt',k0,xl0),deneme:0,ts:Date.now()});await dzKaydet();gLog('düzeltme',k0.kimlik,'kuyrukta',(xl0.sicil||xl0.malzeme)+' → '+(sicil||ad))}else{k0.durum='bekliyor';k0.deneme=0;k0.hata='';k0.ts=Date.now()}
  try{await dbYaz(k0)}catch(e){$('mesaj').className='msg er';$('mesaj').textContent='KAYIT SAKLANAMADI: '+e.message;return}
  dzId=null;notZorla=false;$('dzBan').classList.add('hide');$('sicil').value='';$('ad').value='';$('not').value='';$('mesaj').className='msg ok';$('mesaj').textContent='Düzeltildi: '+(sicil||ad)+' – '+d0.t;hepsi();$('sicil').focus();dzGonder(true);if(gecSn()>0)setTimeout(()=>gonder(false),(gecSn()+2)*1000);else gonder(false);return}}
 const t=Date.now();if(sicil&&sicil===sonS&&t-sonZ<1500){$('sicil').value='';return}sonS=sicil;sonZ=t;
 const gor=new Set(kayit.map(k=>duzelt(k.sicil)).filter(Boolean));const d=durum(sicil,ad,yer,gor);
 let nt=$('not').value;if(!nt){if(d.t==='FAZLA – başka yerde kayıtlı')nt='Başka odadan geldi';else if(d.t==='FAZLA – raporda yok')nt='Sistemde kaydı yok'}
 const n=new Date();
 const k={kimlik:damga(n)+'-'+temiz(LS.get('cihaz','tel')).replace(/[^\w-]/g,'')+'-'+rast(5),yer,tarih:tarihTR(n),saat:saatTR(n),sicil,malzeme:ad,not:nt,durum:'bekliyor',deneme:0,hata:'',medya:[],ts:Date.now()};
 try{await dbYaz(k)}catch(e){$('mesaj').className='msg er';$('mesaj').textContent='KAYIT SAKLANAMADI: '+e.message;return}
 kayit.push(k);LS.set('sonYer',yer);LS.set('yerGecmis',[yer,...LS.get('yerGecmis',[]).filter(x=>x!==yer)].slice(0,50));
 $('mesaj').className='msg '+(d.c==='ok'?'ok':'er');$('mesaj').textContent=(sicil||ad)+' – '+d.t;
 ses(d.c);$('sicil').value='';$('ad').value='';$('not').value='';notZorla=false;sonMS='';
 hepsi();$('ban').className='ban '+d.c;$('ban').innerHTML=esc(d.t);$('ban').classList.toggle('hide',!!d.r);$('kart').innerHTML=kartHtml(d.r,[['Kaydedildi',zamanPil(k.tarih+' '+k.saat,d),1]]);
 medyaGoster(sicil,yer);$('sicil').focus();gonder(false);if(gecSn()>0)setTimeout(()=>gonder(false),(gecSn()+2)*1000)}
function ses(c){try{if(navigator.vibrate)navigator.vibrate(c==='ok'?[30]:c==='muk'?[40,40,40,40,40]:[120,60,120]);
 const AC=window.AudioContext||window.webkitAudioContext;if(!AC)return;const x=ses.x||(ses.x=new AC());if(x.state==='suspended')x.resume();
 const P={ok:{f:[880]},muk:{f:[988,740,988,740,988],w:'square',d:.09,s:.115,g:.26},faz:{f:[660,440]},yer:{f:[400,300]}}[c];if(!P)return;
 let t=x.currentTime;const d=P.d||.16,s=P.s||.19,g0=P.g||.18;
 for(const h of P.f){const o=x.createOscillator(),g=x.createGain();o.type=P.w||'sine';o.frequency.value=h;g.gain.setValueAtTime(g0,t);g.gain.exponentialRampToValueAtTime(.001,t+d);o.connect(g);g.connect(x.destination);o.start(t);o.stop(t+d+.01);t+=s}}catch(e){}}
// ---------- yer önerisi ----------
function oneriKur(inp,box,ac,pick,kaynak){let akt=-1,lst=[];
 const tum=kaynak?(()=>kaynak().map(a=>[normAd(a),a])):(()=>{const kay=[[normAd(TUM),TUM]];const gor=new Set([TUMN]);for(const a of (yerler.length?yerler:LS.get('yerGecmis',[]))){const n=norm(a);if(!gor.has(n)){gor.add(n);kay.push([normAd(a),a])}}return kay});
 const kapat=()=>{box.classList.add('hide');akt=-1;lst=[]};
 const ciz=()=>{box.innerHTML=lst.map((a,i)=>{const s=kaynak?null:yerSet.get(norm(a));return '<div data-i="'+i+'"'+(i===akt?' class="on"':'')+'>'+esc(a)+(s?'<small>'+s.length+' kalem</small>':'')+'</div>'}).join('');box.classList.remove('hide');const o=box.querySelector('.on');if(o)o.scrollIntoView({block:'nearest'})};
 const goster=hepsi=>{const q=hepsi?'':normAd(inp.value);let e=tum().filter(x=>!q||x[0].includes(q)).map(x=>x[1]);if(!hepsi)e=e.slice(0,30);
  if(!e.length||(!hepsi&&q&&e.length===1&&norm(e[0])===norm(inp.value))){kapat();return}lst=e;akt=q?0:-1;ciz()};
 inp.addEventListener('input',()=>goster(false));
 inp.addEventListener('focus',()=>{if(!inp.value)goster(true)});
 inp.addEventListener('dblclick',()=>goster(true));
 ac.addEventListener('pointerdown',e=>{e.preventDefault();if(box.classList.contains('hide')){inp.focus();goster(true)}else kapat()});
 inp.addEventListener('blur',()=>setTimeout(kapat,200));
 inp.addEventListener('keydown',e=>{const acik=!box.classList.contains('hide');
  if(e.key==='ArrowDown'){e.preventDefault();if(!acik){goster(true);akt=0;ciz()}else{akt=Math.min(lst.length-1,akt+1);ciz()}}
  else if(e.key==='ArrowUp'&&acik){e.preventDefault();akt=Math.max(0,akt-1);ciz()}
  else if((e.key==='Enter'||e.key==='Tab')&&acik&&akt>=0){e.preventDefault();const a=lst[akt];kapat();pick(a)}
  else if(e.key==='Escape'&&acik){e.preventDefault();kapat()}});
 box.addEventListener('pointerdown',e=>{const d=e.target.closest('div[data-i]');if(d){e.preventDefault();const a=lst[+d.dataset.i];kapat();pick(a)}})}
function yerSec(a){$('yer').value=a;LS.set('sonYer',a);hepsi();$('sicil').focus()}
oneriKur($('yer'),$('oner'),$('yAc'),yerSec);
$('yer').addEventListener('input',hepsi);
const yanSec=id=>a=>{$(id).value=a;$(id).dispatchEvent(new Event('input',{bubbles:true}))};
oneriKur($('uY'),$('uOner'),$('uYA'),a=>{$('uY').value=a;urunSayfa();$('uY').blur()});
oneriKur($('kyY'),$('kyOner'),$('kyYA'),a=>{$('kyY').value=a;kyListe();$('kyY').blur()});
oneriKur($('not'),$('notO'),$('notA'),a=>{$('not').value=a;onizle()},()=>notlar);
oneriKur($('ad'),$('adO'),$('adA'),a=>{$('ad').value=a;onizle()},()=>malzemeler);
$('yer').addEventListener('change',()=>{LS.set('sonYer',yerAdi());hepsi()});
// ---------- medya (OneDrive: resim/video bul, ekle, değiştir) ----------
const RES_UZ=['jpg','jpeg','png','gif','bmp','webp','heic','tif','tiff','pdf'],VID_UZ=['mp4','mov','avi','wmv','mkv','m4v','3gp','mpg','mpeg','webm'],RES_YER=['jpg','jpeg','png','gif','bmp'],GOS=['jpg','jpeg','png','gif','bmp','webp'];
const resimK=()=>(LS.get('resimK','Sayım İşlemleri/Sayım Fotografları')||'').replace(/^\/+|\/+$/g,''),videoK=()=>(LS.get('videoK','Sayım İşlemleri/Sayım Videoları')||'').replace(/^\/+|\/+$/g,'');
const uz=a=>{const i=a.lastIndexOf('.');return i<0?'':a.slice(i+1).toLowerCase()},gsz=a=>{const i=a.lastIndexOf('.');return i<0?a:a.slice(0,i)};
const sayiBasi=a=>{const m=/^\d+/.exec(a);return m?m[0]:''};
const gUrl=y=>'https://graph.microsoft.com/v1.0/me/drive/root:/'+y.split('/').map(encodeURIComponent).join('/');
let md={h:'sicil',sicil:'',yer:'',bul:{r:null,v:null},sira:0};
async function tara(kl,uzlar,es,jt,d,ust){const sel='id,name,folder,file,webUrl,lastModifiedDateTime,@microsoft.graph.downloadUrl';let url=gUrl(kl)+':/children?$top=999&$select='+sel+',thumbnails&$expand=thumbnails';const ad=[];let yedek=false;
 while(url){let r=await fetch(url,{headers:{Authorization:'Bearer '+jt}});if(r.status===404)return null;
  if(!r.ok&&!yedek){yedek=true;url=gUrl(kl)+':/children?$top=999&$select='+sel;continue}
  if(!r.ok)throw new Error('HTTP '+r.status);const j=await r.json();
  for(const o of j.value){if(o.folder){if(d<3){const a=await tara(kl+'/'+o.name,uzlar,es,jt,d+1,o.name);if(a)ad.push(a)}continue}
   if(!uzlar.includes(uz(o.name)))continue;const p=es(o.name,ust);if(p>0){const t0=o.thumbnails&&o.thumbnails[0];const th=t0&&((t0.large||t0.medium||t0.small||{}).url)||'';
    ad.push({ad:o.name,url:o['@microsoft.graph.downloadUrl'],id:o.id,puan:p,klasor:kl,web:o.webUrl||'',th,tm:o.lastModifiedDateTime||''})}}
  url=j['@odata.nextLink']||null}
 ad.sort((a,b)=>b.puan-a.puan||(b.tm>a.tm?1:b.tm<a.tm?-1:0));return ad[0]||null}
$('medD').open=LS.get('medAc',false);$('medD').addEventListener('toggle',()=>{LS.set('medAc',$('medD').open);if($('medD').open&&md.sicil+md.yer)medyaAra()});
function medyaGoster(s,y){md.sicil=duzelt(s).replace(/[^0-9]/g,'');md.yer=normAd(y);
 if(!md.sicil&&!md.yer){$('medK').classList.add('hide');return}$('medK').classList.remove('hide');if(!md.sicil)md.h='yer';$('hSicil').disabled=!md.sicil;medyaHedef(md.h)}
function medyaHedef(h){md.h=h;$('hSicil').className='b '+(h==='sicil'?'':'g');$('hYer').className='b '+(h==='yer'?'':'g');medyaAra()}
async function medyaAra(){md.yer=normAd(yerAdi());const sr=++md.sira;$('medI').innerHTML='';$('medM').textContent='';md.bul={r:null,v:null};
 $('medS').textContent='Resim / Video — '+(md.h==='sicil'?'Malzeme '+md.sicil:'Yer: '+yerAdi());
 if(!$('medD').open){$('medM').textContent='';return}
 if(!hesap||!navigator.onLine){$('medM').textContent='Giriş yapılınca mevcut resim/video aranır. Yine de yeni ekleyebilirsiniz.';return}
 try{const jt=await belirtec();const s=md.h==='sicil';
  const es=s?(a=>sayiBasi(a)===md.sicil?2:0):((a,u)=>{if(!md.yer)return 0;if(u&&normAd(u)===md.yer)return 3;const x=normAd(gsz(a));return !x?0:(x===md.yer||x.startsWith(md.yer+'_')||x.startsWith(md.yer+' ')?2:0)});
  const[r,v]=await Promise.all([tara(resimK(),s?RES_UZ:RES_YER,es,jt,0),tara(videoK(),VID_UZ,es,jt,0)]);if(sr!==md.sira)return;md.bul={r,v};let h='';
  const gor=(x,vid)=>{if(!x)return'';const th=x.th,ln=x.web||'';
   if(vid&&x.url&&['mp4','mov','m4v','webm'].includes(uz(x.ad)))return '<video controls playsinline preload="metadata"'+(th?' poster="'+esc(th)+'"':'')+' src="'+esc(x.url)+'"></video><div class="msg">🎥 '+esc(x.ad)+'</div>';
   if(th)return (ln?'<a href="'+esc(ln)+'" target="_blank" rel="noopener">':'')+'<img src="'+esc(th)+'" alt="">'+(ln?'</a>':'')+'<div class="msg">'+(vid?'🎥 ':'📷 ')+esc(x.ad)+(ln?' · aç':'')+'</div>';
   if(!vid&&x.url&&GOS.includes(uz(x.ad)))return '<img src="'+esc(x.url)+'"><div class="msg">📷 '+esc(x.ad)+'</div>';
   return '<div class="msg">'+(vid?'Video var: ':'Resim var: ')+esc(x.ad)+(ln?' · <a href="'+esc(ln)+'" target="_blank" rel="noopener">aç</a>':'')+'</div>'};
  h=(r?'<div class="mi">'+gor(r,false)+'</div>':'')+(v?'<div class="mi">'+gor(v,true)+'</div>':'');
  $('medI').innerHTML=h;$('rBtn').textContent=r?'📷 Resmi değiştir':'📷 Fotoğraf ekle';$('vBtn').textContent=v?'🎥 Videoyu değiştir':'🎥 Video ekle';
  $('medM').textContent=(!r&&!v)?'Resim ve video yok; aşağıdan ekleyebilirsiniz.':''}catch(e){$('medM').textContent='Aranamadı: '+e.message}}
async function buluta(kl,ad,f,jt){if(f.size<=4*1024*1024){const r=await fetch(gUrl(kl+'/'+ad)+':/content?@microsoft.graph.conflictBehavior=rename',{method:'PUT',headers:{Authorization:'Bearer '+jt,'Content-Type':f.type||'application/octet-stream'},body:f});if(!r.ok)throw new Error('HTTP '+r.status);return}
 const o=await fetch(gUrl(kl+'/'+ad)+':/createUploadSession',{method:'POST',headers:{Authorization:'Bearer '+jt,'Content-Type':'application/json'},body:JSON.stringify({item:{'@microsoft.graph.conflictBehavior':'rename'}})});if(!o.ok)throw new Error('HTTP '+o.status);
 const up=(await o.json()).uploadUrl,P=5*320*1024;for(let b=0;b<f.size;b+=P){const s=Math.min(b+P,f.size);const r=await fetch(up,{method:'PUT',headers:{'Content-Range':'bytes '+b+'-'+(s-1)+'/'+f.size},body:f.slice(b,s)});if(!r.ok&&r.status!==202)throw new Error('HTTP '+r.status);$('medM').textContent='Yükleniyor %'+Math.round(s/f.size*100)}}
async function eskiyiTasi(e,jt){const kl=await fetch(gUrl(e.klasor)+':/children',{method:'POST',headers:{Authorization:'Bearer '+jt,'Content-Type':'application/json'},body:JSON.stringify({name:'Eski',folder:{},'@microsoft.graph.conflictBehavior':'replace'})}).catch(()=>null);
 let id=null;try{id=kl&&kl.ok?(await kl.json()).id:null}catch(_){}if(!id){const g=await fetch(gUrl(e.klasor+'/Eski'),{headers:{Authorization:'Bearer '+jt}});if(g.ok)id=(await g.json()).id}
 const b={name:'x_'+e.ad};if(id)b.parentReference={id};await fetch('https://graph.microsoft.com/v1.0/me/drive/items/'+e.id,{method:'PATCH',headers:{Authorization:'Bearer '+jt,'Content-Type':'application/json'},body:JSON.stringify(b)})}
const klasorAd=a=>String(a||'').replace(/[\"*:<>?\/\\|]/g,'-').replace(/\s+/g,' ').trim().replace(/[. ]+$/,'').slice(0,80)||'yer';
async function yerKlasorBul(kok,yer,jt){const hedef=norm(yer);let url=gUrl(kok)+':/children?$top=999&$select=name,folder';
 try{while(url){const r=await fetch(url,{headers:{Authorization:'Bearer '+jt}});if(!r.ok)break;const j=await r.json();for(const o of j.value)if(o.folder&&norm(o.name)===hedef)return o.name;url=j['@odata.nextLink']||null}}catch(e){}
 return klasorAd(yer)}
let medKy=[],medGon=false;
async function medKyYukle(){try{medKy=(await kvG('medKuyruk'))||[]}catch(e){medKy=[]}medKyGoster()}
const medKyKaydet=()=>kvS('medKuyruk',medKy);
function medKyGoster(){const n=medKy.length;$('medQ').textContent=n?('⏳ '+n+' resim/video gönderilmeyi bekliyor (internet ve giriş olunca kendiliğinden gider).'):'';$('medQb').classList.toggle('hide',!n)}
async function medKyGonder(elle){if(medGon||!medKy.length)return;if(!hesap||!navigator.onLine){if(elle)$('medM').textContent='İnternet ve Microsoft girişi gerekli.';return}
 medGon=true;let ok=0;
 try{const jt=await belirtec();
  for(const j of medKy.slice()){if(!elle&&(j.deneme||0)>=10)continue;
   try{let kl=j.kok;if(!j.sicilMi)kl=j.kok+'/'+await yerKlasorBul(j.kok,j.yer,jt);$('medM').textContent='Yükleniyor: '+j.ad+'…';await buluta(kl,j.ad,j.blob,jt);
    if(j.eski){try{await eskiyiTasi(j.eski,jt)}catch(_){}}
    medKy=medKy.filter(x=>x.id!==j.id);await medKyKaydet();gLog('medya',j.ad,'ok',kl+'/'+j.ad);ok++}
   catch(e){j.deneme=(j.deneme||0)+1;j.hata=e.message;await medKyKaydet();gLog('medya',j.ad,'hata',e.message);$('medM').textContent='Gönderilemedi, tekrar denenecek: '+e.message;if(/HTTP 40[13]/.test(e.message))break}}
 }catch(e){if(elle)$('medM').textContent='Gönderilemedi: '+e.message}
 medGon=false;medKyGoster();if(ok){$('medM').textContent=ok+' dosya gönderildi. Excel\'de Rapor Güncelle / Medya Bağla ile satıra bağlanır.';setTimeout(medyaAra,1500)}}
async function medyaSec(inp,tur){const f=inp.files&&inp.files[0];inp.value='';if(!f)return;const s=md.h==='sicil';
 if(s&&!md.sicil||!s&&!md.yer){$('medM').textContent='Önce sicil okutun / yer seçin.';return}
 let u=uz(f.name)||(tur===1?'jpg':'mp4');const izin=tur===1?(s?RES_UZ:RES_YER):VID_UZ;if(!izin.includes(u))u=tur===1?'jpg':'mp4';
 const n=new Date();const yc=(yerler.find(x=>norm(x)===norm(yerAdi()))||yerAdi()),ys=klasorAd(yc);const ad=(s?md.sicil:ys.slice(0,40)||'yer')+'_'+damga(n).replace('-','_')+'.'+u;
 const kok=tur===1?resimK():videoK(),eski=tur===1?md.bul.r:md.bul.v;
 medKy.push({id:damga(n)+'-'+rastgele(4),ad,tur,kok,sicilMi:s,yer:yc,blob:f,eski:eski?{id:eski.id,klasor:eski.klasor,ad:eski.ad}:null,deneme:0});
 await medKyKaydet();medKyGoster();
 if(!hesap||!navigator.onLine){$('medM').textContent='Kaydedildi; internet ve giriş olunca kendiliğinden gönderilecek.';return}
 $('medM').textContent='Gönderiliyor: '+ad+'…';await medKyGonder(true)}
$('medQb').onclick=()=>{medKy.forEach(j=>j.deneme=0);medKyGonder(true)};
$('hSicil').onclick=()=>medyaHedef('sicil');$('hYer').onclick=()=>medyaHedef('yer');
$('rBtn').onclick=()=>$('rGir').click();$('vBtn').onclick=()=>$('vGir').click();$('rGir').onchange=e=>medyaSec(e.target,1);$('vGir').onchange=e=>medyaSec(e.target,2);
// ---------- Microsoft / Graph ----------
const TEN='organizations',AUTHB='https://login.microsoftonline.com/'+TEN+'/oauth2/v2.0',SCP='Files.ReadWrite User.Read offline_access openid profile';let hesap=null,tok=LS.get('tok',null),girisHata=LS.get('girisHata','');
const cid=()=>(LS.get('cid','')||'').trim();
const redirUri=()=>location.origin+location.pathname.replace(/index\.html$/,'');
const b64u=b=>btoa(String.fromCharCode(...new Uint8Array(b))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
const rastgele=n=>b64u(crypto.getRandomValues(new Uint8Array(n)));
function hataKaydet(m){girisHata=String(m||'').slice(0,600);LS.set('girisHata',girisHata)}
function jwtKimlik(t){try{const p=JSON.parse(atob(t.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));return p.preferred_username||p.email||p.upn||p.name||''}catch(e){return ''}}
async function tokenAl(govde){const r=await fetch(AUTHB+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams(govde).toString()});const j=await r.json().catch(()=>({}));
 if(!r.ok||!j.access_token)throw new Error((j.error||('HTTP '+r.status))+': '+(j.error_description||'').split('\r')[0].slice(0,400));
 tok={at:j.access_token,rt:j.refresh_token||(tok&&tok.rt)||'',exp:Date.now()+(j.expires_in||3600)*1000,user:(j.id_token?jwtKimlik(j.id_token):'')||(tok&&tok.user)||''};LS.set('tok',tok);hesap={username:tok.user||'hesap'};return tok}
async function girisBaslat(){const ver=rastgele(48),st=rastgele(16);const ch=b64u(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(ver)));
 LS.set('pkce',{ver,st,t:Date.now()});
 location.href=AUTHB+'/authorize?'+new URLSearchParams({client_id:cid(),response_type:'code',redirect_uri:redirUri(),response_mode:'query',scope:SCP,state:st,code_challenge:ch,code_challenge_method:'S256',prompt:'select_account'}).toString()}
async function oturum(){
 try{const q=new URLSearchParams(location.search);
  if(q.get('error')){hataKaydet(q.get('error')+': '+(q.get('error_description')||'').replace(/\+/g,' ').split('\r')[0]);history.replaceState(null,'',location.pathname)}
  else if(q.get('code')){const pk=LS.get('pkce',null);history.replaceState(null,'',location.pathname);
   if(!pk||pk.st!==q.get('state'))throw new Error('Oturum durumu eşleşmedi, tekrar giriş yapın.');
   await tokenAl({client_id:cid(),grant_type:'authorization_code',code:q.get('code'),redirect_uri:redirUri(),code_verifier:pk.ver,scope:SCP});LS.set('pkce',null);hataKaydet('')}}
 catch(e){hataKaydet(e.message)}
 if(!hesap&&tok&&tok.user!==undefined)hesap={username:tok.user||'hesap'};
 bagGoster()}
function bagGoster(){$('girBan').classList.toggle('hide',!!hesap||!cid());$('bag').textContent=hesap?'bağlı':(cid()?'giriş yok':'girişsiz');$('am').textContent=hesap?'Giriş yapıldı: '+hesap.username:(cid()?'Giriş yapılmadı.'+(girisHata?' SON HATA → '+girisHata:''):'Client kimliği girilmedi: girişsiz mod (paylaş/dosya yolu).')}
async function belirtec(){if(!cid()||!tok)throw new Error('Giriş yapılmamış');
 if(tok.at&&tok.exp>Date.now()+60000)return tok.at;
 if(tok.rt){try{return(await tokenAl({client_id:cid(),grant_type:'refresh_token',refresh_token:tok.rt,scope:SCP})).at}catch(e){hataKaydet(e.message)}}
 hesap=null;bagGoster();throw new Error('Oturum süresi doldu: Ayar > Microsoft ile giriş')}
const mobYol=a=>(LS.get('kok','Sayım İşlemleri/Mobil')||'Sayım İşlemleri/Mobil').replace(/^\/+|\/+$/g,'')+'/'+a;
function cihazTur(){const u=navigator.userAgent||'',tp=navigator.maxTouchPoints||0;if(/iPad/.test(u)||(/Macintosh/.test(u)&&tp>1))return 'iPad';if(/iPhone|iPod/.test(u))return 'iPhone';if(/Android/.test(u))return /Mobile/.test(u)?'Android telefon':'Android tablet';if(/Windows/.test(u))return 'Windows PC';if(/Macintosh/.test(u))return 'Mac';if(/CrOS/.test(u))return 'Chromebook';return 'Diğer'}
function kvIcerik(k){return ['yer='+temiz(k.yer),'tarih='+k.tarih,'saat='+k.saat,'sicil='+temiz(k.sicil),'malzeme='+temiz(k.malzeme),'not='+temiz(k.not),'kimlik='+k.kimlik,'cihaz='+temiz(LS.get('cihaz','')),'tur='+cihazTur()].join('\n')+'\n'}
async function yukle(k,jt){const r=await fetch(gUrl(mobYol('Gelen/'+k.kimlik+'.txt'))+':/content?@microsoft.graph.conflictBehavior=fail',{method:'PUT',headers:{Authorization:'Bearer '+jt,'Content-Type':'text/plain; charset=utf-8'},body:new Blob([kvIcerik(k)],{type:'text/plain;charset=utf-8'})});
 if(r.status===409)return;if(!r.ok){let m='';try{m=(await r.json()).error.message}catch(e){}throw new Error('HTTP '+r.status+' '+m)}}
let gon=false;
async function gonder(elle){if(gon)return;if(!navigator.onLine){if(elle)$('gm').textContent='İnternet yok; kayıtlar saklı.';return}
 const gs=gecSn()*1000,bk=kayit.filter(k=>k.durum==='bekliyor'||k.durum==='hata');const b=kayit.filter(k=>k.durum==='bekliyor'&&(elle||!k.ts||Date.now()-k.ts>=gs)).sort((a,b)=>a.kimlik<b.kimlik?-1:1);if(!b.length){if(elle)$('gm').textContent='Gönderilecek kayıt yok.';return}
 if(!hesap){if(elle)$('gm').textContent='Giriş yok: "Dosya olarak paylaş" yolunu kullanın veya Ayarlar\'dan giriş yapın.';return}
 gon=true;let ok=0,h=0,jt;try{jt=await belirtec()}catch(e){gon=false;if(elle)$('gm').textContent='Gönderilemedi: '+e.message;return}
 for(const k of b){try{await yukle(k,jt);k.durum='gitti';k.hata='';k.xl={yer:k.yer,sicil:k.sicil,malzeme:k.malzeme};k.gTs=Date.now();await dbYaz(k);gLog('kayıt',k.kimlik,'ok','Gelen/'+k.kimlik+'.txt');ok++}catch(e){k.deneme=(k.deneme||0)+1;k.hata=e.message;k.durum=k.deneme>=5?'hata':'bekliyor';await dbYaz(k);gLog('kayıt',k.kimlik,'hata',e.message);h++;if(/HTTP 40[13]/.test(e.message))break}}
 gon=false;liste();if(ok){LS.set('sonGon',tarihTR(new Date())+' '+saatTR(new Date()));cihazBildir(true)}dzGonder(false);if(elle||h)$('gm').textContent=ok+' kayıt gönderildi'+(h?', '+h+' gönderilemedi (kuyrukta).':'. Excel açıkken 2 dk içinde alır.')}
$('gonder').onclick=()=>gonder(true);
$('tekrar').onclick=async()=>{for(const k of kayit)if(k.durum==='hata'){k.durum='bekliyor';k.deneme=0;await dbYaz(k)}gonder(true)};
function tsv(b){return '#SAYIM1\n'+b.map(k=>[k.yer,k.sicil,k.malzeme,k.not,isoTar(k),k.saat].map(temiz).join('\t')).join('\n')+'\n'}
$('paylas').onclick=async()=>{const b=kayit.filter(k=>k.durum!=='gitti');if(!b.length){$('gm').textContent='Gönderilecek kayıt yok.';return}
 const ad=temiz(LS.get('cihaz','tel')).replace(/[^\w-]/g,'')+'_'+damga(new Date())+'.txt';const f=new File([tsv(b)],ad,{type:'text/plain'});
 try{if(navigator.canShare&&navigator.canShare({files:[f]}))await navigator.share({files:[f],title:'Sayım kayıtları'});else{const a=document.createElement('a');a.href=URL.createObjectURL(f);a.download=ad;a.click()}
  if(confirm('Dosyayı kaydettiniz / gönderdiniz mi? Evet derseniz bu kayıtlar "gönderildi" işaretlenir.')){for(const k of b){k.durum='gitti';await dbYaz(k)}liste()}
  $('gm').textContent='Dosyayı OneDrive\\Sayım İşlemleri\\Mobil\\Gelen klasörüne koyun; Excel alır (Ayarlar > Mobil Kayıtları Al).'}catch(e){$('gm').textContent='İptal edildi.'}};
async function listeOto(elle){if(!hesap||!navigator.onLine)return false;try{const jt=await belirtec();const meta=LS.get('listeMeta',{});let y,n,b,m,deg=0;
 for(const[a,s]of[['yerler.txt',v=>y=v],['notlar.txt',v=>n=v],['malzemeler.txt',v=>m=v],['beklenen.txt',v=>b=v]]){
  const g=await fetch(gUrl(mobYol(a))+'?$select=lastModifiedDateTime,size',{headers:{Authorization:'Bearer '+jt}});if(!g.ok)continue;const j=await g.json();const imza=j.lastModifiedDateTime+'|'+j.size;
  if(!elle&&meta[a]===imza)continue;const r=await fetch(gUrl(mobYol(a))+':/content',{headers:{Authorization:'Bearer '+jt}});if(r.ok){s(await r.text());meta[a]=imza;deg++}}
 if(deg){if(y!==undefined)await kvS('yerler',y);if(n!==undefined)await kvS('notlar',n);if(m!==undefined)await kvS('malzemeler',m);if(b!==undefined)await kvS('beklenen',b);LS.set('listeMeta',meta);
  listeKur(y!=null?y:null,n!=null?n:null,b!=null?b:null,m!=null?m:null);hepsi();LS.set('listeZaman',Date.now())}
 $('lz').textContent='Son kontrol: '+new Date().toLocaleTimeString('tr-TR')+(deg?' · '+deg+' dosya yenilendi':' · değişiklik yok');return true}catch(e){$('lz').textContent='Liste güncelleme hatası: '+e.message;return false}}
setInterval(()=>listeOto(false),3*60*1000);document.addEventListener('visibilitychange',()=>{if(!document.hidden)listeOto(false)});
$('cek').onclick=async()=>{try{const jt=await belirtec();let y,n,b,m;for(const[a,s]of[['yerler.txt',v=>y=v],['notlar.txt',v=>n=v],['malzemeler.txt',v=>m=v],['beklenen.txt',v=>b=v]]){const r=await fetch(gUrl(mobYol(a))+':/content',{headers:{Authorization:'Bearer '+jt}});if(r.ok)s(await r.text())}
 await kvS('yerler',y);await kvS('notlar',n);await kvS('malzemeler',m);await kvS('beklenen',b);listeKur(y!=null?y:null,n!=null?n:null,b!=null?b:null,m!=null?m:null);hepsi();$('am').textContent='Listeler OneDrive\'dan alındı.'}catch(e){$('am').textContent='Alınamadı: '+e.message+' (Excel\'de Mobil Listeleri Yaz çalıştı mı?)'}};
$('dosya').onchange=async e=>{let y=null,n=null,b=null,m=null;for(const f of e.target.files){const t=await f.text(),a=f.name.toLowerCase();if(a.includes('bekle'))b=t;else if(a.includes('malzeme'))m=t;else if(a.includes('yer'))y=t;else if(a.includes('not'))n=t}
 if(y!=null)await kvS('yerler',y);if(n!=null)await kvS('notlar',n);if(m!=null)await kvS('malzemeler',m);if(b!=null)await kvS('beklenen',b);listeKur(y,n,b,m);hepsi();$('am').textContent='Listeler yüklendi.'};
$('test').onclick=async()=>{try{const jt=await belirtec();const r=await fetch(gUrl(mobYol('TEST-'+damga(new Date())+'.tmp'))+':/content',{method:'PUT',headers:{Authorization:'Bearer '+jt,'Content-Type':'text/plain'},body:'test'});if(!r.ok)throw new Error('HTTP '+r.status);const j=await r.json();await fetch('https://graph.microsoft.com/v1.0/me/drive/items/'+j.id,{method:'DELETE',headers:{Authorization:'Bearer '+jt}});$('am').textContent='Bağlantı çalışıyor (test dosyası yazıldı ve silindi).'}catch(e){$('am').textContent='Bağlantı kurulamadı: '+e.message}};
$('kaydet').onclick=()=>{LS.set('cihaz',$('cihaz').value.trim()||'tel');LS.set('cid',$('cid').value.trim());LS.set('oto',$('oto').checked);LS.set('kok',$('kok').value.trim()||'Sayım İşlemleri/Mobil');LS.set('resimK',$('rk').value.trim()||'Sayım İşlemleri/Sayım Fotografları');LS.set('videoK',$('vk').value.trim()||'Sayım İşlemleri/Sayım Videoları');bagGoster();$('am').textContent+=' Kaydedildi.'};
$('giris').onclick=async()=>{if(!cid()){$('am').textContent='Önce Client kimliğini girip Kaydet\'e basın.';return}if(!(window.crypto&&crypto.subtle)){$('am').textContent='Bu tarayıcı güvenli bağlantı (https) gerektiriyor.';return}hataKaydet('');$('am').textContent='Microsoft giriş sayfasına yönlendiriliyor…';try{await girisBaslat()}catch(e){hataKaydet(e.message);$('am').textContent='Giriş başlatılamadı → '+girisHata}};
$('cikis').onclick=()=>{tok=null;hesap=null;LS.set('tok',null);LS.set('pkce',null);hataKaydet('');bagGoster()};
$('sil').onclick=async()=>{if(!confirm('Gönderilmiş kayıtlar telefondan silinsin mi? (Okutma geçmişi / mükerrer kontrolü de bu kayıtlar için kalkar)'))return;for(const k of kayit.filter(k=>k.durum==='gitti'))await dbSil(k.kimlik);kayit=kayit.filter(k=>k.durum!=='gitti');hepsi()};
$('sifir').onclick=async()=>{const y=yerAdi();if(!y)return;if(!confirm(y+' için telefondaki TÜM kayıtlar silinsin mi? (OneDrive\'a gönderilmiş olanlar Excel\'de kalır)'))return;const n=norm(y);for(const k of kayit.filter(k=>norm(k.yer)===n))await dbSil(k.kimlik);kayit=kayit.filter(k=>norm(k.yer)!==n);hepsi()};
document.addEventListener('click',async e=>{const s=e.target.dataset&&e.target.dataset.sil;if(s&&confirm('Bu kayıt silinsin mi?')){await dbSil(s);kayit=kayit.filter(k=>k.kimlik!==s);hepsi()}});
// ---------- genel ----------
document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>{document.querySelectorAll('nav button,main>section').forEach(x=>x.classList.remove('on'));b.classList.add('on');$(b.dataset.s).classList.add('on');
 $('altbas').textContent={'s-say':'Fiziki Saha Sayım','s-oz':'Sayım Özeti / Kalan Ürün','s-yer':'Yerler / Kişiler · Kişi/Yer Listesi','s-ur':'Ürün Özeti · Malzeme Sorgu','s-kay':'Kayıtlar','s-pano':'Asistan · Pano','s-ay':'Ayarlar'}[b.dataset.s];
 if(b.dataset.s==='s-pano')panoCiz();if(b.dataset.s==='s-oz')ozetSayfa();if(b.dataset.s==='s-ur')urunSayfa();if(b.dataset.s==='s-yer'){yerlerSayfa();kyListe()}if(b.dataset.s==='s-say')$('sicil').focus();liste();window.scrollTo(0,0)});
$('ekle').onclick=ekle;$('temizle').onclick=()=>{for(const i of['sicil','ad','not'])$(i).value='';onizle();$('sicil').focus()};$('gd').onclick=()=>{$('sicil').value='';onizle();$('sicil').focus()};
$('yTem').onclick=()=>{$('yer').value='';LS.set('sonYer','');hepsi();$('yer').focus()};
document.addEventListener('click',e=>{if(e.target.id==='kfT'){e.preventDefault();kfil={yer:'',durum:''};liste()}});
$('sicil').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();ekle()}});
let oz=null,ot=null;$('sicil').addEventListener('input',()=>{clearTimeout(oz);oz=setTimeout(onizle,200);clearTimeout(ot);if(LS.get('oto',false)&&$('sicil').value.replace(/\s/g,'').length>=8)ot=setTimeout(()=>ekle(),700)});$('ad').addEventListener('input',onizle);
// ---------- kamerayla okuma ----------
let cam={akt:false,sur:true,stream:null,zx:null,son:'',sonT:0};
function camBulundu(v){const s=duzelt(v);if(!s)return;const t=Date.now();if(s===cam.son&&t-cam.sonT<2500)return;cam.son=s;cam.sonT=t;$('sicil').value=s;$('camH').textContent='Okundu: '+s;ekle().then(()=>{if(!cam.sur)camKapat()})}
let camZ=null,camIdle=null;
function zxHazir(){return new Promise((res,rej)=>{if(window.ZXing)return res();const e=document.createElement('script');e.src='zxing.min.js';e.onload=()=>window.ZXing?res():rej(new Error('Okuma kütüphanesi bozuk'));e.onerror=()=>rej(new Error('zxing.min.js bulunamadı'));document.head.appendChild(e)})}
function kamHazirla(){const goster=()=>$('kam').classList.remove('hide');if('BarcodeDetector'in window){goster();return}zxHazir().then(goster).catch(()=>{})}
function zxOkuyucu(){if(camZ)return camZ;const Z=ZXing,h=new Map();h.set(Z.DecodeHintType.TRY_HARDER,true);h.set(Z.DecodeHintType.POSSIBLE_FORMATS,[Z.BarcodeFormat.CODE_128,Z.BarcodeFormat.CODE_39,Z.BarcodeFormat.CODE_93,Z.BarcodeFormat.EAN_13,Z.BarcodeFormat.EAN_8,Z.BarcodeFormat.ITF,Z.BarcodeFormat.CODABAR,Z.BarcodeFormat.UPC_A,Z.BarcodeFormat.QR_CODE,Z.BarcodeFormat.DATA_MATRIX]);const r=new Z.MultiFormatReader();r.setHints(h);return camZ=r}
function cvDene(cx,w,h){const Z=ZXing,id=cx.getImageData(0,0,w,h).data,lum=new Uint8ClampedArray(w*h);for(let i=0,j=0;j<w*h;i+=4,j++)lum[j]=(id[i]*77+id[i+1]*150+id[i+2]*29)>>8;
 try{return zxOkuyucu().decode(new Z.BinaryBitmap(new Z.HybridBinarizer(new Z.RGBLuminanceSource(lum,w,h)))).getText()}catch(e){return null}}
async function camAc(){$('camK').classList.remove('hide');cam.akt=true;clearTimeout(camIdle);$('camH').textContent=cam.stream?'Kamera hazırlanıyor…':'Kamera açılıyor… (ilk seferde birkaç saniye sürebilir)';
 try{
  const bd='BarcodeDetector'in window;
  if(!bd)await zxHazir();
  // izin bir kez istenir: akım açık kalır, kapatınca sadece gizlenir
  if(!(cam.stream&&cam.stream.getVideoTracks().some(t=>t.readyState==='live')))cam.stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment',width:{ideal:1920},height:{ideal:1080}},audio:false});
  const v=$('camV');v.srcObject=cam.stream;await v.play();
  try{const tr=cam.stream.getVideoTracks()[0],cp=tr.getCapabilities?tr.getCapabilities():{};if(cp.focusMode&&cp.focusMode.includes('continuous'))await tr.applyConstraints({advanced:[{focusMode:'continuous'}]})}catch(e){}
  $('camH').textContent='Barkodu çerçeveye getirin (düz tutun, 15-25 cm). Okumazsa "Fotoğraf çek ve oku"ya basın.';
  if(bd){const det=new BarcodeDetector();const tur=async()=>{if(!cam.akt)return;try{const r=await det.detect(v);if(r.length)camBulundu(r[0].rawValue)}catch(e){}setTimeout(tur,250)};tur()}
  else{const cv=document.createElement('canvas'),cx=cv.getContext('2d',{willReadFrequently:true});const ac=[0,-4,4,-8,8,-12,12];let n=0;
   const tur=()=>{if(!cam.akt)return;try{const vw=v.videoWidth,vh=v.videoHeight;if(vw){
     const sc=Math.min(1,1280/vw),W=Math.round(vw*sc),H=Math.round(vh*sc),bantY=Math.round(H*.25),bantH=Math.round(H*.5);
     const a=ac[n++%ac.length]*Math.PI/180;cv.width=W;cv.height=bantH;cx.setTransform(1,0,0,1,0,0);cx.fillStyle='#fff';cx.fillRect(0,0,W,bantH);
     cx.translate(W/2,bantH/2);cx.rotate(a);cx.drawImage(v,0,vh*.25,vw,vh*.5,-W/2,-bantH/2,W,bantH);
     let t=cvDene(cx,W,bantH);if(!t&&n%ac.length===1){cv.width=W;cv.height=H;cx.setTransform(1,0,0,1,0,0);cx.drawImage(v,0,0,W,H);t=cvDene(cx,W,H)}
     if(t){if(t===cam.aday){cam.aday=null;camBulundu(t)}else cam.aday=t}}}catch(e){}setTimeout(tur,90)};tur()}
 }catch(e){$('camH').textContent='Kamera açılamadı: '+e.message;}}
function camKapat(){cam.akt=false;try{$('camV').pause()}catch(e){}$('camV').srcObject=null;$('camK').classList.add('hide');$('sicil').focus();
 // akım 10 dk boşta kalırsa bırakılır (pil); o zamana kadar izin tekrar sorulmaz
 clearTimeout(camIdle);camIdle=setTimeout(()=>{if(!cam.akt&&cam.stream){cam.stream.getTracks().forEach(t=>t.stop());cam.stream=null}},600000)}
// fotoğraftan okuma: telefonun kendi kamerası (net odak, yüksek çözünürlük), izin sorulmaz
async function fotoOku(file){$('camH').textContent='Fotoğraf okunuyor…';
 try{await zxHazir();const bmp=await createImageBitmap(file);const cv=document.createElement('canvas'),cx=cv.getContext('2d',{willReadFrequently:true});let t=null;
  for(const mw of [1600,1000,2400]){const sc=Math.min(1,mw/Math.max(bmp.width,bmp.height)),W=Math.round(bmp.width*sc),H=Math.round(bmp.height*sc);
   for(const ang of [0,-3,3,-6,6,-10,10,-15,15,90,-90]){const r=ang*Math.PI/180,q=Math.abs(ang)===90,cw=q?H:W,ch=q?W:H;cv.width=cw;cv.height=ch;cx.setTransform(1,0,0,1,0,0);cx.fillStyle='#fff';cx.fillRect(0,0,cw,ch);cx.translate(cw/2,ch/2);cx.rotate(r);cx.drawImage(bmp,-W/2,-H/2,W,H);
    t=cvDene(cx,cw,ch);if(t)break;await new Promise(r=>setTimeout(r,0))}
   if(t)break}
  if(t){camBulundu(t)}else $('camH').textContent='Fotoğrafta barkod okunamadı. Barkodu daha yakın/net çekin veya sicili elle yazın.'
 }catch(e){$('camH').textContent='Okunamadı: '+e.message}}
function camYetenek(){try{const tr=cam.stream&&cam.stream.getVideoTracks()[0];return tr&&tr.getCapabilities?{tr,cp:tr.getCapabilities()}:null}catch(e){return null}}
$('camT').onclick=async()=>{const y=camYetenek();if(!y||!y.cp.torch){$('camH').textContent='Bu telefon/tarayıcı feneri tarayıcıdan açmıyor.';return}cam.fener=!cam.fener;try{await y.tr.applyConstraints({advanced:[{torch:cam.fener}]});$('camT').textContent=cam.fener?'🔦 Fener: AÇIK':'🔦 Fener'}catch(e){cam.fener=false}};
$('camZ').onclick=async()=>{const y=camYetenek();if(!y||!y.cp.zoom){$('camH').textContent='Bu telefon yakınlaştırmayı tarayıcıdan desteklemiyor; biraz yaklaşın.';return}const z=y.cp.zoom,adim=[1,2,3].map(k=>Math.min(z.max,Math.max(z.min,k)));cam.zi=((cam.zi||0)+1)%adim.length;try{await y.tr.applyConstraints({advanced:[{zoom:adim[cam.zi]}]});$('camZ').textContent='🔍 Yakın: '+adim[cam.zi]+'x'}catch(e){}};
$('camF').onclick=()=>$('camI').click();$('camI').onchange=e=>{const f=e.target.files[0];e.target.value='';if(f)fotoOku(f)};


// ---------- asistan + pano ----------
function panoCiz(){rpOzCiz();const ds=hesaplaTum(),o=yerOzet(TUM);
 const hata=kayit.filter(k=>k.durum==='hata').length,bk=kayit.filter(k=>k.durum==='bekliyor').length,gitti=kayit.filter(k=>k.durum==='gitti').length;
 const ad=new Map();for(const a of yerListe())ad.set(norm(a),a);for(const [n,l] of yerSet){if(!ad.has(n)){const r=l.find(x=>norm(x[9])===n)||l[0];ad.set(n,norm(r[9])===n?r[9]:r[11])}}
 const st=new Map();for(const k of kayit){const n=norm(k.yer);const s=st.get(n)||st.set(n,{o:0,y:0,f:0,m:0}).get(n);s.o++;const d=ds.get(k.kimlik);if(d.t==='YERİNDE')s.y++;else if(d.t==='MÜKERRER')s.m++;else if(d.c==='faz'||d.c==='yer')s.f++}
 const sat=[...ad.entries()].map(x=>{const k=(yerSet.get(x[0])||[]).length,s=st.get(x[0])||{o:0,y:0,f:0,m:0};return{n:x[0],ad:x[1],k,s}}).filter(x=>x.k>0||x.s.o>0);
 const say=sat.filter(x=>x.s.o===0&&x.k>0);
 const g=[];const ekle=(c,t,m,f)=>g.push('<div class="it '+(c==='ok'?'':c)+'"'+(f?' data-pf="'+esc(f)+'" style="cursor:pointer"':'')+'><b>'+esc(t)+'</b><small>'+esc(m)+'</small></div>');
 if(hata)ekle('muk','Gönderilemeyen kayıt ('+hata+')','Dokunun: yeniden denensin.','tekrar');
 if(bk)ekle('faz','Bekleyen kayıt ('+bk+')','İnternet/giriş gelince kendiliğinden gider. Dokunun: şimdi gönder.','gonder');
 if(medKy.length)ekle('faz','Bekleyen resim/video ('+medKy.length+')','Dokunun: şimdi göndermeyi dene.','medya');
 if(!hesap&&cid())ekle('muk','Microsoft girişi yok','Ayar sekmesinden giriş yapın; yoksa hiçbir şey OneDrive\'a gitmez.','ayar');
 if(!bek.length)ekle('faz','Sistem listesi yok','Excel\'de Mobil Listeleri Yaz, sonra Ayar sekmesinden listeleri alın.','ayar');
 if(o.muk)ekle('muk','Mükerrer ('+o.muk+')','Aynı sicil birden fazla okutulmuş. Özet sekmesinde bakın.','oz');
 if(o.faz)ekle('faz','Fazla ('+o.faz+')','Raporda yok ya da başka yerde kayıtlı.','oz');
 if(say.length)ekle('plain','Sayılmamış yer ('+say.length+')','Sıradaki: '+say[0].ad+(say.length>1?'  ·  '+say.slice(1,4).map(x=>x.ad).join('  ·  '):'')+'. Dokunun: ilkine geç.','say:'+say[0].ad);
 else if(sat.length)ekle('ok','Tüm yerlerde okutma var','Listedeki her yerde en az bir okutma yapılmış.');
 if(!g.length)ekle('ok','Her şey yolunda','Yapılacak bir şey görünmüyor.');
 $('asL').innerHTML=g.join('');
 $('pnK').innerHTML=cntHtml(o);const yz=o.kay?Math.round(100*o.yerinde/o.kay):0;$('pnB').style.width=yz+'%';$('pnM').textContent='%'+yz+' tamamlandı ('+o.yerinde+' / '+o.kay+') · '+sat.length+' yer/kişi · '+(sat.length-say.length)+' tanesinde okutma var';
 sat.sort((a,b)=>(a.s.o?1:0)-(b.s.o?1:0)||a.ad.localeCompare(b.ad,'tr'));
 $('pnY').innerHTML='<table class="t pn"><tr><th>Yer / Kişi</th><th></th><th>Yer./Kay.</th></tr>'+sat.map(x=>{const p=x.k?Math.min(100,Math.round(100*x.s.y/x.k)):0;return '<tr data-yy="'+esc(x.ad)+'" style="cursor:pointer"><td>'+esc(x.ad)+'</td><td><div class="pb'+(x.s.o?'':' z')+'"><i style="width:'+p+'%"></i></div></td><td>'+x.s.y+' / '+x.k+(x.s.f?' · F'+x.s.f:'')+(x.s.m?' · M'+x.s.m:'')+'</td></tr>'}).join('')+'</table>';
 const bug=tarihTR(new Date()),bn=kayit.filter(k=>k.tarih===bug).length,son=kayit.length?kayit[kayit.length-1]:null;
 $('pnC').innerHTML='<div class="kv"><b>Kullanıcı / cihaz</b><span>'+esc(LS.get('cihaz','')||'(adsız)')+' · '+cihazTur()+'</span><b>Toplam kayıt</b><span>'+kayit.length+' (gönderilen '+gitti+', bekleyen '+bk+', hatalı '+hata+')</span><b>Bugün</b><span>'+bn+' kayıt</span><b>Son kayıt</b><span>'+(son?esc(son.tarih+' '+son.saat):'—')+'</span><b>Bağlantı</b><span>'+(hesap?'giriş yapıldı':'girişsiz')+'</span></div>'}
document.addEventListener('click',e=>{const a=e.target.closest('[data-pf]');if(a){const f=a.dataset.pf;if(f==='tekrar')$('tekrar').click();else if(f==='gonder')gonder(true);else if(f==='medya')medKyGonder(true);else if(f==='ayar')document.querySelector('nav button[data-s=s-ay]').click();else if(f==='oz')document.querySelector('nav button[data-s=s-oz]').click();else if(f.startsWith('say:')){yerAt(f.slice(4));document.querySelector('nav button[data-s=s-say]').click()}setTimeout(panoCiz,300);return}
 const t=e.target.closest('#pnY tr[data-yy]');if(t){yerAt(t.dataset.yy);document.querySelector('nav button[data-s=s-say]').click()}});


// ---------- gönderim günlüğü, düzeltme kuyruğu, cihaz bildirimi, rapor ----------
let glog=[],dzK=[],dzGon=false,sonBildir=0,rpSon=null;
const dzKaydet=()=>kvS('dzKuyruk',dzK);
function gLog(tur,id,sonuc,det){glog.unshift({t:Date.now(),k:tur,id:id||'',s:sonuc,d:det||''});if(glog.length>500)glog.length=500;kvS('glog',glog)}
async function txtYukle(yol,metin,jt,cb){cb=cb||'fail';const r=await fetch(gUrl(yol)+':/content?@microsoft.graph.conflictBehavior='+cb,{method:'PUT',headers:{Authorization:'Bearer '+jt,'Content-Type':'text/plain; charset=utf-8'},body:new Blob([metin],{type:'text/plain;charset=utf-8'})});
 if(r.status===409&&cb==='fail')return;if(!r.ok){let m='';try{m=(await r.json()).error.message}catch(e){}throw new Error('HTTP '+r.status+' '+m)}}
function dzMetin(islem,k,eski){return ['islem='+islem,'kimlik='+k.kimlik,'eski_yer='+temiz(eski.yer),'eski_sicil='+temiz(eski.sicil),'eski_malzeme='+temiz(eski.malzeme),'tarih='+k.tarih,'saat='+k.saat,'yer='+temiz(k.yer),'sicil='+temiz(k.sicil),'malzeme='+temiz(k.malzeme),'not='+temiz(k.not),'cihaz='+temiz(LS.get('cihaz','')),'tur='+cihazTur()].join('\n')+'\n'}
async function dzGonder(elle){if(dzGon||!dzK.length||!hesap||!navigator.onLine)return;dzGon=true;
 try{const jt=await belirtec();for(const j of dzK.slice()){try{await txtYukle(mobYol('Gelen/'+j.ad),j.metin,jt);dzK=dzK.filter(x=>x!==j);await dzKaydet();gLog('düzeltme',j.kimlik,'ok','Gelen/'+j.ad)}catch(e){j.deneme=(j.deneme||0)+1;j.hata=e.message;await dzKaydet();gLog('düzeltme',j.kimlik,'hata',e.message);if(/HTTP 40[13]/.test(e.message))break}}}catch(e){}
 dzGon=false;liste()}
async function kayitIptal(){const k=kayit.find(x=>x.kimlik===dzId);if(!k)return;
 if(!confirm('Bu kayıt iptal edilsin mi?'+(k.durum==='gitti'?' Excel\'deki satır da temizlenir.':''))) return;
 if(k.durum==='gitti'){const xl=k.xl||{yer:k.yer,sicil:k.sicil,malzeme:k.malzeme};dzK.push({kimlik:k.kimlik,ad:k.kimlik+'_i'+Date.now()+'.txt',metin:dzMetin('iptal',k,xl),deneme:0,ts:Date.now()});await dzKaydet();gLog('iptal',k.kimlik,'kuyrukta',k.sicil||k.malzeme)}else gLog('iptal',k.kimlik,'ok','gönderilmemiş kayıt silindi');
 await dbSil(k.kimlik);kayit=kayit.filter(x=>x.kimlik!==k.kimlik);duzeltBitir();dzGonder(true)}
async function cihazBildir(zorla){if(!hesap||!navigator.onLine)return;if(!zorla&&Date.now()-sonBildir<120000)return;sonBildir=Date.now();
 try{const jt=await belirtec();const ad=(temiz(LS.get('cihaz',''))||'adsiz').replace(/[^\w-]/g,'')||'adsiz',tr=cihazTur(),n=new Date();
  const c=d=>kayit.filter(k=>k.durum===d).length;
  const m=['cihaz='+temiz(LS.get('cihaz','')||'(adsız)'),'tur='+tr,'surum=v10','zaman='+tarihTR(n)+' '+saatTR(n),'hesap='+temiz(hesap.username||''),'yer='+temiz(yerAdi()),'toplam='+kayit.length,'gitti='+c('gitti'),'bekleyen='+c('bekliyor'),'hata='+c('hata'),'medya='+medKy.length,'duzeltme='+dzK.length,'son_gonderim='+temiz(LS.get('sonGon',''))].join('\n')+'\n';
  await txtYukle(mobYol('Cihazlar/'+ad+'_'+tr.replace(/\W/g,'')+'.txt'),m,jt,'replace')}catch(e){}}
async function listeKlasor(yol,jt){let u=gUrl(yol)+':/children?$select=name,size,lastModifiedDateTime&$top=999',o=[];while(u){const r=await fetch(u,{headers:{Authorization:'Bearer '+jt}});if(r.status===404)return[];if(!r.ok)throw new Error('HTTP '+r.status);const j=await r.json();o.push(...j.value);u=j['@odata.nextLink']||null}return o}
const yasS=ms=>{const d=Math.max(0,Math.round(ms/60000));return d<1?'şimdi':d<60?d+' dk önce':d<1440?Math.round(d/60)+' sa önce':Math.round(d/1440)+' gün önce'};
const zS=t=>{const n=new Date(t);return tarihTR(n)+' '+saatTR(n)};
function rpOzCiz(){const c=d=>kayit.filter(k=>k.durum===d).length;
 $('rpOz').innerHTML='<div class="kv"><b>Telefonda kayıt</b><span>'+kayit.length+' (gönderilen '+c('gitti')+', bekleyen '+c('bekliyor')+', hatalı '+c('hata')+')</span><b>Bekleyen resim/video</b><span>'+medKy.length+'</span><b>Bekleyen düzeltme/iptal</b><span>'+dzK.length+'</span><b>Son gönderim</b><span>'+esc(LS.get('sonGon','—'))+'</span><b>Giriş</b><span>'+(hesap?esc(hesap.username):'yok (girişsiz)')+'</span></div>';
 $('glogL').innerHTML=glog.slice(0,100).map(g=>'<div class="it '+(g.s==='hata'?'muk':g.s==='kuyrukta'?'faz':'')+'"><b>'+esc(g.k)+'</b> <span class="msg">'+esc(g.s)+' · '+zS(g.t)+'</span><small>'+esc(g.id)+(g.d?' → '+esc(g.d):'')+'</small></div>').join('')||'<div class="msg">Henüz günlük yok.</div>'}
async function rpSor(){if(!hesap||!navigator.onLine){$('rpM').textContent='Giriş ve internet gerekli.';return}$('rpM').textContent='OneDrive klasörleri okunuyor…';
 try{const jt=await belirtec();const [ge,isl,ha]=await Promise.all([listeKlasor(mobYol('Gelen'),jt),listeKlasor(mobYol('Islendi'),jt),listeKlasor(mobYol('Hata'),jt)]);
  const gi=new Map(ge.map(f=>[f.name,f])),sf=(L,id)=>L.find(f=>f.name===id+'.txt'||f.name.endsWith('_'+id+'.txt'));
  const sat={islendi:[],gelen:[],hata:[],yok:[],bek:[]};
  for(const k of kayit){if(k.durum!=='gitti'){sat.bek.push([k,k.durum==='hata'?'Telefonda gönderilemedi: '+(k.hata||''):'Telefonda sırada']);continue}
   const f1=sf(isl,k.kimlik),f2=sf(ge,k.kimlik),f3=sf(ha,k.kimlik);
   if(f3)sat.hata.push([k,'Hata klasöründe (Excel okuyamadı)']);else if(f1)sat.islendi.push([k,'Excel işledi']);else if(f2)sat.gelen.push([k,'Gelen klasöründe bekliyor, '+yasS(Date.now()-new Date(f2.lastModifiedDateTime))]);else sat.yok.push([k,'Hiçbir klasörde bulunamadı'])}
  const eskiG=ge.length?Math.max(...ge.map(f=>Date.now()-new Date(f.lastModifiedDateTime))):0;
  rpSon={zaman:Date.now(),gelen:ge.length,islendi:isl.length,hata:ha.length,sat,eskiG};
  const kt=(a,c)=>'<div class="it '+c+'"><b>'+a.length+'</b> '+({islendi:'Excel işledi',gelen:'Gelen kutusunda bekliyor',hata:'Hata klasöründe',yok:'Bulunamadı',bek:'Telefonda bekliyor/hatalı'})[a.k]+'</div>';
  const blok=(ad,L,c)=>L.length?'<div class="msg" style="margin-top:8px"><b>'+ad+'</b></div>'+L.slice(0,40).map(([k,m])=>'<div class="it '+c+'"><b>'+esc(k.sicil||k.malzeme)+'</b> <span class="msg">'+esc(k.yer)+' · '+k.tarih+' '+k.saat+'</span><small>'+esc(m)+'</small></div>').join(''):'';
  $('rpL').innerHTML='<div class="kv"><b>Gelen klasörü (herkesten)</b><span>'+ge.length+' dosya'+(ge.length?' · en eskisi '+yasS(eskiG):'')+'</span><b>İşlendi klasörü</b><span>'+isl.length+' dosya</span><b>Hata klasörü</b><span>'+ha.length+' dosya</span><b>Bu telefonun kayıtları</b><span>işlendi '+sat.islendi.length+' · gelende '+sat.gelen.length+' · hata '+sat.hata.length+' · bulunamadı '+sat.yok.length+' · telefonda '+sat.bek.length+'</span><b>Sorgu zamanı</b><span>'+zS(Date.now())+'</span></div>'+blok('Bulunamayanlar',sat.yok,'muk')+blok('Hata klasöründekiler',sat.hata,'muk')+blok('Gelen kutusunda bekleyenler (Excel açık değil / 2 dk içinde alır)',sat.gelen,'faz')+blok('Telefonda bekleyen / hatalı',sat.bek,'faz');
  $('rpM').textContent='Tamam. '+(ge.length&&eskiG>600000?'Uyarı: Gelen kutusundaki dosyalar 10 dakikadan eski; Excel açık mı?':'')}catch(e){$('rpM').textContent='Sorgulanamadı: '+e.message}}
async function cihSor(){if(!hesap||!navigator.onLine){$('cihL').innerHTML='<div class="msg">Giriş ve internet gerekli.</div>';return}$('cihL').innerHTML='<div class="msg">Okunuyor…</div>';
 try{const jt=await belirtec();const L=(await listeKlasor(mobYol('Cihazlar'),jt)).filter(f=>/\.txt$/i.test(f.name));const o=[];
  for(const f of L){try{const r=await fetch(gUrl(mobYol('Cihazlar/'+f.name))+':/content',{headers:{Authorization:'Bearer '+jt}});if(!r.ok)continue;const v={};for(const ln of(await r.text()).split(/\r?\n/)){const i=ln.indexOf('=');if(i>0)v[ln.slice(0,i)]=ln.slice(i+1)}v._yas=Date.now()-new Date(f.lastModifiedDateTime);v._d=f.lastModifiedDateTime;o.push(v)}catch(e){}}
  o.sort((a,b)=>a._yas-b._yas);const bagli=o.filter(v=>v._yas<15*60000).length;
  $('cihL').innerHTML='<div class="msg"><b>'+o.length+'</b> cihaz kayıtlı, <b>'+bagli+'</b> tanesi son 15 dakikada bağlandı.</div>'+(o.map(v=>'<div class="it '+(v._yas<15*60000?'':v._yas<86400000?'faz':'no')+'"><b>'+esc(v.cihaz||'?')+'</b> · '+esc(v.tur||'')+' <span class="msg">'+(v._yas<15*60000?'● BAĞLI':'son görülme '+yasS(v._yas))+'</span><small>Hesap: '+esc(v.hesap||'—')+' · sürüm '+esc(v.surum||'?')+' · toplam '+esc(v.toplam||'0')+' · gönderilen '+esc(v.gitti||'0')+' · bekleyen '+esc(v.bekleyen||'0')+' · hatalı '+esc(v.hata||'0')+' · medya '+esc(v.medya||'0')+' · son yer: '+esc(v.yer||'—')+'</small></div>').join('')||'<div class="msg">Henüz cihaz bildirimi yok. Telefon giriş yapıp kayıt gönderince görünür.</div>')}catch(e){$('cihL').innerHTML='<div class="msg er">Okunamadı: '+esc(e.message)+'</div>'}}
$('rpSor').onclick=rpSor;$('cihSor').onclick=cihSor;
$('rpKop').onclick=async()=>{const sat=rpSon&&rpSon.sat;let t='SAYIM GÖNDERİM RAPORU '+zS(Date.now())+'\nCihaz: '+(LS.get('cihaz','')||'(adsız)')+' / '+cihazTur()+'\nKayıt: '+kayit.length+' · medya bekleyen '+medKy.length+' · düzeltme bekleyen '+dzK.length+'\n';
 if(rpSon)t+='Gelen '+rpSon.gelen+' · İşlendi '+rpSon.islendi+' · Hata '+rpSon.hata+'\nBu telefon: işlendi '+sat.islendi.length+', gelende '+sat.gelen.length+', hata '+sat.hata.length+', bulunamadı '+sat.yok.length+', telefonda '+sat.bek.length+'\n';
 t+='\nGÜNLÜK\n'+glog.slice(0,200).map(g=>zS(g.t)+'\t'+g.k+'\t'+g.s+'\t'+g.id+'\t'+g.d).join('\n');
 try{await navigator.clipboard.writeText(t);$('rpM').textContent='Rapor kopyalandı.'}catch(e){$('rpM').textContent='Kopyalanamadı; ekran görüntüsü alın.'}};

// ---------- yapıştırma: pano -> Excel gibi temiz sicil (nokta, / ve boşluk gider) ----------
function panoTemiz(t){const l=String(t||'').split(/\r?\n/).map(x=>x.trim()).find(x=>x)||'';return duzelt(l.split('\t')[0])}
function yapistir(id,sonra){const el=$(id);const ver=t=>{const v=panoTemiz(t);if(!v){$('mesaj').className='msg er';$('mesaj').textContent='Panoda sicil bulunamadı.';return}el.value=v;el.dispatchEvent(new Event('input',{bubbles:true}));if(sonra)sonra();Application_toast('Yapıştırıldı: '+v)};
 if(navigator.clipboard&&navigator.clipboard.readText)navigator.clipboard.readText().then(ver).catch(()=>{el.focus();el.select();$('mesaj').className='msg';$('mesaj').textContent='Kutuya uzun basıp "Yapıştır" deyin.'});
 else{el.focus();$('mesaj').textContent='Kutuya uzun basıp "Yapıştır" deyin.'}}
for(const id of ['sicil','sq'])$(id).addEventListener('paste',e=>{const t=(e.clipboardData||window.clipboardData).getData('text');e.preventDefault();const v=panoTemiz(t);const el=$(id);el.value=v;el.dispatchEvent(new Event('input',{bubbles:true}))});
$('yap').onclick=()=>yapistir('sicil');$('yapQ').onclick=()=>yapistir('sq');
setTimeout(kamHazirla,800);
$('kam').onclick=()=>{if(!yerAdi()){$('mesaj').className='msg er';$('mesaj').textContent='Önce sayım yerini / kişiyi seçin.';$('yer').focus();return}camAc()};
$('camX').onclick=camKapat;$('camS').onclick=()=>{cam.sur=!cam.sur;$('camS').textContent='Sürekli: '+(cam.sur?'AÇIK':'KAPALI')};
$('sq').addEventListener('input',sorgu);$('ara').addEventListener('input',liste);
window.addEventListener('online',()=>{gonder(false);medKyGonder(false);dzGonder(false);cihazBildir(true)});setInterval(()=>{gonder(false);medKyGonder(false);dzGonder(false);cihazBildir(false)},60000);setTimeout(()=>cihazBildir(true),4000);$('girBanB').onclick=()=>$('giris').click();
(async()=>{$('ru').textContent=location.origin+location.pathname.replace(/index\.html$/,'');
 $('oto').checked=!!LS.get('oto',false);$('cihaz').value=LS.get('cihaz','');$('cid').value=LS.get('cid','');$('kok').value=LS.get('kok','Sayım İşlemleri/Mobil');$('rk').value=LS.get('resimK','Sayım İşlemleri/Sayım Fotografları');$('vk').value=LS.get('videoK','Sayım İşlemleri/Sayım Videoları');
 try{await dbAc()}catch(e){$('mesaj').className='msg er';$('mesaj').textContent='Yerel depolama açılamadı: '+e.message;return}
 if(navigator.storage&&navigator.storage.persist)navigator.storage.persist();
 kayit=(await dbHepsi()).sort((a,b)=>a.kimlik<b.kimlik?-1:1);await medKyYukle();try{dzK=(await kvG('dzKuyruk'))||[];glog=(await kvG('glog'))||[]}catch(e){}
 listeKur(await kvG('yerler')??null,await kvG('notlar')??null,await kvG('beklenen')??null,await kvG('malzemeler')??null);
 $('yer').value='';LS.set('sonYer','');await oturum();hepsi();
 if('serviceWorker'in navigator)navigator.serviceWorker.register('sw.js').catch(()=>{});
 listeOto(false);
 if(!LS.get('cihaz','')){document.querySelector('nav button[data-s=s-ay]').click();$('am').textContent='Önce cihaz adını yazıp Kaydet\'e basın.'}else $('sicil').focus();gonder(false);medKyGonder(false)})();
